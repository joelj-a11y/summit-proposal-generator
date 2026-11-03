#!/usr/bin/env python3
"""
Summit Quote & Proposal Generator
=================================

One web app that serves:
  * a quote calculator page for each team (see TEAMS below)
  * the proposal generator page (/generate), which turns a downloaded quote
    file (.json) into a Word proposal using Summit_Proposal_Template.docx
  * a quote log (/webhook), which adds one row per quote to a Google Sheet,
    using a separate tab for each team

Where things live
-----------------
  app.py                         this file (web server + Word + Google Sheets)
  summit-quote-generator.html    the calculator page, shared by every team
  teams/<team>.js                each team's pricing assumptions (edit these!)
  Summit_Proposal_Template.docx  the Word template, filled in using {{PLACEHOLDERS}}

Render settings
---------------
  Build command:  pip install -r requirements.txt
  Start command:  gunicorn app:app
  Environment:    GOOGLE_SHEET_ID, GOOGLE_SHEETS_CREDENTIALS
"""

import copy
import json
import os
import re
import tempfile
import traceback
from datetime import datetime

import pytz
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from flask import Flask, abort, jsonify, render_template_string, request, send_file, send_from_directory

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_FILE = os.path.join(BASE_DIR, 'Summit_Proposal_Template.docx')
CALCULATOR_FILE = os.path.join(BASE_DIR, 'summit-quote-generator.html')
TEAMS_DIR = os.path.join(BASE_DIR, 'teams')
MOUNTAIN_TZ = pytz.timezone('America/Denver')


# =============================================================================
# TEAMS
# -----------------------------------------------------------------------------
# One entry per team. To add a team:
#   1. Copy teams/summit.js to teams/<new-key>.js and edit the numbers in it
#   2. Add an entry below with the same key
# The Google Sheet tab is created automatically the first time a quote is logged.
# =============================================================================
TEAMS = {
    'summit': {
        'name': 'Summit Services Group',
        'path': '/',                       # web address of the calculator
        'settings_file': 'summit.js',      # file inside the teams/ folder
        'sheet_tab': 'Summit',             # Google Sheet tab for the quote log
    },
    'western-colorado': {
        'name': 'SSG Western Colorado',
        'path': '/western-colorado',
        'settings_file': 'western-colorado.js',
        'sheet_tab': 'SSG Western Colorado',
    },
}
DEFAULT_TEAM = 'summit'


def get_team(team_key):
    """Look up a team, falling back to the default team for old quote files."""
    return TEAMS.get(team_key or DEFAULT_TEAM)


# =============================================================================
# NUMBER HELPERS
# =============================================================================
def to_number(value, default=0.0):
    """Turn '1,250', '1250', 1250 or None into a float."""
    if value is None or value == '':
        return default
    try:
        return float(str(value).replace(',', '').replace('$', ''))
    except (TypeError, ValueError):
        return default


def money(value):
    return f"${value:,.0f}"


def quantity(value):
    """24.0 -> '24', 7.5 -> '7.5'"""
    value = to_number(value)
    return str(int(value)) if float(value).is_integer() else f"{value:.1f}"


# =============================================================================
# PROPOSAL NUMBERS
# -----------------------------------------------------------------------------
# All proposal math lives here, so the Word document and the quote log always
# show the same numbers. Assumptions come from the quote file (which gets them
# from the team's settings file); the values after "or" are fallbacks for quote
# files made before team settings existed.
# =============================================================================
def compute_proposal_numbers(q):
    services = q.get('services', {}) or {}

    def optional(include_key, price_key):
        return to_number(services.get(price_key)) if services.get(include_key) else 0.0

    item_1a = optional('include1A', 'price1A')
    item_1b = optional('include1B', 'price1B')
    item_1c = optional('includePermitting', 'permittingPrice')
    item_5 = optional('include5', 'price5')
    item_6 = optional('include6', 'price6')

    routine = services.get('routine', {}) or {}
    storm = services.get('postStorm', {}) or {}
    post = services.get('postConstruction', {}) or {}

    routine_total = to_number(routine.get('total'))
    storm_total = to_number(storm.get('total'))
    post_total = to_number(post.get('total'))
    routine_rate = to_number(routine.get('rate'))

    # Item 3: flat monthly rate = per-inspection rate x premium x multiplier
    construction_months = to_number(q.get('constructionMonths'))
    weekly = bool(q.get('weeklyInspections'))
    multiplier = to_number(q.get('flatRateWeeklyMultiplier' if weekly else 'flatRateBiWeeklyMultiplier'),
                           5.0 if weekly else 2.5)
    premium = to_number(q.get('flatRatePremium'), 1.1)
    flat_monthly_rate = routine_rate * premium * multiplier
    flat_total = flat_monthly_rate * construction_months

    swmp_subtotal = item_1a + item_1b + item_1c
    inspection_subtotal = routine_total + storm_total
    other_subtotal = item_5 + item_6

    return {
        'item_1a': item_1a, 'item_1b': item_1b, 'item_1c': item_1c,
        'item_5': item_5, 'item_6': item_6,
        'routine': routine, 'storm': storm, 'post': post,
        'routine_total': routine_total, 'storm_total': storm_total, 'post_total': post_total,
        'routine_rate': routine_rate,
        'construction_months': construction_months,
        'flat_monthly_rate': flat_monthly_rate,
        'flat_total': flat_total,
        'swmp_subtotal': swmp_subtotal,
        'inspection_subtotal': inspection_subtotal,
        'other_subtotal': other_subtotal,
        'optional_total': swmp_subtotal + other_subtotal,
        # Option 1: per-inspection billing
        'option1_total': swmp_subtotal + inspection_subtotal + post_total + other_subtotal,
        # Option 2: flat monthly billing (post-construction is NOT included,
        # matching the original proposal logic)
        'option2_total': swmp_subtotal + flat_total + other_subtotal,
    }


# =============================================================================
# WORD PROPOSAL
# -----------------------------------------------------------------------------
# The template contains placeholders like {{CLIENT_NAME}}. Each one is replaced
# with the value below. The full list is in README.md. Placeholders can be moved
# anywhere in the template; an unknown placeholder is left visible in the
# finished document so a typo is easy to spot.
# =============================================================================
PLACEHOLDER_RE = re.compile(r"\{\{\s*([A-Z0-9_]+)\s*\}\}")
SCOPE_PLACEHOLDER = 'SCOPE_OF_WORK'


def build_placeholder_values(q, n):
    services = q.get('services', {}) or {}

    def line(include_key, price):
        """Quantity / rate / total for a one-time optional item."""
        if not services.get(include_key):
            return '', '', ''
        return '1', money(price), money(price)

    a_qty, a_rate, a_total = line('include1A', n['item_1a'])
    b_qty, b_rate, b_total = line('include1B', n['item_1b'])
    c_qty, c_rate, c_total = line('includePermitting', n['item_1c'])
    i5_qty, i5_rate, i5_total = line('include5', n['item_5'])
    i6_qty, i6_rate, i6_total = line('include6', n['item_6'])

    prepared_contact = ' / '.join(x for x in [q.get('summitPhone', ''), q.get('summitEmail', '')] if x)

    return {
        # Project & client information
        'PROJECT_NAME': q.get('projectName', ''),
        'CLIENT_NAME': q.get('clientName', ''),
        'PROJECT_CITY': q.get('projectCity', ''),
        'PROJECT_STATE': q.get('projectState', 'Colorado'),
        'PROJECT_COUNTY': q.get('projectCounty', ''),
        'PROJECT_ADDRESS': q.get('projectAddress', ''),
        'PROJECT_TYPE': q.get('projectType', ''),
        'DISTURBED_ACRES': q.get('disturbedAcres', ''),
        'SITE_CONTACT_NAME': q.get('siteContactName', ''),
        'SITE_CONTACT_PHONE': q.get('siteContactPhone', ''),
        'SITE_CONTACT_EMAIL': q.get('siteContactEmail', ''),
        'PREPARED_BY': q.get('summitContact', ''),
        'PREPARED_BY_CONTACT': prepared_contact,
        'TEAM_NAME': q.get('teamName', ''),
        'DATE': datetime.now(MOUNTAIN_TZ).strftime('%B %d, %Y'),
        # Items 1A, 1B, 1C
        'ITEM_1A_QTY': a_qty, 'ITEM_1A_RATE': a_rate, 'ITEM_1A_TOTAL': a_total,
        'ITEM_1B_QTY': b_qty, 'ITEM_1B_RATE': b_rate, 'ITEM_1B_TOTAL': b_total,
        'ITEM_1C_QTY': c_qty, 'ITEM_1C_RATE': c_rate, 'ITEM_1C_TOTAL': c_total,
        'SWMP_SUBTOTAL': money(n['swmp_subtotal']),
        # Items 2A, 2B, 3
        'ROUTINE_QTY': quantity(n['routine'].get('qty')),
        'ROUTINE_RATE': money(to_number(n['routine'].get('rate'))),
        'ROUTINE_TOTAL': money(n['routine_total']),
        'STORM_QTY': quantity(n['storm'].get('qty')),
        'STORM_RATE': money(to_number(n['storm'].get('rate'))),
        'STORM_TOTAL': money(n['storm_total']),
        'STORM_EVENTS_PER_MONTH': f"{to_number(q.get('postStormPerMonth'), 0.6):g}",
        'INSPECTION_SUBTOTAL': money(n['inspection_subtotal']),
        'FLAT_MONTHS': quantity(n['construction_months']),
        'FLAT_MONTHLY_RATE': money(n['flat_monthly_rate']),
        'FLAT_TOTAL': money(n['flat_total']),
        # Item 4
        'POST_QTY': quantity(n['post'].get('qty')),
        'POST_RATE': money(to_number(n['post'].get('rate'))),
        'POST_TOTAL': money(n['post_total']),
        'POST_SUBTOTAL': money(n['post_total']),
        # Items 5, 6
        'ITEM_5_DESC': services.get('tableDescription5', '') if services.get('include5') else '',
        'ITEM_5_QTY': i5_qty, 'ITEM_5_RATE': i5_rate, 'ITEM_5_TOTAL': i5_total,
        'ITEM_6_DESC': services.get('tableDescription6', '') if services.get('include6') else '',
        'ITEM_6_QTY': i6_qty, 'ITEM_6_RATE': i6_rate, 'ITEM_6_TOTAL': i6_total,
        'OTHER_SUBTOTAL': money(n['other_subtotal']),
        # Totals and notes
        'OPTION_1_TOTAL': money(n['option1_total']),
        'OPTION_2_TOTAL': money(n['option2_total']),
        'CONSULTING_RATE': money(to_number(q.get('consultingHourlyRate'), 85)),
    }


def scope_of_work_items(q):
    """(item number, description) pairs for the Scope of Work section."""
    s = q.get('services', {}) or {}
    candidates = [
        ('1.A', 'include1A', 'description1A'),
        ('1.B', 'include1B', 'description1B'),
        ('1.C', 'includePermitting', 'permittingDescription'),
        ('5', 'include5', 'scopeDescription5'),
        ('6', 'include6', 'scopeDescription6'),
    ]
    return [(num, s.get(desc_key)) for num, inc_key, desc_key in candidates
            if s.get(inc_key) and s.get(desc_key)]


def iter_paragraphs(container):
    """Every paragraph in a document, including inside tables (and nested tables)."""
    for paragraph in container.paragraphs:
        yield paragraph
    for table in container.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from iter_paragraphs(cell)


def all_paragraphs(doc):
    yield from iter_paragraphs(doc)
    for section in doc.sections:
        for part in (section.header, section.footer):
            yield from iter_paragraphs(part)


def replace_in_paragraph(paragraph, values, unknown):
    """Replace {{PLACEHOLDERS}} in one paragraph, even if Word split one across
    several text runs. The replacement keeps the formatting of the run where the
    placeholder starts."""
    full_text = ''.join(run.text for run in paragraph.runs)
    if '{{' not in full_text:
        return
    for match in reversed(list(PLACEHOLDER_RE.finditer(full_text))):
        key = match.group(1)
        if key == SCOPE_PLACEHOLDER:
            continue
        if key not in values:
            unknown.add(key)
            continue
        start, end = match.span()
        position, first = 0, True
        for run in paragraph.runs:
            run_start, run_end = position, position + len(run.text)
            position = run_end
            if run_end <= start or run_start >= end:
                continue
            a, b = max(start, run_start) - run_start, min(end, run_end) - run_start
            if first:
                run.text = run.text[:a] + str(values[key]) + run.text[b:]
                first = False
            else:
                run.text = run.text[:a] + run.text[b:]


def fill_scope_of_work(doc, items):
    """Turn the {{SCOPE_OF_WORK}} paragraph into one paragraph per selected item."""
    for paragraph in list(iter_paragraphs(doc)):
        if '{{' + SCOPE_PLACEHOLDER + '}}' not in paragraph.text.replace(' ', ''):
            continue
        template_p = paragraph._p
        run_props = paragraph.runs[0]._r.rPr if paragraph.runs else None
        for number, description in items:
            new_p = copy.deepcopy(template_p)
            for attr in list(new_p.attrib):          # avoid duplicate Word paragraph IDs
                if attr.endswith('}paraId') or attr.endswith('}textId'):
                    del new_p.attrib[attr]
            for r in new_p.findall(qn('w:r')):
                new_p.remove(r)
            template_p.addprevious(new_p)
            new_paragraph = Paragraph(new_p, paragraph._parent)
            number_run = new_paragraph.add_run(f'{number} ')
            text_run = new_paragraph.add_run(description)
            for run in (number_run, text_run):
                if run_props is not None:
                    run._r.insert(0, copy.deepcopy(run_props))
            number_run.bold = True
        template_p.getparent().remove(template_p)


def generate_proposal_docx(quote_data):
    """Fill the Word template with one quote and return the path of the new file."""
    if not os.path.exists(TEMPLATE_FILE):
        raise Exception("Template file not found. Please upload Summit_Proposal_Template.docx to the repository.")

    numbers = compute_proposal_numbers(quote_data)
    values = build_placeholder_values(quote_data, numbers)

    doc = Document(TEMPLATE_FILE)
    fill_scope_of_work(doc, scope_of_work_items(quote_data))
    unknown = set()
    for paragraph in all_paragraphs(doc):
        replace_in_paragraph(paragraph, values, unknown)
    if unknown:
        print(f"WARNING: template has placeholders with no value: {sorted(unknown)}")

    handle, output_file = tempfile.mkstemp(suffix='.docx')
    os.close(handle)
    doc.save(output_file)
    return output_file


# =============================================================================
# GOOGLE SHEETS QUOTE LOG
# =============================================================================
LOG_COLUMNS = [
    'Logged (Mountain Time)', 'Team', 'Prepared By', 'Client', 'Project', 'City', 'County',
    'Project Type', 'Disturbed Acres', 'Nearest Inspector', 'Inspection Frequency',
    'Construction Months', 'Post-Construction Months', 'Calculated Rate', 'Rate Sheet Price',
    'Manual Override Used', 'Routine Rate (Final)', 'Post-Construction Rate (Final)',
    'Optional Services Total', 'Per-Inspection Option Total', 'Flat Monthly Rate',
    'Flat Rate Option Total',
]
_verified_tabs = set()   # tabs already checked since the server started


def get_sheets_service():
    """Connect to the Google Sheets API using the service account in GOOGLE_SHEETS_CREDENTIALS."""
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    creds_json = os.environ.get('GOOGLE_SHEETS_CREDENTIALS')
    if not creds_json:
        raise Exception('GOOGLE_SHEETS_CREDENTIALS is not set in Render')
    credentials = service_account.Credentials.from_service_account_info(
        json.loads(creds_json), scopes=['https://www.googleapis.com/auth/spreadsheets'])
    return build('sheets', 'v4', credentials=credentials, cache_discovery=False)


def ensure_tab(service, spreadsheet_id, tab):
    """Create the team's tab (with a header row) if it does not exist yet."""
    if tab in _verified_tabs:
        return
    meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id, fields='sheets.properties.title').execute()
    titles = [s['properties']['title'] for s in meta.get('sheets', [])]
    if tab not in titles:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'requests': [{'addSheet': {'properties': {'title': tab}}}]}).execute()
    header = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=f"'{tab}'!A1:A1").execute()
    if not header.get('values'):
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range=f"'{tab}'!A1",
            valueInputOption='RAW', body={'values': [LOG_COLUMNS]}).execute()
    _verified_tabs.add(tab)


def build_log_row(q, team):
    n = compute_proposal_numbers(q)
    log = q.get('log', {}) or {}
    return [
        datetime.now(MOUNTAIN_TZ).strftime('%Y-%m-%d %H:%M'),
        team['name'],
        q.get('summitContact', ''),
        q.get('clientName', ''),
        q.get('projectName', ''),
        q.get('projectCity', ''),
        q.get('projectCounty', ''),
        q.get('projectType', ''),
        to_number(q.get('disturbedAcres')),
        q.get('closestInspector', ''),
        'Weekly' if q.get('weeklyInspections') else 'Bi-weekly',
        to_number(q.get('constructionMonths')),
        to_number(q.get('postConstructionMonths')),
        to_number(log.get('calculatedRate')),
        '' if log.get('rateSheetPrice') in (None, '') else to_number(log.get('rateSheetPrice')),
        'Yes' if log.get('overrideUsed') else 'No',
        n['routine_rate'],
        to_number(n['post'].get('rate')),
        round(n['optional_total'], 2),
        round(n['option1_total'], 2),
        round(n['flat_monthly_rate'], 2),
        round(n['option2_total'], 2),
    ]


# =============================================================================
# WEB PAGES
# =============================================================================
PROPOSAL_GENERATOR_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{{ team_name }} | Proposal Generator</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
        }
        .container {
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
            max-width: 600px;
            width: 100%;
            padding: 40px;
        }
        .back-link {
            display: inline-block;
            margin-bottom: 20px;
            color: #667eea;
            text-decoration: none;
            font-size: 14px;
        }
        .back-link:hover { text-decoration: underline; }
        h1 {
            color: #2d3748;
            font-size: 28px;
            margin-bottom: 10px;
            text-align: center;
        }
        .subtitle {
            color: #718096;
            text-align: center;
            margin-bottom: 30px;
            font-size: 14px;
        }
        .upload-area {
            border: 3px dashed #cbd5e0;
            border-radius: 12px;
            padding: 40px;
            text-align: center;
            transition: all 0.3s;
            cursor: pointer;
            background: #f7fafc;
        }
        .upload-area:hover {
            border-color: #667eea;
            background: #edf2f7;
        }
        .upload-area.dragover {
            border-color: #667eea;
            background: #e6f0ff;
        }
        .upload-icon { font-size: 48px; margin-bottom: 20px; }
        .upload-text { color: #4a5568; font-size: 16px; margin-bottom: 10px; }
        .upload-subtext { color: #a0aec0; font-size: 14px; }
        #fileInput { display: none; }
        .file-info {
            margin-top: 20px;
            padding: 15px;
            background: #e6f0ff;
            border-radius: 8px;
            display: none;
        }
        .file-info.show { display: block; }
        .file-name { color: #2d3748; font-weight: 600; margin-bottom: 5px; }
        .file-details { color: #718096; font-size: 14px; }
        .button {
            width: 100%;
            padding: 15px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            border-radius: 8px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            margin-top: 20px;
            transition: transform 0.2s;
            display: none;
        }
        .button.show { display: block; }
        .button:hover { transform: translateY(-2px); }
        .button:disabled { background: #cbd5e0; cursor: not-allowed; }
        .status {
            margin-top: 20px;
            padding: 15px;
            border-radius: 8px;
            text-align: center;
            display: none;
        }
        .status.show { display: block; }
        .status.processing { background: #fef5e7; color: #d68910; }
        .status.success { background: #d4edda; color: #155724; }
        .status.error { background: #f8d7da; color: #721c24; }
        .spinner {
            border: 3px solid #f3f3f3;
            border-top: 3px solid #667eea;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin: 20px auto;
            display: none;
        }
        .spinner.show { display: block; }
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        .instructions {
            margin-top: 30px;
            padding: 20px;
            background: #f7fafc;
            border-radius: 8px;
            border-left: 4px solid #667eea;
        }
        .instructions h3 {
            color: #2d3748;
            font-size: 16px;
            margin-bottom: 10px;
        }
        .instructions ol {
            margin-left: 20px;
            color: #4a5568;
            font-size: 14px;
            line-height: 1.8;
        }
    </style>
</head>
<body>
    <div class="container">
        <a href="{{ back_url }}" class="back-link">← Back to {{ team_name }} Quote Calculator</a>
        
        <h1>📄 Proposal Generator</h1>
        <p class="subtitle">Upload quote JSON to generate Word proposal</p>
        
        <div class="upload-area" id="uploadArea">
            <div class="upload-icon">📁</div>
            <div class="upload-text">Click to select or drag & drop</div>
            <div class="upload-subtext">Upload your quote JSON file</div>
        </div>
        
        <input type="file" id="fileInput" accept=".json">
        
        <div class="file-info" id="fileInfo">
            <div class="file-name" id="fileName"></div>
            <div class="file-details" id="fileDetails"></div>
        </div>
        
        <button class="button" id="generateBtn">Generate Proposal</button>
        <div class="spinner" id="spinner"></div>
        <div class="status" id="status"></div>
        
        <div class="instructions">
            <h3>📋 How to use:</h3>
            <ol>
                <li>After generating a quote, click "Download Quote Data"</li>
                <li>Upload the JSON file here (drag & drop or click)</li>
                <li>Click "Generate Proposal"</li>
                <li>Download your Word document!</li>
            </ol>
        </div>
    </div>
    
    <script>
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('fileInput');
        const fileInfo = document.getElementById('fileInfo');
        const fileName = document.getElementById('fileName');
        const fileDetails = document.getElementById('fileDetails');
        const generateBtn = document.getElementById('generateBtn');
        const spinner = document.getElementById('spinner');
        const status = document.getElementById('status');
        let selectedFile = null;

        uploadArea.addEventListener('click', () => fileInput.click());
        fileInput.addEventListener('change', (e) => handleFile(e.target.files[0]));
        
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });
        uploadArea.addEventListener('dragleave', () => uploadArea.classList.remove('dragover'));
        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            handleFile(e.dataTransfer.files[0]);
        });

        function handleFile(file) {
            if (!file) return;
            if (!file.name.endsWith('.json')) {
                showStatus('error', '❌ Please upload a JSON file');
                return;
            }
            selectedFile = file;
            fileName.textContent = file.name;
            fileDetails.textContent = `${(file.size / 1024).toFixed(1)} KB`;
            fileInfo.classList.add('show');
            generateBtn.classList.add('show');
            status.classList.remove('show');
        }

        generateBtn.addEventListener('click', async () => {
            if (!selectedFile) return;
            try {
                const text = await selectedFile.text();
                const quoteData = JSON.parse(text);
                
                showStatus('processing', '⏳ Generating proposal... (first use may take 30 sec)');
                generateBtn.disabled = true;
                spinner.classList.add('show');

                const formData = new FormData();
                formData.append('quoteData', JSON.stringify(quoteData));

                const response = await fetch('/api/generate-proposal', {
                    method: 'POST',
                    body: formData
                });

                if (!response.ok) {
                    const error = await response.json();
                    throw new Error(error.error || 'Server error');
                }

                const blob = await response.blob();
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `${quoteData.clientName}_${quoteData.projectName}_Proposal`.replace(/[^a-zA-Z0-9_]/g, '_') + '.docx';
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);

                showStatus('success', '✅ Proposal generated! Check your Downloads folder.');
                setTimeout(() => resetForm(), 3000);

            } catch (error) {
                console.error('Error:', error);
                showStatus('error', '❌ Error: ' + error.message);
            } finally {
                generateBtn.disabled = false;
                spinner.classList.remove('show');
            }
        });

        function showStatus(type, message) {
            status.className = 'status show ' + type;
            status.textContent = message;
        }

        function resetForm() {
            selectedFile = null;
            fileInput.value = '';
            fileInfo.classList.remove('show');
            generateBtn.classList.remove('show');
            status.classList.remove('show');
        }
    </script>
</body>
</html>"""


def serve_calculator(team_key):
    team = TEAMS[team_key]
    try:
        with open(CALCULATOR_FILE, 'r', encoding='utf-8') as f:
            html = f.read()
    except FileNotFoundError:
        return ("<h1>Calculator file not found</h1>"
                "<p>Please upload summit-quote-generator.html to the repository.</p>"), 500
    # The page loads teams/summit.js by default; point it at this team's settings
    return html.replace('src="teams/summit.js"', f'src="/teams/{team["settings_file"]}"')


@app.route('/')
def index():
    return serve_calculator('summit')


@app.route('/<team_path>')
def team_calculator(team_path):
    for key, team in TEAMS.items():
        if team['path'] == '/' + team_path:
            return serve_calculator(key)
    abort(404)


@app.route('/teams/<path:filename>')
def team_settings(filename):
    allowed = {team['settings_file'] for team in TEAMS.values()}
    if filename not in allowed:
        abort(404)
    response = send_from_directory(TEAMS_DIR, filename, mimetype='application/javascript')
    response.headers['Cache-Control'] = 'no-cache'   # settings changes show up right away
    return response


@app.route('/generate')
def proposal_page():
    team = get_team(request.args.get('team')) or TEAMS[DEFAULT_TEAM]
    return render_template_string(PROPOSAL_GENERATOR_HTML, team_name=team['name'], back_url=team['path'])


@app.route('/api/generate-proposal', methods=['POST'])
def generate_proposal():
    try:
        quote_data_str = request.form.get('quoteData')
        if not quote_data_str:
            return jsonify({'error': 'No quote data provided'}), 400
        quote_data = json.loads(quote_data_str)
        output_file = generate_proposal_docx(quote_data)

        date_str = datetime.now(MOUNTAIN_TZ).strftime('%Y%m%d')
        inspector_name = quote_data.get('closestInspector', '')
        parts = [quote_data.get('clientName', 'Client'), quote_data.get('projectName', 'Project'),
                 inspector_name.split()[-1] if inspector_name else '', date_str]
        filename = re.sub(r'[^A-Za-z0-9_-]', '_', '_'.join(p for p in parts if p)) + '.docx'

        return send_file(output_file,
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                         as_attachment=True, download_name=filename)
    except Exception as e:
        print(f"Error generating proposal: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/webhook', methods=['POST'])
def webhook():
    """Add one row to the team's tab in the quote log."""
    try:
        q = request.get_json(silent=True)
        if not q:
            return jsonify({'error': 'No data provided'}), 400
        team = get_team(q.get('team'))
        if not team:
            return jsonify({'error': f"Unknown team '{q.get('team')}'"}), 400
        spreadsheet_id = os.environ.get('GOOGLE_SHEET_ID')
        if not spreadsheet_id:
            return jsonify({'error': 'GOOGLE_SHEET_ID is not set in Render'}), 500

        service = get_sheets_service()
        ensure_tab(service, spreadsheet_id, team['sheet_tab'])
        result = service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id, range=f"'{team['sheet_tab']}'!A1",
            valueInputOption='RAW', insertDataOption='INSERT_ROWS',
            body={'values': [build_log_row(q, team)]}).execute()
        return jsonify({'success': True, 'tab': team['sheet_tab'],
                        'updatedRange': result.get('updates', {}).get('updatedRange')}), 200
    except Exception as e:
        print(f"ERROR in webhook: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/health')
def health():
    """Quick check that every file and setting is in place."""
    return jsonify({
        'status': 'healthy',
        'template_found': os.path.exists(TEMPLATE_FILE),
        'calculator_found': os.path.exists(CALCULATOR_FILE),
        'team_settings_found': {key: os.path.exists(os.path.join(TEAMS_DIR, t['settings_file']))
                                for key, t in TEAMS.items()},
        'google_sheet_id_set': bool(os.environ.get('GOOGLE_SHEET_ID')),
        'google_credentials_set': bool(os.environ.get('GOOGLE_SHEETS_CREDENTIALS')),
    })


if __name__ == '__main__':
    # Local testing only. On Render, gunicorn starts the app (see top of file).
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)

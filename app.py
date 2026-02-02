#!/usr/bin/env python3
"""
Summit Quote & Proposal Generator - Combined App
Serves both the quote calculator and proposal generator
"""

from flask import Flask, request, send_file, jsonify, render_template_string
import json
import os
import tempfile
from datetime import datetime
import pytz
from xml.dom import minidom
import zipfile

app = Flask(__name__)

# Read the calculator HTML file
def get_calculator_html():
    """Load the quote calculator HTML"""
    try:
        with open('summit-quote-generator.html', 'r', encoding='utf-8') as f:
            return f.read()
    except FileNotFoundError:
        return """
        <html><body style="font-family: Arial; padding: 50px; text-align: center;">
        <h1>⚠️ Calculator file not found</h1>
        <p>Please upload summit-quote-generator.html to the repository</p>
        </body></html>
        """

# Proposal Generator HTML
PROPOSAL_GENERATOR_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Summit Proposal Generator</title>
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
        <a href="/" class="back-link">← Back to Quote Calculator</a>
        
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
                a.download = `${quoteData.clientName}_${quoteData.projectName}_Proposal.docx`.replace(/[^a-zA-Z0-9_]/g, '_');
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
</html>
"""

def set_cell_text(cell, text, align_right=False):
    """Set text content of a table cell with optional right alignment"""
    paras = cell.getElementsByTagName('w:p')
    if not paras:
        return
    para = paras[0]
    
    # Clear existing content
    runs = para.getElementsByTagName('w:r')
    for run in list(runs):
        para.removeChild(run)
    
    # Add or update paragraph properties for alignment
    pPr = para.getElementsByTagName('w:pPr')
    if pPr:
        # Remove existing pPr
        for pr in list(pPr):
            para.removeChild(pr)
    
    if align_right:
        # Create paragraph properties with right justification
        doc = cell.ownerDocument
        pPr = doc.createElement('w:pPr')
        jc = doc.createElement('w:jc')
        jc.setAttribute('w:val', 'right')
        pPr.appendChild(jc)
        # Insert pPr as first child of paragraph
        if para.firstChild:
            para.insertBefore(pPr, para.firstChild)
        else:
            para.appendChild(pPr)
    
    # Add text content
    doc = cell.ownerDocument
    run = doc.createElement('w:r')
    t = doc.createElement('w:t')
    t.appendChild(doc.createTextNode(str(text)))
    run.appendChild(t)
    para.appendChild(run)

def generate_proposal_docx(quote_data):
    """Generate Word proposal from quote data"""
    
    try:
        # Check if template exists
        template_path = 'Enhanced_Pricing_Template_v3.docx'
        if not os.path.exists(template_path):
            raise Exception("Template file not found. Please upload Enhanced_Pricing_Template_v3.docx to the repository.")
        
        # Create temp directory
        template_dir = tempfile.mkdtemp()
        
        # Unpack DOCX
        with zipfile.ZipFile(template_path, 'r') as zip_ref:
            zip_ref.extractall(template_dir)
        
        # Parse document
        doc_path = os.path.join(template_dir, 'word', 'document.xml')
        doc = minidom.parse(doc_path)
        
        # Get tables
        tables = doc.getElementsByTagName('w:tbl')
        print(f"DEBUG: Found {len(tables)} tables in template")
        
        if len(tables) < 4:
            raise Exception(f"Template needs at least 4 tables, found only {len(tables)}")
        
        # TABLE 1: Project Information (6 rows)
        info_table = tables[0]
        info_rows = info_table.getElementsByTagName('w:tr')
        print(f"DEBUG: Table 1 has {len(info_rows)} rows")
        
        # Row 0: Project Name, Client Name
        cells = info_rows[0].getElementsByTagName('w:tc')
        if len(cells) >= 4:
            set_cell_text(cells[1], quote_data.get('projectName', ''))
            set_cell_text(cells[3], quote_data.get('clientName', ''))
        
        # Row 1: City, Client Contact
        cells = info_rows[1].getElementsByTagName('w:tc')
        if len(cells) >= 4:
            set_cell_text(cells[1], quote_data.get('projectCity', ''))
            set_cell_text(cells[3], '')  # Client contact not in our data
        
        # Row 2: State, Contact Phone
        cells = info_rows[2].getElementsByTagName('w:tc')
        if len(cells) >= 4:
            set_cell_text(cells[1], 'Colorado')  # Default to Colorado
            set_cell_text(cells[3], '')  # Contact phone not in our data
        
        # Row 3: Jurisdiction, Contact Email
        cells = info_rows[3].getElementsByTagName('w:tc')
        if len(cells) >= 4:
            set_cell_text(cells[1], quote_data.get('projectCounty', ''))
            set_cell_text(cells[3], '')  # Contact email not in our data
        
        # Row 4: Project Type, Prepared by
        cells = info_rows[4].getElementsByTagName('w:tc')
        if len(cells) >= 4:
            set_cell_text(cells[1], quote_data.get('projectType', ''))
            set_cell_text(cells[3], quote_data.get('summitContact', ''))
        
        # Row 5: Size in Acres, Phone/email
        cells = info_rows[5].getElementsByTagName('w:tc')
        if len(cells) >= 4:
            set_cell_text(cells[1], quote_data.get('disturbedAcres', ''))
            contact_info = f"{quote_data.get('summitPhone', '')} / {quote_data.get('summitEmail', '')}"
            set_cell_text(cells[3], contact_info)
        
        services = quote_data.get('services', {})
        
        # TABLE 2: SWMP Services (Items 1A, 1B, 1C)
        swmp_table = tables[1]
        swmp_rows = swmp_table.getElementsByTagName('w:tr')
        
        # Row 1 (1A): SWMP Narrative
        if services.get('include1A'):
            price1A = float(services.get('price1A', 0)) if services.get('price1A') else 0
            if len(swmp_rows) > 1:
                cells = swmp_rows[1].getElementsByTagName('w:tc')
                if len(cells) >= 5:
                    set_cell_text(cells[2], '1', align_right=True)  # Quantity
                    set_cell_text(cells[3], f'${price1A:,.0f}', align_right=True)  # Rate
                    set_cell_text(cells[4], f'${price1A:,.0f}', align_right=True)  # Total
        
        # Row 2 (1B): SWMP Site Plan
        if services.get('include1B'):
            price1B = float(services.get('price1B', 0)) if services.get('price1B') else 0
            if len(swmp_rows) > 2:
                cells = swmp_rows[2].getElementsByTagName('w:tc')
                if len(cells) >= 5:
                    set_cell_text(cells[2], '1', align_right=True)  # Quantity
                    set_cell_text(cells[3], f'${price1B:,.0f}', align_right=True)  # Rate
                    set_cell_text(cells[4], f'${price1B:,.0f}', align_right=True)  # Total
        
        # Row 3 (1C): Permitting
        if services.get('includePermitting'):
            permitting_price = float(services.get('permittingPrice', 0)) if services.get('permittingPrice') else 0
            if len(swmp_rows) > 3:
                cells = swmp_rows[3].getElementsByTagName('w:tc')
                if len(cells) >= 5:
                    set_cell_text(cells[2], '1', align_right=True)  # Quantity - right aligned
                    set_cell_text(cells[3], f'${permitting_price:,.0f}', align_right=True)  # Rate - right aligned
                    set_cell_text(cells[4], f'${permitting_price:,.0f}', align_right=True)  # Total - right aligned
        
        # TABLE 3: Inspection Services (Items 2A, 2B, 2C)
        inspection_table = tables[2]
        inspection_rows = inspection_table.getElementsByTagName('w:tr')
        print(f"DEBUG: Table 3 (inspections) has {len(inspection_rows)} rows")
        
        # Row 1 (2A): Routine Inspections
        routine = services.get('routine', {})
        if routine and len(inspection_rows) > 1:
            cells = inspection_rows[1].getElementsByTagName('w:tc')
            if len(cells) >= 5:
                set_cell_text(cells[2], routine.get('qty', ''), align_right=True)
                set_cell_text(cells[3], f"${routine.get('rate', 0):,.0f}", align_right=True)
                set_cell_text(cells[4], f"${routine.get('total', 0):,.0f}", align_right=True)
        
        # Row 2 (2B): Post-Storm Inspections
        storm = services.get('postStorm', {})
        if storm and len(inspection_rows) > 2:
            cells = inspection_rows[2].getElementsByTagName('w:tc')
            if len(cells) >= 5:
                set_cell_text(cells[2], storm.get('qty', ''), align_right=True)
                set_cell_text(cells[3], f"${storm.get('rate', 0):,.0f}", align_right=True)
                set_cell_text(cells[4], f"${storm.get('total', 0):,.0f}", align_right=True)
        
        # Row 3 (2C): Post-Construction Inspections
        post = services.get('postConstruction', {})
        if post and len(inspection_rows) > 3:
            cells = inspection_rows[3].getElementsByTagName('w:tc')
            if len(cells) >= 5:
                set_cell_text(cells[2], post.get('qty', ''), align_right=True)
                set_cell_text(cells[3], f"${post.get('rate', 0):,.0f}", align_right=True)
                set_cell_text(cells[4], f"${post.get('total', 0):,.0f}", align_right=True)
        
        # Calculate inspection subtotal
        routine_total = routine.get('total', 0) if routine else 0
        storm_total = storm.get('total', 0) if storm else 0
        post_total = post.get('total', 0) if post else 0
        inspection_subtotal = routine_total + storm_total + post_total
        
        # Row 4: Inspection Subtotal
        if len(inspection_rows) > 4:
            cells = inspection_rows[4].getElementsByTagName('w:tc')
            if len(cells) >= 5:
                set_cell_text(cells[4], f"${inspection_subtotal:,.0f}", align_right=True)
        
        # TABLE 4: Other Services (Items 4, 5)
        other_table = tables[3]
        other_rows = other_table.getElementsByTagName('w:tr')
        
        # Row 1 (Item 4)
        if services.get('include4'):
            price4 = float(services.get('price4', 0)) if services.get('price4') else 0
            if len(other_rows) > 1:
                cells = other_rows[1].getElementsByTagName('w:tc')
                if len(cells) >= 5:
                    set_cell_text(cells[2], '1', align_right=True)  # Quantity
                    set_cell_text(cells[3], f'${price4:,.0f}', align_right=True)  # Rate
                    set_cell_text(cells[4], f'${price4:,.0f}', align_right=True)  # Total
        
        # Row 2 (Item 5)
        if services.get('include5'):
            price5 = float(services.get('price5', 0)) if services.get('price5') else 0
            if len(other_rows) > 2:
                cells = other_rows[2].getElementsByTagName('w:tc')
                if len(cells) >= 5:
                    set_cell_text(cells[2], '1', align_right=True)  # Quantity
                    set_cell_text(cells[3], f'${price5:,.0f}', align_right=True)  # Rate
                    set_cell_text(cells[4], f'${price5:,.0f}', align_right=True)  # Total
        
        # TABLE 5: Project Total
        total_table = tables[4]
        total_rows = total_table.getElementsByTagName('w:tr')
        if len(total_rows) > 1:
            cells = total_rows[1].getElementsByTagName('w:tc')
            if len(cells) >= 3:
                project_total = quote_data.get('projectTotal', 0)
                set_cell_text(cells[2], f"${project_total:,.0f}", align_right=True)
        
        # Fill descriptions in SCOPE OF WORK DESCRIPTION section
        all_paras = doc.getElementsByTagName('w:p')
        
        # Find the SCOPE OF WORK DESCRIPTION section
        scope_para_index = None
        for i, para in enumerate(all_paras):
            texts = para.getElementsByTagName('w:t')
            text_content = ''.join([t.firstChild.nodeValue if t.firstChild else '' for t in texts])
            
            if 'SCOPE OF WORK DESCRIPTION' in text_content:
                scope_para_index = i
                break
        
        if scope_para_index is not None:
            # Collect all descriptions to add
            descriptions_to_add = []
            
            if services.get('include1A') and services.get('description1A'):
                descriptions_to_add.append(('1.A', services.get('description1A')))
            
            if services.get('include1B') and services.get('description1B'):
                descriptions_to_add.append(('1.B', services.get('description1B')))
            
            if services.get('includePermitting') and services.get('permittingDescription'):
                descriptions_to_add.append(('1.C', services.get('permittingDescription')))
            
            if services.get('include4') and services.get('description4'):
                descriptions_to_add.append(('4', services.get('description4')))
            
            if services.get('include5') and services.get('description5'):
                descriptions_to_add.append(('5', services.get('description5')))
            
            # Add each description to consecutive paragraphs
            for idx, (item_num, description) in enumerate(descriptions_to_add):
                para_offset = scope_para_index + 1 + idx
                if para_offset < len(all_paras):
                    target_para = all_paras[para_offset]
                    
                    # Clear any existing content
                    for run in list(target_para.getElementsByTagName('w:r')):
                        target_para.removeChild(run)
                    
                    # Add item number in bold
                    run1 = doc.createElement('w:r')
                    rpr1 = doc.createElement('w:rPr')
                    bold1 = doc.createElement('w:b')
                    rpr1.appendChild(bold1)
                    run1.appendChild(rpr1)
                    t1 = doc.createElement('w:t')
                    t1.setAttribute('xml:space', 'preserve')
                    t1.appendChild(doc.createTextNode(f'{item_num} '))
                    run1.appendChild(t1)
                    target_para.appendChild(run1)
                    
                    # Add the description (not bold)
                    run2 = doc.createElement('w:r')
                    t2 = doc.createElement('w:t')
                    t2.setAttribute('xml:space', 'preserve')
                    t2.appendChild(doc.createTextNode(description))
                    run2.appendChild(t2)
                    target_para.appendChild(run2)
        
        print("DEBUG: Successfully filled all template fields")
        
        # Write updated document
        with open(doc_path, 'w', encoding='utf-8') as f:
            f.write(doc.toxml())
        
        # Repack as DOCX
        output_file = tempfile.mktemp(suffix='.docx')
        with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as docx:
            for root, dirs, files in os.walk(template_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, template_dir)
                    docx.write(file_path, arcname)
        
        print("DEBUG: Successfully created output file")
        return output_file
        
    except Exception as e:
        print(f"ERROR in generate_proposal_docx: {e}")
        import traceback
        traceback.print_exc()
        raise

@app.route('/')
def index():
    """Serve the quote calculator"""
    return get_calculator_html()

@app.route('/generate')
def proposal_page():
    """Serve the proposal generator page"""
    try:
        return render_template_string(PROPOSAL_GENERATOR_HTML)
    except Exception as e:
        print(f"Error serving proposal page: {e}")
        import traceback
        traceback.print_exc()
        return f"""
        <html><body style="font-family: Arial; padding: 50px;">
        <h1>Error Loading Proposal Generator</h1>
        <p>Error: {str(e)}</p>
        <p><a href="/">← Back to Calculator</a></p>
        </body></html>
        """, 500

@app.route('/api/generate-proposal', methods=['POST'])
def generate_proposal():
    """API endpoint to generate proposals"""
    try:
        quote_data_str = request.form.get('quoteData')
        if not quote_data_str:
            return jsonify({'error': 'No quote data provided'}), 400
        
        quote_data = json.loads(quote_data_str)
        output_file = generate_proposal_docx(quote_data)
        
        # Generate filename with date
        client_name = quote_data.get('clientName', 'Client')
        project_name = quote_data.get('projectName', 'Project')
        mountain_tz = pytz.timezone('America/Denver')
        today = datetime.now(mountain_tz)
        date_str = today.strftime('%Y%m%d')
        filename = f"{client_name}_{project_name}_{date_str}.docx".replace(' ', '_')
        
        return send_file(
            output_file,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=filename
        )
    
    except Exception as e:
        print(f"Error generating proposal: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/health')
def health():
    """Health check endpoint"""
    template_exists = os.path.exists('Enhanced_Pricing_Template_Blank__2_.docx')
    calculator_exists = os.path.exists('summit-quote-generator.html')
    
    # List all files in current directory
    files = os.listdir('.')
    
    return jsonify({
        'status': 'healthy',
        'template_found': template_exists,
        'calculator_found': calculator_exists,
        'files_in_directory': files,
        'current_directory': os.getcwd()
    })

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    print(f"🚀 Starting Summit Quote & Proposal Generator on port {port}")
    app.run(host='0.0.0.0', port=port)

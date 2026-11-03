/* =============================================================================
   TEAM SETTINGS: Summit Services Group
   -----------------------------------------------------------------------------
   Every pricing assumption the quote calculator uses lives in this file.
   To change a number: edit it, save, and push to GitHub. Render redeploys.

   Rules for editing (so the page keeps working):
     * Change only the values to the right of the colons.
     * Keep the commas at the ends of lines, and the quote marks around text.
     * Numbers have no $ or % signs: write 100, not $100.
   ========================================================================== */

window.TEAM_SETTINGS = {

  // Must match a team key in app.py (this is how quotes reach the right Google Sheet tab)
  teamKey: 'summit',
  teamName: 'Summit Services Group',
  pageTitle: 'Summit Inspection Quote Generator',


  // ---------------------------------------------------------------------------
  // PRICING
  // ---------------------------------------------------------------------------
  pricing: {
    hourlyRate: 100,               // $/hour for drive time and inspection time (staff can change per quote)
    adminFee: 50,                  // $ added to every inspection
    roundUpTo: 5,                  // round inspection rates UP to the nearest $5
    postConstructionPercent: 85,   // post-construction rate = this % of the routine rate (staff can change per quote)
    consultingHourlyRate: 85,      // $/hour shown in the proposal's Notes (item 4 of Notes and assumptions)
  },


  // ---------------------------------------------------------------------------
  // DRIVE TIME
  // Distance is measured "as the crow flies" from the nearest inspector's home.
  // ---------------------------------------------------------------------------
  driveTime: {
    averageSpeedMph: 45,           // average driving speed used to turn miles into hours
    roundTripThresholdMiles: 50,   // under this distance bill one-way miles; at or over it bill the round trip
    roadDistanceFactor: 1.0,       // roads are longer than a straight line; 1.0 = no adjustment
                                   // (example: 1.3 adds 30% to every distance)
  },


  // ---------------------------------------------------------------------------
  // TIME ON SITE (hours per inspection, by disturbed acres)
  // Each line means "less than maxAcres acres takes this many hours".
  // Projects at or above the last maxAcres get a "custom quote" message.
  // ---------------------------------------------------------------------------
  inspectionTime: [
    { maxAcres: 5,  hours: 0.5 },
    { maxAcres: 10, hours: 0.65 },
    { maxAcres: 25, hours: 1.0 },
    { maxAcres: 50, hours: 1.5 },
    { maxAcres: 75, hours: 2.0 },
  ],

  // Linear projects take longer to walk, so their on-site time is multiplied
  linearProjects: {
    types: ['Transmission Line', 'Pipeline', 'Highway'],
    timeMultiplier: 1.2,           // 1.2 = 20% more time on site
  },


  // ---------------------------------------------------------------------------
  // INSPECTION COUNTS
  // ---------------------------------------------------------------------------
  inspectionFrequency: {
    biWeeklyPerMonth: 2,           // routine inspections per month, bi-weekly schedule (the default)
    weeklyPerMonth: 4,             // routine inspections per month, weekly schedule
    postStormPerMonth: 0.6,        // budgeted storm-event inspections per month of construction
                                   // (also printed in the proposal's item 2B description)
    postConstructionPerMonth: 1,   // post-construction inspections per month
  },


  // ---------------------------------------------------------------------------
  // FLAT MONTHLY RATE (proposal item 3)
  // Flat monthly rate = routine rate x premium x multiplier
  // ---------------------------------------------------------------------------
  flatRate: {
    biWeeklyMultiplier: 2.5,       // default multiplier for bi-weekly sites (staff can change per quote)
    weeklyMultiplier: 5.0,         // default multiplier for weekly sites (staff can change per quote)
    premium: 1.1,                  // 1.1 = flat rate costs 10% more than paying per inspection
  },


  // ---------------------------------------------------------------------------
  // RATE SHEET (used only for the side-by-side comparison on screen)
  // ---------------------------------------------------------------------------
  rateSheet: [
    { min: 0,  max: 1,  price: 105, label: 'Less than 1 acre' },
    { min: 1,  max: 2,  price: 125, label: '1-2 acres' },
    { min: 2,  max: 4,  price: 140, label: '2-4 acres' },
    { min: 4,  max: 5,  price: 140, label: '4-5 acres' },
    { min: 5,  max: 10, price: 170, label: '5-10 acres' },
    { min: 10, max: 25, price: 200, label: '10-25 acres' },
    { min: 25, max: 50, price: 230, label: '25-50 acres' },
    { min: 50, max: 75, price: 260, label: '50-75 acres' },
  ],
  closeMatchPercent: 10,           // within this % of the rate sheet counts as a "close match"


  // ---------------------------------------------------------------------------
  // PROJECT TYPES (the first one is the default in the dropdown)
  // ---------------------------------------------------------------------------
  projectTypes: ['Residential', 'Commercial', 'Multi-Family', 'Transmission Line', 'Pipeline', 'Highway'],


  // ---------------------------------------------------------------------------
  // INSPECTORS
  // Source: Inspector_home_locations (2).xlsx, November 2026 roster.
  // lat/lon are approximate TOWN-CENTER coordinates, not exact home addresses.
  // active: false keeps someone on the list without assigning them quotes
  // (office staff were excluded in the original calculator).
  // ---------------------------------------------------------------------------
  inspectors: [
    { name: 'Alexander C McNiel',      town: 'Denver',           lat: 39.7316, lon: -104.9576, active: true },
    { name: 'Anthony Cipperly',        town: 'Palisade',         lat: 39.1103, lon: -108.3517, active: true },
    { name: 'Britlyn R Romero',        town: 'Fruita',           lat: 39.1592, lon: -108.7289, active: true },
    { name: 'Brooks M Bosco',          town: 'Fruita',           lat: 39.1592, lon: -108.7289, active: true },
    { name: 'Camden J. Sherrill',      town: 'Grand Junction',   lat: 39.0639, lon: -108.5506, active: true },
    { name: 'Dustin M Graf',           town: 'Grand Junction',   lat: 39.0639, lon: -108.5506, active: true },
    { name: 'Haley Hinger',            town: 'Lafayette',        lat: 39.9936, lon: -105.0897, active: true },
    { name: 'Jesse R Tietz',           town: 'Grand Junction',   lat: 39.0639, lon: -108.5506, active: true },
    { name: 'Katherine Bergen',        town: 'Longmont',         lat: 40.1672, lon: -105.1019, active: true },
    { name: 'Max Hall',                town: 'Colorado Springs', lat: 38.8339, lon: -104.8214, active: true },
    { name: 'Natasha N Jennings',      town: 'Fruita',           lat: 39.1592, lon: -108.7289, active: true },
    { name: 'Nick Brach',              town: 'Loma',             lat: 39.1950, lon: -108.8126, active: true },
    { name: 'Robert Kyle West',        town: 'Bailey',           lat: 39.4081, lon: -105.4769, active: true },
    { name: 'Roxanne L Fromang',       town: 'Rangely',          lat: 40.0875, lon: -108.8048, active: true },
    { name: 'Taylor Hickman',          town: 'Broomfield',       lat: 39.9205, lon: -105.0867, active: true },
    { name: 'Travis Williams',         town: 'Fruita',           lat: 39.1592, lon: -108.7289, active: true },
    // Office staff (on the roster, not assigned quotes):
    { name: 'Daniel C Kintz',          town: 'Grand Junction',   lat: 39.0639, lon: -108.5506, active: false },
    { name: 'Matthew T Fought',        town: 'Grand Junction',   lat: 39.0639, lon: -108.5506, active: false },
    { name: 'Mr. Andrew S Wallenkamp', town: 'Denver',           lat: 39.7589, lon: -104.9671, active: false },
    { name: 'Stefan P Majewski',       town: 'Denver',           lat: 39.7539, lon: -105.0047, active: false },
  ],
};

// Masters Pick'em — Google Apps Script Backend
// Deploy as: Web App, Execute as: Me, Who has access: Anyone
//
// SETUP: Set SHEET_ID below to your Google Sheet's ID (from the URL).

const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID_HERE';

// ── Sheet column layouts ────────────────────────────────────────────────────
// Golfers sheet:  rank | name | odds | tier
// Picks sheet:    name | tierA | tierB | tierC | alternate | birdieGuess | timestamp

const GOLFERS_SHEET = 'Golfers';
const PICKS_SHEET   = 'Picks';

// ── doGet ───────────────────────────────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action || '';
  try {
    if (action === 'golfers') return respond(getGolfers());
    if (action === 'picks')   return respond(getPicks());
    return respond({ error: 'Unknown action' }, 400);
  } catch (err) {
    return respond({ error: err.message }, 500);
  }
}

// ── doPost ──────────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action || '';
    if (action === 'submit')          return respond(submitPick(payload));
    if (action === 'updateBirdies')   return respond(updateBirdies(payload));
    return respond({ error: 'Unknown action' }, 400);
  } catch (err) {
    return respond({ error: err.message }, 500);
  }
}

// ── Golfers ─────────────────────────────────────────────────────────────────
function getGolfers() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(GOLFERS_SHEET);
  const rows  = sheet.getDataRange().getValues();
  // skip header row
  return rows.slice(1).map(r => ({
    rank: r[0],
    name: r[1],
    odds: r[2],
    tier: r[3],
  }));
}

// ── Picks ───────────────────────────────────────────────────────────────────
function getPicks() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(PICKS_SHEET);
  const rows  = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  return rows.slice(1).map(r => ({
    name:        r[0],
    tierA:       r[1],
    tierB:       r[2],
    tierC:       r[3],
    alternate:   r[4],
    birdieGuess: r[5],
    timestamp:   r[6],
  }));
}

function submitPick(payload) {
  const { name, tierA, tierB, tierC, alternate, birdieGuess } = payload;
  if (!name || !tierA || !tierB || !tierC || !alternate) {
    throw new Error('Missing required fields');
  }
  const ts    = new Date().toISOString();
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(PICKS_SHEET);
  const rows  = sheet.getDataRange().getValues();

  // Find existing pick by player name (case-insensitive)
  const normalName = name.trim().toLowerCase();
  let found = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === normalName) {
      found = i + 1; // 1-indexed sheet row
      break;
    }
  }

  const rowData = [name.trim(), tierA, tierB, tierC, alternate, birdieGuess || '', ts];
  if (found > -1) {
    sheet.getRange(found, 1, 1, rowData.length).setValues([rowData]);
    return { status: 'updated', name: name.trim() };
  } else {
    sheet.appendRow(rowData);
    return { status: 'created', name: name.trim() };
  }
}

// ── Admin: set actual total birdies (for tiebreaker resolution) ─────────────
function updateBirdies(payload) {
  const { birdies, adminKey } = payload;
  if (adminKey !== getAdminKey_()) throw new Error('Unauthorized');
  const ss   = SpreadsheetApp.openById(SHEET_ID);
  const prop = PropertiesService.getScriptProperties();
  prop.setProperty('ACTUAL_BIRDIES', String(birdies));
  return { status: 'ok', birdies };
}

function getActualBirdies() {
  const prop = PropertiesService.getScriptProperties();
  const v    = prop.getProperty('ACTUAL_BIRDIES');
  return v ? Number(v) : null;
}

// Store admin key in Script Properties (set manually via Apps Script UI → Project Settings)
function getAdminKey_() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_KEY') || 'masters2026';
}

// ── Helpers ─────────────────────────────────────────────────────────────────
function respond(data, code) {
  const output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ── One-time setup: initialise the spreadsheet ──────────────────────────────
function initSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // Golfers tab
  let golfers = ss.getSheetByName(GOLFERS_SHEET);
  if (!golfers) golfers = ss.insertSheet(GOLFERS_SHEET);
  if (golfers.getLastRow() === 0) {
    golfers.appendRow(['rank', 'name', 'odds', 'tier']);
    const data = GOLFER_DATA.map((g, i) => [g.rank, g.name, g.odds, g.tier]);
    golfers.getRange(2, 1, data.length, 4).setValues(data);
  }

  // Picks tab
  let picks = ss.getSheetByName(PICKS_SHEET);
  if (!picks) picks = ss.insertSheet(PICKS_SHEET);
  if (picks.getLastRow() === 0) {
    picks.appendRow(['name', 'tierA', 'tierB', 'tierC', 'alternate', 'birdieGuess', 'timestamp']);
  }

  Logger.log('Sheet initialized.');
}

// ── Golfer seed data ─────────────────────────────────────────────────────────
const GOLFER_DATA = [
  // Tier A (1-10)
  { rank:  1, name: 'Scottie Scheffler',          odds: '+350',    tier: 'A' },
  { rank:  2, name: 'Rory McIlroy',               odds: '+700',    tier: 'A' },
  { rank:  3, name: 'Bryson DeChambeau',           odds: '+1000',   tier: 'A' },
  { rank:  4, name: 'Jon Rahm',                    odds: '+1200',   tier: 'A' },
  { rank:  5, name: 'Ludvig Aberg',                odds: '+1600',   tier: 'A' },
  { rank:  6, name: 'Xander Schauffele',           odds: '+1800',   tier: 'A' },
  { rank:  7, name: 'Tommy Fleetwood',             odds: '+1800',   tier: 'A' },
  { rank:  8, name: 'Collin Morikawa',             odds: '+2200',   tier: 'A' },
  { rank:  9, name: 'Cameron Young',               odds: '+2700',   tier: 'A' },
  { rank: 10, name: 'Patrick Reed',                odds: '+3000',   tier: 'A' },
  // Tier B (11-20)
  { rank: 11, name: 'Matt Fitzpatrick',            odds: '+3000',   tier: 'B' },
  { rank: 12, name: 'Justin Rose',                 odds: '+3000',   tier: 'B' },
  { rank: 13, name: 'Chris Gotterup',              odds: '+3500',   tier: 'B' },
  { rank: 14, name: 'Hideki Matsuyama',            odds: '+3500',   tier: 'B' },
  { rank: 15, name: 'Viktor Hovland',              odds: '+3500',   tier: 'B' },
  { rank: 16, name: 'Brooks Koepka',               odds: '+3800',   tier: 'B' },
  { rank: 17, name: 'Justin Thomas',               odds: '+4000',   tier: 'B' },
  { rank: 18, name: 'Robert MacIntyre',            odds: '+4000',   tier: 'B' },
  { rank: 19, name: 'Jordan Spieth',               odds: '+4000',   tier: 'B' },
  { rank: 20, name: 'Tyrrell Hatton',              odds: '+4000',   tier: 'B' },
  // Tier C (21-30)
  { rank: 21, name: 'Shane Lowry',                 odds: '+4500',   tier: 'C' },
  { rank: 22, name: 'Patrick Cantlay',             odds: '+5000',   tier: 'C' },
  { rank: 23, name: 'Joaquin Niemann',             odds: '+5000',   tier: 'C' },
  { rank: 24, name: 'Ben Griffin',                 odds: '+5500',   tier: 'C' },
  { rank: 25, name: 'Si Woo Kim',                  odds: '+6000',   tier: 'C' },
  { rank: 26, name: 'Jake Knapp',                  odds: '+6000',   tier: 'C' },
  { rank: 27, name: 'Corey Conners',               odds: '+6000',   tier: 'C' },
  { rank: 28, name: 'Akshay Bhatia',               odds: '+6000',   tier: 'C' },
  { rank: 29, name: 'Russell Henley',              odds: '+6600',   tier: 'C' },
  { rank: 30, name: 'Min Woo Lee',                 odds: '+6600',   tier: 'C' },
  // Alternate / C+ (31-101)
  { rank: 31, name: 'Max Homa',                    odds: '+6600',   tier: 'Alt' },
  { rank: 32, name: 'Jason Day',                   odds: '+6600',   tier: 'Alt' },
  { rank: 33, name: 'Cameron Smith',               odds: '+6600',   tier: 'Alt' },
  { rank: 34, name: 'Adam Scott',                  odds: '+6600',   tier: 'Alt' },
  { rank: 35, name: 'Sepp Straka',                 odds: '+7000',   tier: 'Alt' },
  { rank: 36, name: 'Sam Burns',                   odds: '+7000',   tier: 'Alt' },
  { rank: 37, name: 'Daniel Berger',               odds: '+7000',   tier: 'Alt' },
  { rank: 38, name: 'Sung-Jae Im',                 odds: '+8000',   tier: 'Alt' },
  { rank: 39, name: 'Sahith Theegala',             odds: '+8000',   tier: 'Alt' },
  { rank: 40, name: 'Gary Woodland',               odds: '+8000',   tier: 'Alt' },
  { rank: 41, name: 'Marco Penge',                 odds: '+8000',   tier: 'Alt' },
  { rank: 42, name: 'Wyndham Clark',               odds: '+8000',   tier: 'Alt' },
  { rank: 43, name: 'Will Zalatoris',              odds: '+8000',   tier: 'Alt' },
  { rank: 44, name: 'Nicolai Hojgaard',            odds: '+9000',   tier: 'Alt' },
  { rank: 45, name: 'Jacob Bridgeman',             odds: '+9000',   tier: 'Alt' },
  { rank: 46, name: 'J.J. Spaun',                  odds: '+9000',   tier: 'Alt' },
  { rank: 47, name: 'Harris English',              odds: '+9000',   tier: 'Alt' },
  { rank: 48, name: 'Dustin Johnson',              odds: '+9000',   tier: 'Alt' },
  { rank: 49, name: 'Sergio Garcia',               odds: '+10000',  tier: 'Alt' },
  { rank: 50, name: 'Matt McCarty',                odds: '+10000',  tier: 'Alt' },
  { rank: 51, name: 'Alex Noren',                  odds: '+10000',  tier: 'Alt' },
  { rank: 52, name: 'Tony Finau',                  odds: '+10000',  tier: 'Alt' },
  { rank: 53, name: 'Maverick McNealy',            odds: '+11000',  tier: 'Alt' },
  { rank: 54, name: 'Ryan Gerard',                 odds: '+11000',  tier: 'Alt' },
  { rank: 55, name: 'Ryan Fox',                    odds: '+12000',  tier: 'Alt' },
  { rank: 56, name: 'Kurt Kitayama',               odds: '+12000',  tier: 'Alt' },
  { rank: 57, name: 'Keegan Bradley',              odds: '+12000',  tier: 'Alt' },
  { rank: 58, name: 'Harry Hall',                  odds: '+13000',  tier: 'Alt' },
  { rank: 59, name: 'Aaron Rai',                   odds: '+13000',  tier: 'Alt' },
  { rank: 60, name: 'Pierceson Coody',             odds: '+13000',  tier: 'Alt' },
  { rank: 61, name: 'John Keefer',                 odds: '+13000',  tier: 'Alt' },
  { rank: 62, name: 'Rasmus Neergaard-Petersen',   odds: '+13000',  tier: 'Alt' },
  { rank: 63, name: 'Davis Thompson',              odds: '+14000',  tier: 'Alt' },
  { rank: 64, name: 'Tom McKibbin',                odds: '+14000',  tier: 'Alt' },
  { rank: 65, name: 'Taylor Pendrith',             odds: '+15000',  tier: 'Alt' },
  { rank: 66, name: 'Tiger Woods',                 odds: '+15000',  tier: 'Alt' },
  { rank: 67, name: 'Rasmus Hojgaard',             odds: '+15000',  tier: 'Alt' },
  { rank: 68, name: 'Carlos Ortiz',                odds: '+15000',  tier: 'Alt' },
  { rank: 69, name: 'Brian Harman',                odds: '+15000',  tier: 'Alt' },
  { rank: 70, name: 'Sam Stevens',                 odds: '+15000',  tier: 'Alt' },
  { rank: 71, name: 'Phil Mickelson',              odds: '+20000',  tier: 'Alt' },
  { rank: 72, name: 'Michael Kim',                 odds: '+20000',  tier: 'Alt' },
  { rank: 73, name: 'Casey Jarvis',                odds: '+20000',  tier: 'Alt' },
  { rank: 74, name: 'Nico Echavarria',             odds: '+20000',  tier: 'Alt' },
  { rank: 75, name: 'Andrew Novak',                odds: '+20000',  tier: 'Alt' },
  { rank: 76, name: 'Aldrich Potgieter',           odds: '+20000',  tier: 'Alt' },
  { rank: 77, name: 'Jayden Schaper',              odds: '+20000',  tier: 'Alt' },
  { rank: 78, name: 'Nick Taylor',                 odds: '+25000',  tier: 'Alt' },
  { rank: 79, name: 'Max Greyserman',              odds: '+25000',  tier: 'Alt' },
  { rank: 80, name: 'Haotong Li',                  odds: '+25000',  tier: 'Alt' },
  { rank: 81, name: 'Bubba Watson',                odds: '+25000',  tier: 'Alt' },
  { rank: 82, name: 'Sami Valimaki',               odds: '+30000',  tier: 'Alt' },
  { rank: 83, name: 'Davis Riley',                 odds: '+30000',  tier: 'Alt' },
  { rank: 84, name: 'Kristoffer Reitan',           odds: '+30000',  tier: 'Alt' },
  { rank: 85, name: 'Charl Schwartzel',            odds: '+35000',  tier: 'Alt' },
  { rank: 86, name: 'Brian Campbell',              odds: '+40000',  tier: 'Alt' },
  { rank: 87, name: 'Michael Brennan',             odds: '+40000',  tier: 'Alt' },
  { rank: 88, name: 'Zach Johnson',                odds: '+40000',  tier: 'Alt' },
  { rank: 89, name: 'Danny Willett',               odds: '+50000',  tier: 'Alt' },
  { rank: 90, name: 'Angel Cabrera',               odds: '+50000',  tier: 'Alt' },
  { rank: 91, name: 'Vijay Singh',                 odds: '+100000', tier: 'Alt' },
  { rank: 92, name: 'Mike Weir',                   odds: '+100000', tier: 'Alt' },
  { rank: 93, name: 'Fred Couples',                odds: '+100000', tier: 'Alt' },
  { rank: 94, name: 'Brandon Holtz',               odds: '+200000', tier: 'Alt' },
  { rank: 95, name: 'Mason Howell',                odds: '+200000', tier: 'Alt' },
  { rank: 96, name: 'Jose Maria Olazabal',         odds: '+200000', tier: 'Alt' },
  { rank: 97, name: 'Mateo Pulcini',               odds: '+200000', tier: 'Alt' },
  { rank: 98, name: 'Jackson Herrington',          odds: '+200000', tier: 'Alt' },
  { rank: 99, name: 'Ethan Fang',                  odds: '+200000', tier: 'Alt' },
  { rank:100, name: 'Pongsapak Laopakdee',         odds: '+200000', tier: 'Alt' },
  { rank:101, name: 'Naoyuki Kataoka',             odds: '+200000', tier: 'Alt' },
];

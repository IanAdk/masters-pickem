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
    if (action === 'config')  return respond(getConfig());
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
    if (action === 'submit')        return respond(submitPick(payload));
    if (action === 'updateBirdies') return respond(updateBirdies(payload));
    if (action === 'resetPicks')    return respond(resetPicks(payload));
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
  return PropertiesService.getScriptProperties().getProperty('ADMIN_KEY') || 'Analysis';
}

function getConfig() {
  const prop    = PropertiesService.getScriptProperties();
  const birdies = prop.getProperty('ACTUAL_BIRDIES');
  return { actualBirdies: birdies ? Number(birdies) : null };
}

function resetPicks(payload) {
  if ((payload.adminKey || '') !== getAdminKey_()) throw new Error('Unauthorized');
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(PICKS_SHEET);
  const last  = sheet.getLastRow();
  if (last > 1) sheet.deleteRows(2, last - 1);
  return { status: 'ok', deleted: last - 1 };
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
  { rank:  1, name: 'Scottie Scheffler',          odds: '+510',    tier: 'A' },
  { rank:  2, name: 'Jon Rahm',                   odds: '+900',    tier: 'A' },
  { rank:  3, name: 'Bryson DeChambeau',           odds: '+1050',   tier: 'A' },
  { rank:  4, name: 'Rory McIlroy',               odds: '+1175',   tier: 'A' },
  { rank:  5, name: 'Ludvig Aberg',                odds: '+1650',   tier: 'A' },
  { rank:  6, name: 'Xander Schauffele',           odds: '+1750',   tier: 'A' },
  { rank:  7, name: 'Cameron Young',               odds: '+2200',   tier: 'A' },
  { rank:  8, name: 'Tommy Fleetwood',             odds: '+2250',   tier: 'A' },
  { rank:  9, name: 'Matt Fitzpatrick',            odds: '+2300',   tier: 'A' },
  { rank: 10, name: 'Hideki Matsuyama',            odds: '+2700',   tier: 'A' },
  // Tier B (11-20)
  { rank: 11, name: 'Collin Morikawa',             odds: '+3100',   tier: 'B' },
  { rank: 12, name: 'Min Woo Lee',                 odds: '+3300',   tier: 'B' },
  { rank: 13, name: 'Justin Rose',                 odds: '+3500',   tier: 'B' },
  { rank: 14, name: 'Robert MacIntyre',            odds: '+3500',   tier: 'B' },
  { rank: 15, name: 'Brooks Koepka',               odds: '+3700',   tier: 'B' },
  { rank: 16, name: 'Patrick Reed',                odds: '+4200',   tier: 'B' },
  { rank: 17, name: 'Jordan Spieth',               odds: '+4200',   tier: 'B' },
  { rank: 18, name: 'Chris Gotterup',              odds: '+4300',   tier: 'B' },
  { rank: 19, name: 'Viktor Hovland',              odds: '+4500',   tier: 'B' },
  { rank: 20, name: 'Si Woo Kim',                  odds: '+5000',   tier: 'B' },
  // Tier C (21-30)
  { rank: 21, name: 'Akshay Bhatia',               odds: '+5100',   tier: 'C' },
  { rank: 22, name: 'Russell Henley',              odds: '+5400',   tier: 'C' },
  { rank: 23, name: 'Justin Thomas',               odds: '+5900',   tier: 'C' },
  { rank: 24, name: 'Adam Scott',                  odds: '+6000',   tier: 'C' },
  { rank: 25, name: 'Patrick Cantlay',             odds: '+6400',   tier: 'C' },
  { rank: 26, name: 'Jake Knapp',                  odds: '+6400',   tier: 'C' },
  { rank: 27, name: 'Shane Lowry',                 odds: '+6600',   tier: 'C' },
  { rank: 28, name: 'Jason Day',                   odds: '+6800',   tier: 'C' },
  { rank: 29, name: 'J.J. Spaun',                  odds: '+6800',   tier: 'C' },
  { rank: 30, name: 'Sam Burns',                   odds: '+7000',   tier: 'C' },
  // Alternate (31-91)
  { rank: 31, name: 'Nicolai Hojgaard',            odds: '+7400',   tier: 'Alt' },
  { rank: 32, name: 'Sepp Straka',                 odds: '+7600',   tier: 'Alt' },
  { rank: 33, name: 'Maverick McNealy',            odds: '+7800',   tier: 'Alt' },
  { rank: 34, name: 'Tyrrell Hatton',              odds: '+8000',   tier: 'Alt' },
  { rank: 35, name: 'Jacob Bridgeman',             odds: '+8400',   tier: 'Alt' },
  { rank: 36, name: 'Corey Conners',               odds: '+8400',   tier: 'Alt' },
  { rank: 37, name: 'Kurt Kitayama',               odds: '+10000',  tier: 'Alt' },
  { rank: 38, name: 'Harris English',              odds: '+10000',  tier: 'Alt' },
  { rank: 39, name: 'Ben Griffin',                 odds: '+11000',  tier: 'Alt' },
  { rank: 40, name: 'Cameron Smith',               odds: '+11000',  tier: 'Alt' },
  { rank: 41, name: 'Sung-Jae Im',                 odds: '+11500',  tier: 'Alt' },
  { rank: 42, name: 'Gary Woodland',               odds: '+12000',  tier: 'Alt' },
  { rank: 43, name: 'Max Homa',                    odds: '+12000',  tier: 'Alt' },
  { rank: 44, name: 'Daniel Berger',               odds: '+12000',  tier: 'Alt' },
  { rank: 45, name: 'Rasmus Hojgaard',             odds: '+13000',  tier: 'Alt' },
  { rank: 46, name: 'Keegan Bradley',              odds: '+14000',  tier: 'Alt' },
  { rank: 47, name: 'Marco Penge',                 odds: '+14000',  tier: 'Alt' },
  { rank: 48, name: 'Harry Hall',                  odds: '+15000',  tier: 'Alt' },
  { rank: 49, name: 'Ryan Gerard',                 odds: '+15500',  tier: 'Alt' },
  { rank: 50, name: 'Alex Noren',                  odds: '+16000',  tier: 'Alt' },
  { rank: 51, name: 'Sam Stevens',                 odds: '+17000',  tier: 'Alt' },
  { rank: 52, name: 'Nick Taylor',                 odds: '+19000',  tier: 'Alt' },
  { rank: 53, name: 'Ryan Fox',                    odds: '+20000',  tier: 'Alt' },
  { rank: 54, name: 'Wyndham Clark',               odds: '+20000',  tier: 'Alt' },
  { rank: 55, name: 'Michael Kim',                 odds: '+21000',  tier: 'Alt' },
  { rank: 56, name: 'Max Greyserman',              odds: '+21000',  tier: 'Alt' },
  { rank: 57, name: 'Brian Harman',                odds: '+21000',  tier: 'Alt' },
  { rank: 58, name: 'Kristoffer Reitan',           odds: '+22000',  tier: 'Alt' },
  { rank: 59, name: 'Casey Jarvis',                odds: '+23000',  tier: 'Alt' },
  { rank: 60, name: 'Carlos Ortiz',                odds: '+23000',  tier: 'Alt' },
  { rank: 61, name: 'Sergio Garcia',               odds: '+25000',  tier: 'Alt' },
  { rank: 62, name: 'Dustin Johnson',              odds: '+25000',  tier: 'Alt' },
  { rank: 63, name: 'Aaron Rai',                   odds: '+25000',  tier: 'Alt' },
  { rank: 64, name: 'Haotong Li',                  odds: '+27000',  tier: 'Alt' },
  { rank: 65, name: 'Matt McCarty',                odds: '+28000',  tier: 'Alt' },
  { rank: 66, name: 'Andrew Novak',                odds: '+29000',  tier: 'Alt' },
  { rank: 67, name: 'Tom McKibbin',                odds: '+29000',  tier: 'Alt' },
  { rank: 68, name: 'Rasmus Neergaard-Petersen',   odds: '+31000',  tier: 'Alt' },
  { rank: 69, name: 'Nico Echavarria',             odds: '+32500',  tier: 'Alt' },
  { rank: 70, name: 'Sami Valimaki',               odds: '+36000',  tier: 'Alt' },
  { rank: 71, name: 'Aldrich Potgieter',           odds: '+36000',  tier: 'Alt' },
  { rank: 72, name: 'John Keefer',                 odds: '+36000',  tier: 'Alt' },
  { rank: 73, name: 'Michael Brennan',             odds: '+39000',  tier: 'Alt' },
  { rank: 74, name: 'Bubba Watson',                odds: '+52500',  tier: 'Alt' },
  { rank: 75, name: 'Zach Johnson',                odds: '+52500',  tier: 'Alt' },
  { rank: 76, name: 'Charl Schwartzel',            odds: '+62500',  tier: 'Alt' },
  { rank: 77, name: 'Davis Riley',                 odds: '+82500',  tier: 'Alt' },
  { rank: 78, name: 'Mason Howell',                odds: '+200000', tier: 'Alt' },
  { rank: 79, name: 'Danny Willett',               odds: '+225000', tier: 'Alt' },
  { rank: 80, name: 'Angel Cabrera',               odds: '+300000', tier: 'Alt' },
  { rank: 81, name: 'Brian Campbell',              odds: '+325000', tier: 'Alt' },
  { rank: 82, name: 'Ethan Fang',                  odds: '+325000', tier: 'Alt' },
  { rank: 83, name: 'Pongsapak Laopakdee',         odds: '+350000', tier: 'Alt' },
  { rank: 84, name: 'Naoyuki Kataoka',             odds: '+450000', tier: 'Alt' },
  { rank: 85, name: 'Brandon Holtz',               odds: '+500000', tier: 'Alt' },
  { rank: 86, name: 'Vijay Singh',                 odds: '+500000', tier: 'Alt' },
  { rank: 87, name: 'Mike Weir',                   odds: '+500000', tier: 'Alt' },
  { rank: 88, name: 'Fred Couples',                odds: '+500000', tier: 'Alt' },
  { rank: 89, name: 'Jose Maria Olazabal',         odds: '+500000', tier: 'Alt' },
  { rank: 90, name: 'Mateo Pulcini',               odds: '+500000', tier: 'Alt' },
  { rank: 91, name: 'Jackson Herrington',          odds: '+500000', tier: 'Alt' },
];

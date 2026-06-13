/**
 * ════════════════════════════════════════════════════════════════
 *  FinanceOS — Google Sheet builder
 * ════════════════════════════════════════════════════════════════
 *  HOW TO USE (one time):
 *    1. Open script.google.com → New project
 *    2. Delete any code there, paste this whole file, press Save
 *    3. In the function dropdown (top toolbar) pick  setupFinanceOS
 *    4. Press ▶ Run, authorise when asked (it's your own script)
 *    5. Open Google Drive — your new sheet "FinanceOS Data" is ready
 *
 *  The script creates 8 tabs, all formulas, named ranges, the
 *  💰 FinanceOS menu, the daily 7:30-8:30am briefing email and the
 *  "last updated" stamp. You never need to run it again.
 * ════════════════════════════════════════════════════════════════
 */

var SHEET_NAME = 'FinanceOS Data';

/* ──────────────────────────────────────────────
   MAIN — run this once
   ────────────────────────────────────────────── */
function setupFinanceOS() {
  var ss = SpreadsheetApp.create(SHEET_NAME);
  PropertiesService.getScriptProperties().setProperty('FINANCEOS_SS_ID', ss.getId());

  buildConfig_(ss);
  buildAccounts_(ss);
  buildEquityMF_(ss);
  buildGold_(ss);
  buildFixedIncome_(ss);
  buildLiabilities_(ss);   // includes the Card EMIs section
  buildSubscriptions_(ss);
  buildCoreExpenses_(ss);
  buildNetWorth_(ss);
  buildTax_(ss);

  // Remove the default empty "Sheet1"
  var s1 = ss.getSheetByName('Sheet1');
  if (s1) ss.deleteSheet(s1);

  installTriggers_(ss);

  Logger.log('✅ Done! Open Google Drive and look for the sheet named "' + SHEET_NAME + '"');
  Logger.log('Sheet URL: ' + ss.getUrl());
}

function getSS_() {
  var id = PropertiesService.getScriptProperties().getProperty('FINANCEOS_SS_ID');
  return SpreadsheetApp.openById(id);
}

/**
 * RECOVERY: if you accidentally ran setupFinanceOS (which makes a NEW sheet),
 * paste your ORIGINAL sheet ID below and run this to re-link the script to it,
 * then it adds the new tabs. Find the ID in your real sheet's URL between /d/ and /edit.
 * After running, delete the empty new sheet from Drive and redeploy the web app.
 */
function relinkToMySheetAndAddModules() {
  var MY_SHEET_ID = 'PASTE_YOUR_ORIGINAL_SHEET_ID_HERE';
  PropertiesService.getScriptProperties().setProperty('FINANCEOS_SS_ID', MY_SHEET_ID);
  addNewModules();
}

/**
 * Run this ONCE to add the new Predict-engine tabs to your EXISTING sheet
 * (Subscriptions, Card EMIs, Card Swipes) without rebuilding everything.
 * Safe to run again — it skips tabs that already exist.
 * After running, redeploy the web app: Deploy → Manage deployments → ✏️ → New version.
 */
function addNewModules() {
  var ss = getSS_();
  if (!ss.getSheetByName('SUBSCRIPTIONS')) buildSubscriptions_(ss);
  if (!ss.getSheetByName('CORE_EXPENSES')) buildCoreExpenses_(ss);
  addCardEmiSection_(ss);   // adds the EMI section into the existing LIABILITIES tab

  // Retire the old standalone tabs — EMIs now live in LIABILITIES, swipes are derived
  ['CARD_EMIS', 'CARD_SWIPES'].forEach(function (n) {
    var old = ss.getSheetByName(n);
    if (old) { try { ss.deleteSheet(old); } catch (e) {} }
  });

  // Fold subscription burn into Monthly Obligations (if not already)
  try {
    var nw = ss.getSheetByName('NET_WORTH');
    var f = nw.getRange('B11').getFormula();
    if (f && f.indexOf('SUBSCRIPTION_BURN') === -1) nw.getRange('B11').setFormula(f + '+SUBSCRIPTION_BURN');
  } catch (e) {}
  toast_('Modules updated ✓ EMIs now live in LIABILITIES. Redeploy the web app (Deploy → Manage deployments → New version).');
}

/* Adds the Card EMIs section (rows 22+) to an existing LIABILITIES tab if absent. */
function addCardEmiSection_(ss) {
  var sh = ss.getSheetByName('LIABILITIES');
  if (!sh) return;
  if (String(sh.getRange('A22').getValue()).toLowerCase().indexOf('card emi') !== -1) return; // already there
  sh.getRange('A22').setValue('Card EMIs').setFontWeight('bold').setFontSize(12)
    .setNote('EMIs running on your credit cards. Pick the card from the dropdown — it must match a card above. The app derives your "pay in full" swipe amount automatically.');
  sh.getRange('A23:H23').setValues([[
    'On Card', 'Description', 'Monthly Installment', 'Total Months', 'Months Paid',
    'Interest Rate %', 'Months Remaining', 'Balance Remaining'
  ]]).setFontWeight('bold').setBackground('#ffdad6');
  for (var er = 24; er <= 40; er++) {
    sh.getRange('G' + er).setFormula('=IF(C' + er + '="","",MAX(0,D' + er + '-E' + er + '))');
    sh.getRange('H' + er).setFormula('=IF(C' + er + '="","",C' + er + '*G' + er + ')');
  }
  sh.getRange('C24:C40').setNumberFormat('₹#,##,##0');
  sh.getRange('H24:H40').setNumberFormat('₹#,##,##0');
  sh.getRange('F24:F40').setNumberFormat('0.0%');
  var cardRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sh.getRange('A2:A7'), true).setAllowInvalid(false).build();
  sh.getRange('A24:A40').setDataValidation(cardRule);
}

function todayPlus_(days) {
  var d = new Date();
  d.setDate(d.getDate() + days);
  d.setHours(0, 0, 0, 0);
  return d;
}

/* ──────────────────────────────────────────────
   TAB: CONFIG
   ────────────────────────────────────────────── */
function buildConfig_(ss) {
  var sh = ss.insertSheet('CONFIG');
  sh.getRange('A1').setValue('FINANCEOS CONFIG').setFontWeight('bold').setFontSize(14);
  var rows = [
    ['Sheet Version', '1.0'],
    ['Owner', 'AB'],
    ['Monthly Take-Home Salary', ''],          // B4 — user fills
    ['Emergency Fund Multiplier', 6],          // B5
    ['Gemini API Key', ''],                    // B6 — user fills
    ['Last Refreshed', ''],                    // B7
    ['App Access Token', '']                   // B8 — generated from the menu (private mode)
  ];
  sh.getRange(2, 1, rows.length, 2).setValues(rows);
  sh.getRange('B7').setFormula('=NOW()');
  sh.getRange('B4').setBackground('#fff2cc').setNote('Fill in your monthly take-home salary (numbers only, e.g. 95000)');
  sh.getRange('B6').setBackground('#fff2cc').setNote('Paste your free Gemini API key from aistudio.google.com');
  sh.getRange('B8').setNote('Used by the FinanceOS app in private mode. Create it via the 💰 FinanceOS menu → 🔐 Generate App Token');
  sh.getRange('A2:A8').setFontWeight('bold');
  sh.setColumnWidth(1, 230).setColumnWidth(2, 260);

  ss.setNamedRange('MONTHLY_SALARY', sh.getRange('B4'));
  ss.setNamedRange('EF_MULTIPLIER', sh.getRange('B5'));
}

/* ──────────────────────────────────────────────
   TAB: ACCOUNTS
   ────────────────────────────────────────────── */
function buildAccounts_(ss) {
  var sh = ss.insertSheet('ACCOUNTS');
  var today = todayPlus_(0);
  sh.getRange('A1:E1').setValues([['Account Name', 'Bank', 'Type', 'Balance', 'Last Updated']])
    .setFontWeight('bold').setBackground('#d6e3ff');
  sh.getRange(2, 1, 3, 5).setValues([
    ['Salary Account', 'HDFC', 'Savings', 95000, today],
    ['Emergency FD', 'SBI', 'FD', 87000, today],
    ['Liquid MF', 'Zerodha', 'Liquid Fund', 0, today]
  ]);
  sh.getRange('D2:D100').setNumberFormat('₹#,##,##0');
  sh.getRange('E2:E100').setNumberFormat('yyyy-mm-dd');

  // Helper total cell → named range (named ranges must point at cells)
  sh.getRange('G1').setValue('TOTAL LIQUID CASH').setFontWeight('bold');
  sh.getRange('G2').setFormula('=SUM(D2:D100)').setNumberFormat('₹#,##,##0').setFontWeight('bold');
  ss.setNamedRange('LIQUID_CASH', sh.getRange('G2'));
  sh.setColumnWidths(1, 7, 150);
}

/* ──────────────────────────────────────────────
   TAB: EQUITY_MF
   ────────────────────────────────────────────── */
function buildEquityMF_(ss) {
  var sh = ss.insertSheet('EQUITY_MF');
  sh.getRange('A1:K1').setValues([[
    'Name', 'Type', 'Ticker', 'Units', 'Buy Price', 'Current Price',
    'Invested Amount', 'Current Value', 'Abs Return %', 'CAGR %', 'Investment Date'
  ]]).setFontWeight('bold').setBackground('#d6e3ff');

  // Two kinds of rows:
  //  · MFs/SIPs — units & prices don't apply (SIPs buy at a different NAV every
  //    month, and Google Finance has no Indian MF NAVs). Just type the two
  //    numbers your fund app shows: Invested Amount + Current Value.
  //  · Stocks/ETFs — Units + Buy Price entered once; Current Price is live via
  //    GOOGLEFINANCE and the rest calculates itself.
  var data = [
    ['Parag Parikh Flexi Cap', 'MF-SIP',      '—',         '', '',   '', 71672, 214300, '', '', new Date(2022, 3, 1)],
    ['Mirae Asset Large Cap',  'MF-Lumpsum',  '—',         '', '',   '', 88410, 98400,  '', '', new Date(2021, 7, 15)],
    ['NIFTYBEES',              'ETF',         'NIFTYBEES', 85, 210,  '', '',    '',     '', '', new Date(2023, 0, 10)],
    ['RELIANCE',               'Stock',       'RELIANCE',  12, 2480, '', '',    '',     '', '', new Date(2023, 5, 20)]
  ];
  sh.getRange(2, 1, data.length, 11).setValues(data);

  // MF/SIP rows: Invested + Current Value are typed in (yellow = yours to update)
  sh.getRange('G2:H3').setBackground('#fff2cc')
    .setNote('Type these straight from your fund app (Groww / Coin / Kuvera): total amount invested, and what it\'s worth right now. Update once a month.');

  // Stock/ETF rows: live price + auto-calculated amounts
  for (var r = 4; r <= 5; r++) {
    sh.getRange('F' + r).setFormula('=IFERROR(GOOGLEFINANCE(CONCATENATE("NSE:",C' + r + '),"price"),"—")');
    sh.getRange('G' + r).setFormula('=D' + r + '*E' + r);
    sh.getRange('H' + r).setFormula('=IFERROR(D' + r + '*F' + r + ',G' + r + ')');
  }
  // Returns work for both kinds — they only need Invested (G) and Current (H)
  for (var r2 = 2; r2 <= 5; r2++) {
    sh.getRange('I' + r2).setFormula('=IFERROR((H' + r2 + '-G' + r2 + ')/G' + r2 + ',0)');
    sh.getRange('J' + r2).setFormula('=IFERROR(((H' + r2 + '/G' + r2 + ')^(365/MAX(1,TODAY()-K' + r2 + ')))-1,0)');
  }
  sh.getRange('F1').setNote('Only used for stocks/ETFs (live from Google Finance). Leave blank for MFs/SIPs.');
  sh.getRange('G1').setNote('MFs/SIPs: type it. Stocks/ETFs: calculates itself.');
  sh.getRange('J1').setNote('For SIPs this CAGR is approximate — it treats the whole amount as invested on the Investment Date. "Abs Return %" is the exact number.');
  sh.getRange('G2:H100').setNumberFormat('₹#,##,##0');
  sh.getRange('I2:J100').setNumberFormat('0.0%');
  sh.getRange('K2:K100').setNumberFormat('yyyy-mm-dd');

  sh.getRange('M1').setValue('EQUITY TOTAL').setFontWeight('bold');
  sh.getRange('M2').setFormula('=SUM(H2:H100)').setNumberFormat('₹#,##,##0').setFontWeight('bold');
  ss.setNamedRange('EQUITY_TOTAL', sh.getRange('M2'));
  sh.setColumnWidths(1, 13, 140);
}

/* ──────────────────────────────────────────────
   TAB: GOLD
   ────────────────────────────────────────────── */
function buildGold_(ss) {
  var sh = ss.insertSheet('GOLD');
  sh.getRange('A1:I1').setValues([[
    'Type', 'Grams', 'Buy Rate/g', 'Buy Date', 'Invested',
    'Current Rate/g', 'Current Value', 'Return %', 'CAGR %'
  ]]).setFontWeight('bold').setBackground('#ffdea8');

  var data = [
    ['Physical Gold',  8,   6200, new Date(2021, 2, 15), 49600],
    ['Gold SIP Accum', 4.2, 6800, new Date(2022, 8, 1),  28560],
    ['SGB 2023',       2,   5900, new Date(2023, 6, 10), 11800]
  ];
  sh.getRange(2, 1, data.length, 5).setValues(data);

  // Shared live-rate cell (K2). All rows reference it.
  sh.getRange('K1').setValue('LIVE GOLD RATE/g').setFontWeight('bold');
  sh.getRange('K2').setFormula(
    '=IFERROR(VALUE(REGEXREPLACE(IMPORTXML("https://www.goodreturns.in/gold-rates-in-hyderabad.html",' +
    '"//span[contains(@class,\'gold-rate-per-gram\')][1]"),"[^0-9.]","")),7240)'
  ).setNumberFormat('₹#,##0');
  sh.getRange('K3').setValue('If formula fails, update Current Rate/g manually')
    .setFontStyle('italic').setFontColor('#7a5900');

  for (var r = 2; r <= 4; r++) {
    sh.getRange('F' + r).setFormula('=$K$2');
    sh.getRange('G' + r).setFormula('=B' + r + '*F' + r);
    sh.getRange('H' + r).setFormula('=IFERROR((G' + r + '-E' + r + ')/E' + r + ',0)');
    sh.getRange('I' + r).setFormula('=IFERROR(((G' + r + '/E' + r + ')^(365/MAX(1,TODAY()-D' + r + ')))-1,0)');
  }
  sh.getRange('C2:C100').setNumberFormat('₹#,##0');
  sh.getRange('E2:G100').setNumberFormat('₹#,##,##0');
  sh.getRange('H2:I100').setNumberFormat('0.0%');
  sh.getRange('D2:D100').setNumberFormat('yyyy-mm-dd');

  sh.getRange('K5').setValue('GOLD TOTAL').setFontWeight('bold');
  sh.getRange('K6').setFormula('=SUM(G2:G100)').setNumberFormat('₹#,##,##0').setFontWeight('bold');
  ss.setNamedRange('GOLD_TOTAL', sh.getRange('K6'));
  sh.setColumnWidths(1, 11, 130);
}

/* ──────────────────────────────────────────────
   TAB: FIXED_INCOME
   ────────────────────────────────────────────── */
function buildFixedIncome_(ss) {
  var sh = ss.insertSheet('FIXED_INCOME');
  sh.getRange('A1:K1').setValues([[
    'Type', 'Account', 'Principal', 'Rate %', 'Start Date', 'Maturity Date',
    'Current Balance', 'Projected Maturity', 'Days Left', 'Months Left', 'Status'
  ]]).setFontWeight('bold').setBackground('#b7f0d4');

  sh.getRange(2, 1, 2, 7).setValues([
    ['PPF', 'SBI PPF', 42000,    0.071, new Date(2015, 3, 1),  new Date(2030, 2, 31), 42000],
    ['RD',  'HDFC RD', '4200/mo', 0.068, new Date(2024, 9, 1),  new Date(2025, 7, 31), 46200]
  ]);

  for (var r = 2; r <= 3; r++) {
    sh.getRange('I' + r).setFormula('=IF(F' + r + '="","",F' + r + '-TODAY())');
    sh.getRange('J' + r).setFormula('=IF(F' + r + '="","",IF(F' + r + '<TODAY(),0,DATEDIF(TODAY(),F' + r + ',"M")))');
    // Projected maturity: compound the current balance monthly for the months left.
    // (For RDs add future deposits manually; for FDs/PPF this simple compounding is a fair projection.)
    sh.getRange('H' + r).setFormula('=IF(OR(G' + r + '="",J' + r + '=""),"",ROUND(G' + r + '*(1+D' + r + '/12)^MAX(0,J' + r + '),0))');
  }
  // Status: PPF is dormant (manual flag); RD uses the automatic formula
  sh.getRange('K2').setValue('DORMANT ⚠️').setFontColor('#7a5900').setFontWeight('bold')
    .setNote('DORMANT — reactivate with ₹500 minimum');
  sh.getRange('K3').setFormula('=IF(I3="","",IF(I3<0,"MATURED",IF(I3<30,"DUE SOON ⚠️","Active")))');

  sh.getRange('D2:D100').setNumberFormat('0.0%');
  sh.getRange('E2:F100').setNumberFormat('yyyy-mm-dd');
  sh.getRange('G2:H100').setNumberFormat('₹#,##,##0');

  sh.getRange('M1').setValue('FIXED INCOME TOTAL').setFontWeight('bold');
  sh.getRange('M2').setFormula('=SUM(G2:G100)').setNumberFormat('₹#,##,##0').setFontWeight('bold');
  ss.setNamedRange('FIXED_TOTAL', sh.getRange('M2'));
  sh.setColumnWidths(1, 13, 140);
}

/* ──────────────────────────────────────────────
   TAB: LIABILITIES  (cards + loans)
   ────────────────────────────────────────────── */
function buildLiabilities_(ss) {
  var sh = ss.insertSheet('LIABILITIES');

  // ── Section A: credit cards (rows 1-7)
  sh.getRange('A1:J1').setValues([[
    'Card', 'Network', 'Limit', 'Outstanding', 'Due Date',
    'Available', 'Utilisation %', 'Days to Due', 'Status', 'Last Paid'
  ]]).setFontWeight('bold').setBackground('#ffdad6');
  sh.getRange('J2:J7').setNumberFormat('yyyy-mm-dd');

  var cards = [
    ['ICICI Coral',      'Visa',       150000, 18000,  todayPlus_(14)],
    ['ICICI Amazon Pay', 'Visa',       80000,  24800,  todayPlus_(9)],
    ['Kotak',            'Mastercard', 200000, 16000,  todayPlus_(20)],
    ['Amex',             'Amex',       300000, 132000, todayPlus_(5)],
    ['Bajaj',            'RuPay',      100000, 71000,  todayPlus_(2)],
    ['Scapia',           'Visa',       50000,  0,      todayPlus_(25)]
  ];
  sh.getRange(2, 1, cards.length, 5).setValues(cards);
  for (var r = 2; r <= 7; r++) {
    sh.getRange('F' + r).setFormula('=C' + r + '-D' + r);
    sh.getRange('G' + r).setFormula('=IFERROR(D' + r + '/C' + r + ',0)');
    sh.getRange('H' + r).setFormula('=E' + r + '-TODAY()');
    sh.getRange('I' + r).setFormula(
      '=IF(H' + r + '<0,"OVERDUE ⛔",IF(H' + r + '<=3,"PAY NOW 🚨",IF(H' + r + '<=7,"DUE SOON ⚠️","OK ✓")))');
  }
  sh.getRange('C2:D7').setNumberFormat('₹#,##,##0');
  sh.getRange('F2:F7').setNumberFormat('₹#,##,##0');
  sh.getRange('G2:G7').setNumberFormat('0%');
  sh.getRange('E2:E7').setNumberFormat('yyyy-mm-dd');

  // ── Section B: loans (row 9 headers, row 10+ data) — blank row 8 separates
  sh.getRange('A9:L9').setValues([[
    'Loan', 'Lender', 'Principal', 'ROI %', 'EMI', 'Start Date', 'Tenure Months',
    'EMIs Paid', 'Remaining Principal', 'Months Left', 'Debt Free Date', 'Total Interest Left'
  ]]).setFontWeight('bold').setBackground('#eaddff');

  sh.getRange(10, 1, 1, 8).setValues([[
    'Personal Loan', 'IndMoney', 700000, 0.14, 12400, new Date(2024, 0, 15), 60, 17
  ]]);
  // Remaining principal = principal still owed after the EMIs already paid
  sh.getRange('I10').setFormula('=IFERROR(-CUMPRINC(D10/12,G10,C10,H10+1,G10,0),C10)');
  sh.getRange('J10').setFormula('=IFERROR(ROUND(NPER(D10/12,-E10,I10),0),G10-H10)');
  sh.getRange('K10').setFormula('=EDATE(TODAY(),J10)');
  sh.getRange('L10').setFormula('=E10*J10-I10');
  sh.getRange('C10:C20').setNumberFormat('₹#,##,##0');
  sh.getRange('E10:E20').setNumberFormat('₹#,##,##0');
  sh.getRange('I10:I20').setNumberFormat('₹#,##,##0');
  sh.getRange('L10:L20').setNumberFormat('₹#,##,##0');
  sh.getRange('D10:D20').setNumberFormat('0.0%');
  sh.getRange('F10:F20').setNumberFormat('yyyy-mm-dd');
  sh.getRange('K10:K20').setNumberFormat('mmm yyyy');

  // ── Section C: Card EMIs (rows 22+) — lives in THIS tab so it can never
  //    drift from the card list above. The "On Card" column is a dropdown
  //    locked to the cards in Section A.
  sh.getRange('A22').setValue('Card EMIs').setFontWeight('bold').setFontSize(12)
    .setNote('EMIs running on your credit cards. Pick the card from the dropdown — it must match a card above. The app derives your "pay in full" swipe amount as Outstanding − remaining EMI balance, so you never log swipes by hand.');
  sh.getRange('A23:H23').setValues([[
    'On Card', 'Description', 'Monthly Installment', 'Total Months', 'Months Paid',
    'Interest Rate %', 'Months Remaining', 'Balance Remaining'
  ]]).setFontWeight('bold').setBackground('#ffdad6');
  sh.getRange(24, 1, 2, 6).setValues([
    ['Amex',  'iPhone 15 Pro', 8500, 12, 3, 0.16],
    ['Bajaj', 'Laptop',        5200, 9,  2, 0.18]
  ]);
  for (var er = 24; er <= 40; er++) {
    sh.getRange('G' + er).setFormula('=IF(C' + er + '="","",MAX(0,D' + er + '-E' + er + '))');
    sh.getRange('H' + er).setFormula('=IF(C' + er + '="","",C' + er + '*G' + er + ')');
  }
  sh.getRange('C24:C40').setNumberFormat('₹#,##,##0');
  sh.getRange('H24:H40').setNumberFormat('₹#,##,##0');
  sh.getRange('F24:F40').setNumberFormat('0.0%');
  // Dropdown: "On Card" must be one of the cards in Section A
  var cardRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sh.getRange('A2:A7'), true).setAllowInvalid(false).build();
  sh.getRange('A24:A40').setDataValidation(cardRule);

  // ── Totals → named ranges
  sh.getRange('N1').setValue('TOTALS').setFontWeight('bold');
  sh.getRange('M2').setValue('Credit Card Outstanding');
  sh.getRange('N2').setFormula('=SUM(D2:D7)').setNumberFormat('₹#,##,##0');
  sh.getRange('M3').setValue('Loan Remaining');
  sh.getRange('N3').setFormula('=SUM(I10:I20)').setNumberFormat('₹#,##,##0');
  sh.getRange('M4').setValue('Total Debt').setFontWeight('bold');
  sh.getRange('N4').setFormula('=N2+N3').setNumberFormat('₹#,##,##0').setFontWeight('bold');
  ss.setNamedRange('CC_OUTSTANDING', sh.getRange('N2'));
  ss.setNamedRange('LOAN_REMAINING', sh.getRange('N3'));
  ss.setNamedRange('TOTAL_DEBT', sh.getRange('N4'));
  sh.setColumnWidths(1, 14, 135);
}

/* ──────────────────────────────────────────────
   TAB: SUBSCRIPTIONS  (recurring services)
   ────────────────────────────────────────────── */
function buildSubscriptions_(ss) {
  var sh = ss.insertSheet('SUBSCRIPTIONS');
  sh.getRange('A1:G1').setValues([[
    'Service', 'Amount', 'Cycle', 'Billing Day', 'Category', 'Monthly Cost', 'Notes'
  ]]).setFontWeight('bold').setBackground('#eaddff');
  sh.getRange('A1').setNote('This tab tracks ALL recurring auto-debits: subscriptions, insurance premiums, ' +
    'and SIPs. Use the Category column (Entertainment / Insurance / Investment / …). Annual plans are ' +
    'auto-converted to a monthly burn. They flow into your Upcoming Payments and Next-Month forecast.');
  // Cycle = Monthly or Annual; Billing Day = day of month (1-31)
  var data = [
    ['Netflix',            649,  'Monthly', 12, 'Entertainment', '', ''],
    ['ChatGPT Plus',       1700, 'Monthly', 20, 'Productivity',  '', ''],
    ['Amazon Prime',       1499, 'Annual',  8,  'Shopping',      '', ''],
    ['Medical Insurance',  2200, 'Monthly', 5,  'Insurance',     '', ''],
    ['Term Life Insurance',1800, 'Monthly', 14, 'Insurance',     '', ''],
    ['Parag Parikh SIP',   5000, 'Monthly', 1,  'Investment',    '', '']
  ];
  sh.getRange(2, 1, data.length, 7).setValues(data);
  for (var r = 2; r <= data.length + 1; r++) {
    // Monthly Cost normalises annual plans to a per-month burn
    sh.getRange('F' + r).setFormula('=IF(LOWER(C' + r + ')="annual",B' + r + '/12,B' + r + ')');
  }
  sh.getRange('B2:B100').setNumberFormat('₹#,##,##0');
  sh.getRange('F2:F100').setNumberFormat('₹#,##,##0');

  sh.getRange('I1').setValue('MONTHLY SUBSCRIPTION BURN').setFontWeight('bold');
  sh.getRange('I2').setFormula('=SUM(F2:F100)').setNumberFormat('₹#,##,##0').setFontWeight('bold');
  ss.setNamedRange('SUBSCRIPTION_BURN', sh.getRange('I2'));
  sh.setColumnWidths(1, 9, 130);
}

/* ──────────────────────────────────────────────
   TAB: CORE_EXPENSES  (monthly living baseline)
   Synced from the app's Predict tab; can also be edited here.
   ────────────────────────────────────────────── */
function buildCoreExpenses_(ss) {
  var sh = ss.insertSheet('CORE_EXPENSES');
  sh.getRange('A1:B1').setValues([['Item', 'Monthly Amount']])
    .setFontWeight('bold').setBackground('#b7f0d4');
  sh.getRange(2, 1, 4, 2).setValues([
    ['Rent & Housing', ''],
    ['Utilities (power, water, internet, mobile)', ''],
    ['Transport (fuel, transit, cabs)', ''],
    ['Groceries & Household', '']
  ]);
  sh.getRange('A6').setValue('TOTAL CORE FIXED / MONTH').setFontWeight('bold');
  sh.getRange('B6').setFormula('=SUM(B2:B5)').setFontWeight('bold');
  sh.getRange('B2:B6').setNumberFormat('₹#,##,##0');
  sh.getRange('B2:B5').setBackground('#fff2cc')
    .setNote('Set these in the FinanceOS app (Predict tab) — they sync here automatically. You can also type them directly.');
  ss.setNamedRange('CORE_FIXED_TOTAL', sh.getRange('B6'));
  sh.setColumnWidth(1, 330).setColumnWidth(2, 150);
}

/* Writes the four core-expense amounts (called by doGet?action=saveCore). */
function writeCoreExpenses_(params) {
  var ss = getSS_();
  var sh = ss.getSheetByName('CORE_EXPENSES');
  if (!sh) { buildCoreExpenses_(ss); sh = ss.getSheetByName('CORE_EXPENSES'); }
  var map = { rent: /rent|housing/i, utilities: /utilit/i, transport: /transport|fuel|commute/i, groceries: /grocer|household/i };
  var labels = sh.getRange('A2:A5').getValues();
  for (var i = 0; i < labels.length; i++) {
    var label = String(labels[i][0]);
    for (var key in map) {
      if (map[key].test(label) && params[key] !== undefined && params[key] !== '') {
        sh.getRange(2 + i, 2).setValue(Number(params[key]) || 0);
      }
    }
  }
}

/* Logs a credit-card bill payment (called by doGet?action=logPayment&card=&paid=).
   new outstanding = max(0, current − paid); advances each active EMI on that card by
   one month; rolls the due date forward ~1 month so it stops showing overdue. */
function writeBillPayment_(params) {
  var ss = getSS_();
  var sh = ss.getSheetByName('LIABILITIES');
  if (!sh) throw new Error('no LIABILITIES');
  var card = String(params.card || '').trim();
  var paid = Number(params.paid) || 0;
  if (!card) throw new Error('no card');

  // Find the card row in Section A (A2:A7)
  var cardNames = sh.getRange('A2:A7').getValues();
  var row = -1;
  for (var i = 0; i < cardNames.length; i++) {
    if (String(cardNames[i][0]).trim().toLowerCase() === card.toLowerCase()) { row = 2 + i; break; }
  }
  if (row === -1) throw new Error('card not found');

  var outstanding = Number(sh.getRange(row, 4).getValue()) || 0; // col D
  var newOut = Math.max(0, outstanding - paid);
  sh.getRange(row, 4).setValue(newOut);

  // Roll the due date forward one month (clears overdue/pay-now)
  var dueCell = sh.getRange(row, 5); // col E
  var due = dueCell.getValue();
  if (due instanceof Date) dueCell.setValue(new Date(due.getFullYear(), due.getMonth() + 1, due.getDate()));

  // Stamp the paid date (col J)
  sh.getRange(row, 10).setValue(new Date());

  // Advance each active EMI on this card by one month (Months Paid +1, capped)
  var emiRange = sh.getRange('A24:E40').getValues(); // A=On Card … E=Months Paid
  for (var j = 0; j < emiRange.length; j++) {
    if (String(emiRange[j][0]).trim().toLowerCase() !== card.toLowerCase()) continue;
    var totalM = Number(emiRange[j][3]) || 0;     // col D
    var paidM = Number(emiRange[j][4]) || 0;      // col E
    if (paidM < totalM) sh.getRange(24 + j, 5).setValue(paidM + 1);
  }
  return newOut;
}

/* ──────────────────────────────────────────────
   TAB: NET_WORTH  (all pulled from named ranges)
   ────────────────────────────────────────────── */
function buildNetWorth_(ss) {
  var sh = ss.insertSheet('NET_WORTH');
  sh.getRange('A1:C1').setValues([['Label', 'Value', 'Notes']])
    .setFontWeight('bold').setBackground('#d6e3ff');

  var rows = [
    ['Total Liquid Cash',       '=LIQUID_CASH',  ''],
    ['Total Equity & MF',       '=EQUITY_TOTAL', ''],
    ['Total Gold',              '=GOLD_TOTAL',   ''],
    ['Total Fixed Income',      '=FIXED_TOTAL',  ''],
    ['Total Assets',            '=SUM(B2:B5)',   ''],
    ['Credit Card Outstanding', '=CC_OUTSTANDING', ''],
    ['Loan Remaining',          '=LOAN_REMAINING', ''],
    ['Total Liabilities',       '=TOTAL_DEBT',   ''],
    ['NET WORTH',               '=B6-B9',        'Assets − Liabilities'],
    ['Monthly Obligations',     '=12400+18000+4200+SUBSCRIPTION_BURN', 'EMI + avg card spend + RD + subscriptions'],
    ['Emergency Fund Target',   '=B11*EF_MULTIPLIER', '6× monthly obligations'],
    ['EF Status',               '=IF(LIQUID_CASH>=B12,"SAFE ✅",IF(LIQUID_CASH>=B12*0.5,"VULNERABLE ⚠️","CRITICAL 🚨"))', '']
  ];
  for (var i = 0; i < rows.length; i++) {
    var r = i + 2;
    sh.getRange('A' + r).setValue(rows[i][0]);
    if (String(rows[i][1]).charAt(0) === '=') sh.getRange('B' + r).setFormula(rows[i][1]);
    else sh.getRange('B' + r).setValue(rows[i][1]);
    sh.getRange('C' + r).setValue(rows[i][2]);
  }
  sh.getRange('B2:B12').setNumberFormat('₹#,##,##0');
  sh.getRange('A10:B10').setFontWeight('bold').setFontSize(14);
  sh.getRange('A2:A13').setFontWeight('bold');

  ss.setNamedRange('NET_WORTH_TOTAL', sh.getRange('B10'));
  ss.setNamedRange('MONTHLY_OBLIGATIONS', sh.getRange('B11'));
  ss.setNamedRange('EF_TARGET', sh.getRange('B12'));
  sh.setColumnWidth(1, 220).setColumnWidth(2, 160).setColumnWidth(3, 260);
}

/* ──────────────────────────────────────────────
   TAB: TAX  — FY2025-26 old vs new regime engine
   ────────────────────────────────────────────── */
function buildTax_(ss) {
  var sh = ss.insertSheet('TAX');
  sh.getRange('A1').setValue('TAX OPTIMIZER · FY 2025-26').setFontWeight('bold').setFontSize(14);

  // ── Inputs (yellow = fill in your numbers)
  sh.getRange('A2').setValue('INPUTS — fill the yellow cells').setFontWeight('bold').setBackground('#fff2cc');
  var inputs = [
    ['Gross CTC',                 1320000],
    ['Basic Salary (40% default)', ''],     // formula below
    ['HRA Received',              264000],
    ['Rent Paid (yearly)',        240000],
    ['80C Investments',           150000],
    ['80D Health Premium',        25000],
    ['Section 24 Home Loan Interest', 0]
  ];
  sh.getRange(3, 1, inputs.length, 2).setValues(inputs);
  sh.getRange('B4').setFormula('=B3*0.4');
  sh.getRange('B3:B9').setBackground('#fff2cc').setNumberFormat('₹#,##,##0');
  sh.getRange('A3:A9').setFontWeight('bold');

  // ── Old regime
  sh.getRange('A11').setValue('OLD REGIME').setFontWeight('bold').setBackground('#ffdad6');
  sh.getRange('A12').setValue('Standard Deduction');
  sh.getRange('B12').setValue(50000);
  sh.getRange('A13').setValue('HRA Exemption');
  sh.getRange('B13').setFormula('=MAX(0,MIN(B4*0.5,B5,B6-B4*0.1))');
  sh.getRange('A14').setValue('Taxable Income');
  sh.getRange('B14').setFormula('=MAX(0,B3-B12-B13-MIN(B7,150000)-B8-MIN(B9,200000))');
  sh.getRange('A15').setValue('Tax (slabs)');
  sh.getRange('B15').setFormula(
    '=IF(B14<=250000,0,IF(B14<=500000,(B14-250000)*0.05,' +
    'IF(B14<=1000000,12500+(B14-500000)*0.2,112500+(B14-1000000)*0.3)))');
  sh.getRange('A16').setValue('87A Rebate (taxable ≤ ₹5L)');
  sh.getRange('B16').setFormula('=IF(B14<=500000,-MIN(B15,12500),0)');
  sh.getRange('A17').setValue('Cess 4%');
  sh.getRange('B17').setFormula('=(B15+B16)*0.04');
  sh.getRange('A18').setValue('OLD REGIME TOTAL TAX').setFontWeight('bold');
  sh.getRange('B18').setFormula('=ROUND(B15+B16+B17,0)').setFontWeight('bold').setFontColor('#ba1a1a');

  // ── New regime FY2025-26
  sh.getRange('A20').setValue('NEW REGIME (FY 2025-26)').setFontWeight('bold').setBackground('#b7f0d4');
  sh.getRange('A21').setValue('Standard Deduction');
  sh.getRange('B21').setValue(75000);
  sh.getRange('A22').setValue('Taxable Income (no other deductions)');
  sh.getRange('B22').setFormula('=MAX(0,B3-B21)');
  sh.getRange('A23').setValue('Tax (slabs 0/5/10/15/20/25/30)');
  sh.getRange('B23').setFormula(
    '=IF(B22<=400000,0,IF(B22<=800000,(B22-400000)*0.05,' +
    'IF(B22<=1200000,20000+(B22-800000)*0.1,' +
    'IF(B22<=1600000,60000+(B22-1200000)*0.15,' +
    'IF(B22<=2000000,120000+(B22-1600000)*0.2,' +
    'IF(B22<=2400000,200000+(B22-2000000)*0.25,300000+(B22-2400000)*0.3))))))');
  sh.getRange('A24').setValue('87A Rebate (income ≤ ₹12L → zero tax)');
  sh.getRange('B24').setFormula('=IF(B22<=1200000,-B23,0)');
  sh.getRange('A25').setValue('Cess 4%');
  sh.getRange('B25').setFormula('=(B23+B24)*0.04');
  sh.getRange('A26').setValue('NEW REGIME TOTAL TAX').setFontWeight('bold');
  sh.getRange('B26').setFormula('=ROUND(B23+B24+B25,0)').setFontWeight('bold').setFontColor('#1a7f4b');

  // ── Recommendation
  sh.getRange('A28').setValue('RECOMMENDATION').setFontWeight('bold');
  sh.getRange('B28').setFormula(
    '=IF(B26<B18,"✅ NEW REGIME — saves ₹"&TEXT(B18-B26,"##,##,##0"),' +
    '"✅ OLD REGIME — saves ₹"&TEXT(B26-B18,"##,##,##0"))')
    .setFontWeight('bold').setFontSize(12);

  sh.getRange('B12:B26').setNumberFormat('₹#,##,##0');
  sh.setColumnWidth(1, 290).setColumnWidth(2, 200);
}

/* ──────────────────────────────────────────────
   TRIGGERS  (installable — work from a standalone script)
   ────────────────────────────────────────────── */
function installTriggers_(ss) {
  // Clean out old triggers from previous runs
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });

  ScriptApp.newTrigger('onOpenTrigger').forSpreadsheet(ss).onOpen().create();
  ScriptApp.newTrigger('onEditTrigger').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('sendMorningBriefing').timeBased().everyDays(1).atHour(7).nearMinute(45).create();
}

/** Adds the custom menu when the sheet opens */
function onOpenTrigger(e) {
  SpreadsheetApp.getUi()
    .createMenu('💰 FinanceOS')
    .addItem('🔄 Refresh Live Rates', 'refreshLiveRates')
    .addItem('📧 Send Morning Briefing', 'sendMorningBriefing')
    .addItem('📊 Recalculate Net Worth', 'recalcNetWorth')
    .addItem('📸 Export Monthly Snapshot', 'exportMonthlySnapshot')
    .addItem('🔐 Generate App Token', 'generateAppToken')
    .addItem('➕ Add new data tabs', 'addNewModules')
    .addToUi();
}

/** Stamps CONFIG!B7 on every edit */
function onEditTrigger(e) {
  try {
    var ss = e.source;
    var cfg = ss.getSheetByName('CONFIG');
    if (!cfg) return;
    // Don't stamp when the user is editing the stamp cell itself
    if (e.range && e.range.getSheet().getName() === 'CONFIG' && e.range.getA1Notation() === 'B7') return;
    cfg.getRange('B7').setValue('Last updated: ' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm'));
  } catch (err) { /* never block user edits */ }
}

/* ──────────────────────────────────────────────
   MENU ACTIONS
   ────────────────────────────────────────────── */

/** Forces GOOGLEFINANCE / IMPORTXML to recalculate by re-setting their formulas */
function refreshLiveRates() {
  var ss = getSS_();
  var gold = ss.getSheetByName('GOLD');
  if (gold) {
    var f = gold.getRange('K2').getFormula();
    gold.getRange('K2').setFormula('');
    SpreadsheetApp.flush();
    gold.getRange('K2').setFormula(f);
  }
  var eq = ss.getSheetByName('EQUITY_MF');
  if (eq) {
    ['F4', 'F5'].forEach(function (a1) {
      var f2 = eq.getRange(a1).getFormula();
      if (f2) {
        eq.getRange(a1).setFormula('');
        SpreadsheetApp.flush();
        eq.getRange(a1).setFormula(f2);
      }
    });
  }
  SpreadsheetApp.flush();
  toast_('Live rates refreshed ✓');
}

/** Recalculates / refreshes the NET_WORTH tab */
function recalcNetWorth() {
  var ss = getSS_();
  SpreadsheetApp.flush();
  var cfg = ss.getSheetByName('CONFIG');
  if (cfg) cfg.getRange('B7').setValue('Last updated: ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm'));
  toast_('Net worth recalculated: ' + inr_(num_(ss, 'NET_WORTH_TOTAL')));
}

/** Copies NET_WORTH values into a Snapshots tab with today's date */
function exportMonthlySnapshot() {
  var ss = getSS_();
  var nw = ss.getSheetByName('NET_WORTH');
  var snaps = ss.getSheetByName('SNAPSHOTS');
  if (!snaps) {
    snaps = ss.insertSheet('SNAPSHOTS');
    snaps.getRange('A1:E1').setValues([['Date', 'Net Worth', 'Total Assets', 'Total Liabilities', 'Liquid Cash']])
      .setFontWeight('bold').setBackground('#eaddff');
    snaps.getRange('B2:E1000').setNumberFormat('₹#,##,##0');
    snaps.getRange('A2:A1000').setNumberFormat('yyyy-mm-dd');
  }
  var row = snaps.getLastRow() + 1;
  snaps.getRange(row, 1, 1, 5).setValues([[
    new Date(),
    nw.getRange('B10').getValue(),
    nw.getRange('B6').getValue(),
    nw.getRange('B9').getValue(),
    nw.getRange('B2').getValue()
  ]]);
  toast_('Snapshot saved ✓');
}

/* ──────────────────────────────────────────────
   MORNING BRIEFING EMAIL (daily, 7:30-8:30am)
   ────────────────────────────────────────────── */
var TIPS = [
  'Pay credit cards in full, always. Revolving at 36-42% APR is the most expensive money in India.',
  'Automate the SIP the day after salary lands — what you don\'t see, you don\'t spend.',
  'Keep card utilisation under 30% per card. It\'s the single fastest CIBIL booster.',
  'Review one subscription today. ₹499/month cancelled = ₹5,988/year invested.',
  'An emergency fund isn\'t an investment — it\'s insurance. Park it in a liquid fund or FD and forget returns.'
];

function sendMorningBriefing() {
  var ss = getSS_();
  var nw = ss.getSheetByName('NET_WORTH');
  var liab = ss.getSheetByName('LIABILITIES');
  var snaps = ss.getSheetByName('SNAPSHOTS');
  var email = Session.getEffectiveUser().getEmail();
  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'dd MMM yyyy');

  var netWorth = nw.getRange('B10').getValue();
  var efStatus = nw.getRange('B13').getValue();
  var liquid = nw.getRange('B2').getValue();
  var efTarget = nw.getRange('B12').getValue();

  // Delta vs last snapshot
  var deltaHtml = '';
  if (snaps && snaps.getLastRow() > 1) {
    var last = snaps.getRange(snaps.getLastRow(), 2).getValue();
    var delta = netWorth - last;
    var up = delta >= 0;
    deltaHtml = '<span style="color:' + (up ? '#1a7f4b' : '#ba1a1a') + ';font-weight:600">' +
      (up ? '▲ +' : '▼ −') + inr_(Math.abs(delta)).replace('₹', '₹') + ' since last snapshot</span>';
  } else {
    deltaHtml = '<span style="color:#888">Take your first snapshot (💰 FinanceOS menu) to start tracking the trend</span>';
  }

  // Payments due this week (cards rows 2-7 + loan EMI)
  var dueRows = '';
  var data = liab.getRange('A2:H7').getValues();
  data.forEach(function (r) {
    var name = r[0], outst = r[3], days = r[7];
    if (name && typeof days === 'number' && days >= 0 && days <= 7 && outst > 0) {
      var color = days <= 3 ? '#ba1a1a' : '#7a5900';
      dueRows += '<tr><td style="padding:8px 12px;border-bottom:1px solid #eee">' + name +
        '</td><td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;font-weight:600">' + inr_(outst) +
        '</td><td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;color:' + color + ';font-weight:600">' +
        days + ' day' + (days === 1 ? '' : 's') + '</td></tr>';
    }
  });
  if (!dueRows) dueRows = '<tr><td colspan="3" style="padding:12px;color:#1a7f4b;font-weight:600">Nothing due this week ✓</td></tr>';

  var tip = TIPS[new Date().getDate() % TIPS.length];
  var efPct = efTarget > 0 ? Math.round(liquid / efTarget * 100) : 0;

  var html =
    '<div style="font-family:\'Google Sans\',Roboto,Arial,sans-serif;max-width:520px;margin:0 auto;background:#f8f9ff;border-radius:24px;padding:28px">' +
    '<h2 style="margin:0 0 4px;color:#2d6bc6">FinanceOS</h2>' +
    '<p style="margin:0 0 22px;color:#888;font-size:13px">Your morning briefing · ' + today + '</p>' +

    '<div style="background:linear-gradient(140deg,#d6e3ff,#eaddff);border-radius:20px;padding:20px;margin-bottom:16px">' +
    '<div style="font-size:11px;letter-spacing:1.5px;color:#001945;opacity:.7">NET WORTH</div>' +
    '<div style="font-size:34px;color:#001945;margin:4px 0">' + inr_(netWorth) + '</div>' +
    deltaHtml + '</div>' +

    '<div style="background:#fff;border-radius:20px;padding:18px;margin-bottom:16px;border:1px solid #eee">' +
    '<div style="font-size:11px;letter-spacing:1.5px;color:#888;margin-bottom:8px">DUE THIS WEEK</div>' +
    '<table style="width:100%;border-collapse:collapse;font-size:14px">' + dueRows + '</table></div>' +

    '<div style="background:#fff;border-radius:20px;padding:18px;margin-bottom:16px;border:1px solid #eee">' +
    '<div style="font-size:11px;letter-spacing:1.5px;color:#888">EMERGENCY FUND</div>' +
    '<div style="font-size:16px;margin-top:6px">' + efStatus + ' · ' + efPct + '% of target (' + inr_(liquid) + ' of ' + inr_(efTarget) + ')</div></div>' +

    '<div style="background:#ffdea8;border-radius:20px;padding:18px;color:#7a5900">' +
    '<div style="font-size:11px;letter-spacing:1.5px;margin-bottom:6px">💡 TODAY\'S TIP</div>' +
    '<div style="font-size:14px;line-height:1.5">' + tip + '</div></div>' +

    '<p style="color:#bbb;font-size:11px;margin-top:20px;text-align:center">Sent automatically by your FinanceOS sheet</p></div>';

  MailApp.sendEmail({
    to: email,
    subject: 'FinanceOS · Your morning briefing · ' + today,
    htmlBody: html
  });
  toast_('Briefing emailed to ' + email + ' ✓');
}

/* ──────────────────────────────────────────────
   PRIVATE MODE — personal web-app API
   Deploy → New deployment → Web app → Execute as: Me →
   Who has access: Anyone. The sheet itself stays fully
   Restricted; data is only returned for the correct token.
   ────────────────────────────────────────────── */

/** Creates a long random token in CONFIG!B8 — copy it into the app's Settings */
function generateAppToken() {
  var ss = getSS_();
  var cfg = ss.getSheetByName('CONFIG');
  if (!cfg) { toast_('CONFIG tab not found'); return; }
  if (!cfg.getRange('A8').getValue()) {
    cfg.getRange('A8').setValue('App Access Token').setFontWeight('bold');
    cfg.getRange('B8').setNote('Used by the FinanceOS app in private mode — paste into ⚙️ Settings on your phone');
  }
  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  cfg.getRange('B8').setValue(token);
  toast_('Token created ✓ Copy it from CONFIG cell B8 into the app');
  Logger.log('App access token: ' + token);
  return token;
}

function getAppToken_() {
  try {
    var v = getSS_().getSheetByName('CONFIG').getRange('B8').getValue();
    return v ? String(v).trim() : '';
  } catch (e) { return ''; }
}

/** The app calls this URL. Returns all tabs as JSON (or JSONP via &callback=) */
function doGet(e) {
  var params = (e && e.parameter) || {};
  var token = (params.token || '').trim();
  var stored = getAppToken_();
  var payload;

  if (!stored) {
    payload = { error: 'no_token_set' };
  } else if (token !== stored) {
    payload = { error: 'bad_token' };
  } else if (params.action === 'saveCore') {
    // App is pushing the Core Fixed Expenses baseline to the sheet
    try { writeCoreExpenses_(params); payload = { ok: true }; }
    catch (err) { payload = { error: 'write_failed', detail: String(err) }; }
  } else if (params.action === 'logPayment') {
    // App is logging a credit-card bill payment
    try { var no = writeBillPayment_(params); payload = { ok: true, outstanding: no }; }
    catch (err) { payload = { error: 'write_failed', detail: String(err) }; }
  } else {
    var ss = getSS_();
    var tabs = {};
    ['CONFIG', 'ACCOUNTS', 'EQUITY_MF', 'GOLD', 'FIXED_INCOME', 'LIABILITIES',
     'SUBSCRIPTIONS', 'CORE_EXPENSES', 'NET_WORTH', 'TAX']
      .forEach(function (name) {
        var sh = ss.getSheetByName(name);
        if (!sh) return;
        var rows = sh.getDataRange().getDisplayValues();
        // Never send secrets back out, even to the token holder
        if (name === 'CONFIG') {
          rows = rows.map(function (r) {
            if (String(r[0]).indexOf('Gemini API Key') !== -1 ||
                String(r[0]).indexOf('App Access Token') !== -1) {
              var c = r.slice(); c[1] = '•••'; return c;
            }
            return r;
          });
        }
        tabs[name] = rows;
      });
    payload = { tabs: tabs, at: new Date().toISOString() };
  }

  var json = JSON.stringify(payload);
  var cb = params.callback;
  if (cb && /^[\w.]+$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

/* ──────────────────────────────────────────────
   Small helpers
   ────────────────────────────────────────────── */
function num_(ss, namedRange) {
  try { return ss.getRangeByName(namedRange).getValue() || 0; } catch (e) { return 0; }
}
function inr_(n) {
  n = Math.round(Number(n) || 0);
  var neg = n < 0; n = Math.abs(n);
  var s = String(n);
  if (s.length > 3) {
    var last3 = s.slice(-3);
    var rest = s.slice(0, -3).replace(/\B(?=(\d{2})+(?!\d))/g, ',');
    s = rest + ',' + last3;
  }
  return (neg ? '−₹' : '₹') + s;
}
function toast_(msg) {
  try { getSS_().toast(msg, 'FinanceOS', 5); } catch (e) { Logger.log(msg); }
}

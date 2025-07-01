/*════════════════════  CONFIG (safe to edit)  ════════════════════*/
var CURR_SHEET_NAME = 'sheet1'
var PEOPLE = ['Harsh', 'Sourav', 'Gautam', 'Tirth', 'Rahul']

/*================================================================*/

const NUM_PPL      = PEOPLE.length;
const SPARE_ROWS   = 3;
const HEADER_COLOR = '#eeeeee';
const LOCKED_BG    = '#f2f2f2';
const CUR_FMT      = '$#,##0.00';
const ST_BG    = '#fdff32';
const GT_BG    = '#abff32';
const COL_WIDTH    = 75;               // equal width for every column

const SUMMARY_LABEL = ['SubTotal', 'Tax $', 'Grand Total'];

function createGroceryTracker() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CURR_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CURR_SHEET_NAME);
  sheet.clearContents(); sheet.clearFormats();

  /*── indices & columns ─────────────────────────────────────────*/
  const HDR1 = SPARE_ROWS + 2;              // first header row
  const HDR2 = HDR1 + 1;                    // second header row
  const DATA = HDR2 + 1;                    // first data row
  const [SUB, TAX, GRD] = [2, 3, 4];

  const PRICE_COL  = 2;                     // column B
  const chkStart   = PRICE_COL + 1;         // first checkbox col
  const contStart  = chkStart + NUM_PPL;    // first contribution col
  const totalCols  = contStart + NUM_PPL - 1;

  /*── two-row header with merged banners ────────────────────────*/
  const top = Array(totalCols).fill('');
  top[0] = 'Item'; top[1] = 'Price';
  top[chkStart  - 1] = 'Members';
  top[contStart - 1] = 'Contributions';

  const sub = Array(totalCols).fill('');
  PEOPLE.forEach((n, i) => {
    sub[chkStart  - 1 + i] = n;
    sub[contStart - 1 + i] = n;
  });

  sheet.getRange(HDR1, 1, 1, totalCols).setValues([top]);
  sheet.getRange(HDR2, 1, 1, totalCols).setValues([sub]);

  sheet.getRange(HDR1, 1, 2, totalCols)
       .setFontWeight('bold')
       .setBackground(HEADER_COLOR)
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle')
       .setBorder(true, true, true, true, true, true);

  sheet.getRange(HDR1, 1, 2, 1).merge();                   // Item
  sheet.getRange(HDR1, 2, 2, 1).merge();                   // Price
  sheet.getRange(HDR1, chkStart , 1, NUM_PPL).merge();     // Consumptions
  sheet.getRange(HDR1, contStart, 1, NUM_PPL).merge();     // Contributions

  /*── summary block rows 2–4 ────────────────────────────────────*/
  SUMMARY_LABEL.forEach((lab, i) =>
    sheet.getRange(SUB + i, 1).setValue(lab).setFontWeight('bold')
  );

  const priceL = colLtr(PRICE_COL);

  sheet.getRange(TAX, PRICE_COL)
       .setValue(0)
       .setNumberFormat(CUR_FMT)
       .setHorizontalAlignment('center');

  sheet.getRange(SUB, PRICE_COL)
       .setFormula(`=SUM(${priceL}${DATA}:${priceL})`)
       .setNumberFormat(CUR_FMT)
       .setHorizontalAlignment('center');

  sheet.getRange(GRD, PRICE_COL)
       .setFormula(`=${priceL}${SUB}+${priceL}${TAX}`)
       .setNumberFormat(CUR_FMT)
       .setHorizontalAlignment('center');

  /*── per-person summary cells ─────────────────────────────────*/
  PEOPLE.forEach((_, i) => {
    const col = contStart + i, L = colLtr(col);

    sheet.getRange(SUB, col)
         .setFormula(`=SUM(${L}${DATA}:${L})`)
         .setNumberFormat(CUR_FMT)
         .setHorizontalAlignment('center');

    sheet.getRange(GRD, col)
         .setFormula(
           `=IFERROR(${L}${SUB}+` +
           `${L}${SUB}/${priceL}${SUB}*${priceL}${TAX},0)`
         )
         .setNumberFormat(CUR_FMT)
         .setHorizontalAlignment('center')
         .setFontWeight('bold');
  });

  /*── check-boxes + contribution formulas ──────────────────────*/
  const maxRows = sheet.getMaxRows();
  sheet.getRange(DATA, chkStart, maxRows - DATA + 1, NUM_PPL).insertCheckboxes();

  const denom = PEOPLE.map((_, j) =>
    `N(${colLtr(chkStart + j)}${DATA}:${colLtr(chkStart + j)})`).join('+');

  PEOPLE.forEach((_, i) => {
    const cbCol = chkStart + i, ctCol = contStart + i, cbL = colLtr(cbCol);

    sheet.getRange(DATA, ctCol).setFormula(
      `=ARRAYFORMULA(IF(${cbL}${DATA}:${cbL}=TRUE,` +
      `${priceL}${DATA}:${priceL}/(${denom}),0))`
    );

    sheet.getRange(DATA, ctCol, maxRows - DATA + 1, 1)
         .setNumberFormat(CUR_FMT)
         .setHorizontalAlignment('center');
  });

  /*── price column format ──────────────────────────────────────*/
  sheet.getRange(DATA, PRICE_COL, maxRows - DATA + 1, 1)
       .setNumberFormat(CUR_FMT)
       .setHorizontalAlignment('center');

  /*── grey out non-editable cells ───────────────────────────────*/
  sheet.getRange(SUB, 1, 1, totalCols).setBackground(ST_BG); // SubTotal row
  sheet.getRange(GRD, 1, 1, totalCols).setBackground(GT_BG); // GrandTotal row
  sheet.getRange(DATA, contStart, maxRows - DATA + 1, NUM_PPL)
       .setBackground(LOCKED_BG);

  /*── uniform column width + alignment ─────────────────────────*/
  sheet.setColumnWidths(1, totalCols, COL_WIDTH);
  sheet.getRange(1, 1, maxRows, totalCols)
       .setHorizontalAlignment('center');

  /*── freeze two header rows (no filter to avoid error) ─────────*/
  sheet.setFrozenRows(HDR2);
}


/* helper: column index → letter */
function colLtr(n) {
  let s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - 1) / 26 | 0; }
  return s;
}

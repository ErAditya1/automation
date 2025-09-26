// fillLoanDisbursementFromExcel.js
// Node.js + Puppeteer + xlsx
//
// Usage example at the bottom.

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

/**
 * Read excel file and return array of row objects (header -> value)
 * @param {string} filePath
 */
function readExcelRows(filePath) {
  const wb = XLSX.readFile(filePath);
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: '' }); // array of objects
  return rows;
}

/**
 * Default mapping between friendly Excel header and actual form field id/name.
 * If your Excel header matches the element id exactly, you can use an empty mapping {}
 * or pass your own mapping to the main function.
 *
 * Example keys: Admissionno, ValueDate, Product, AccountNo, ActivityType, VoucherType, ChequeNo,
 * ChequeDate, VoucherNo, ContraProduct, ContraAccountNo, SocietyVoucherNo, Narration, TotalAmount,
 * MoreDisbursementDetaos_ApplicationNo, MoreDisbursementDetaos_LoanNo, MoreDisbursementDetaos_DueDate, ...
 */
const DEFAULT_FIELD_MAP = {
  Admissionno: 'Admissionno',
  AdmissionNoHidden: 'AdmissionNoPkey', // example if you export hidden keys
  ValueDate: 'ValueDate',
  Product: 'Product',
  AccountNo: 'AccountNo',
  TempAccountNo: 'TempAccountNo',
  ActivityType: 'ActivityType',
  VoucherType: 'VoucherType',
  ChequeNo: 'ChequeNo',
  ChequeDate: 'ChequeDate',
  VoucherNo: 'VoucherNo',
  ContraProduct: 'ContraProduct',
  ContraAccountNo: 'ContraAccountNo',
  SocietyVoucherNo: 'SocietyVoucherNo',
  Narration: 'Narration',
  TotalAmount: 'TotalAmount',
  AmountInWords: 'AmountInWords',
  // MoreDisbursement details (modal fields / inputs)
  MoreDisbursementDetaos_ApplicationNo: 'MoreDisbursementDetaos_ApplicationNo',
  MoreDisbursementDetaos_AdmissionNoPkey: 'MoreDisbursementDetaos_AdmissionNoPkey',
  MoreDisbursementDetaos_ProductSlNo: 'MoreDisbursementDetaos_ProductSlNo',
  MoreDisbursementDetaos_LoanNo: 'MoreDisbursementDetaos_LoanNo',
  MoreDisbursementDetaos_OldLoanNo: 'MoreDisbursementDetaos_OldLoanNo',
  MoreDisbursementDetaos_DPNDate: 'MoreDisbursementDetaos_DPNDate',
  MoreDisbursementDetaos_DPNNo: 'MoreDisbursementDetaos_DPNNo',
  MoreDisbursementDetaos_DueDate: 'MoreDisbursementDetaos_DueDate',
  MoreDisbursementDetaos_DateOfAdvice: 'MoreDisbursementDetaos_DateOfAdvice',
  MoreDisbursementDetaos_DebitSlipNo: 'MoreDisbursementDetaos_DebitSlipNo',
  MoreDisbursementDetaos_LedgerFolioNo: 'MoreDisbursementDetaos_LedgerFolioNo',
  MoreDisbursementDetaos_DCCBSBNo: 'MoreDisbursementDetaos_DCCBSBNo',
  MoreDisbursementDetaos_PolicyName: 'MoreDisbursementDetaos_PolicyName',
  MoreDisbursementDetaos_ROI: 'MoreDisbursementDetaos_ROI',
  MoreDisbursementDetaos_PenalROI: 'MoreDisbursementDetaos_PenalROI',
  MoreDisbursementDetaos_IOAROI: 'MoreDisbursementDetaos_IOAROI',
  MoreDisbursementDetaos_DisbursmentAmount: 'MoreDisbursementDetaos_DisbursmentAmount',
  // checkboxes
  IdTransfer: 'IdTransfer', // checkbox (IsTransfer)
  IdCash: 'IdCash',
  IdCheck: 'IdCheck'
};

/**
 * Helper: wait for selector and click (safe)
 */
async function safeClick(page, selector, opts = {}) {
  try {
    await page.waitForSelector(selector, { timeout: opts.timeout || 5000 });
    await page.evaluate((sel) => {
      const el = document.querySelector(sel);
      if (el) el.scrollIntoView({ block: 'center', inline: 'center' });
    }, selector);
    await page.click(selector, { delay: opts.delay || 50 });
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * Helper: set input or textarea value
 */
async function setInputValue(page, selector, value) {
  if (value === undefined || value === null) return false;
  value = String(value);
  try {
    await page.waitForSelector(selector, { timeout: 4000 });
    // some fields are readonly; remove readonly attribute temporarily if present
    await page.evaluate((sel, val) => {
      const el = document.querySelector(sel);
      if (!el) return;
      // if it's a select, set option instead
      if (el.tagName.toLowerCase() === 'select') return;
      el.focus();
      if (el.readOnly) el.removeAttribute('readonly');
      if (el.type === 'checkbox' || el.type === 'radio') {
        el.checked = val === 'true' || val === '1' || val === true;
        el.dispatchEvent(new Event('change', { bubbles: true }));
        return;
      }
      el.value = val;
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
    }, selector, value);
    // type to trigger UI scripts (only if not too long)
    await page.focus(selector).catch(() => {});
    // short pause
    await page.waitForTimeout(120);
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * Helper: set/select a select element by value or visible text
 */
async function selectOption(page, selector, value) {
  if (value === undefined || value === null) return false;
  value = String(value);
  try {
    await page.waitForSelector(selector, { timeout: 4000 });
    // try selecting by value first
    const selected = await page.select(selector, value).catch(() => []);
    if (selected && selected.length) {
      // trigger change
      await page.evaluate((sel) => {
        const el = document.querySelector(sel);
        if (el) el.dispatchEvent(new Event('change', { bubbles: true }));
      }, selector);
      await page.waitForTimeout(150);
      return true;
    }
    // otherwise attempt to match option by visible text
    const didByText = await page.evaluate(
      ({ sel, text }) => {
        const el = document.querySelector(sel);
        if (!el) return false;
        const options = Array.from(el.options || []);
        const opt = options.find(o => (o.text || '').trim().toLowerCase() === (text || '').trim().toLowerCase());
        if (opt) {
          el.value = opt.value;
          el.dispatchEvent(new Event('change', { bubbles: true }));
          return true;
        }
        // try contains
        const opt2 = options.find(o => (o.text || '').toLowerCase().includes((text || '').toLowerCase()));
        if (opt2) {
          el.value = opt2.value;
          el.dispatchEvent(new Event('change', { bubbles: true }));
          return true;
        }
        return false;
      },
      { sel: selector, text: value }
    );
    await page.waitForTimeout(100);
    return didByText;
  } catch (e) {
    return false;
  }
}

/**
 * Set checkbox true/false
 */
async function setCheckbox(page, selector, truthy) {
  try {
    await page.waitForSelector(selector, { timeout: 3000 });
    await page.evaluate((sel, val) => {
      const el = document.querySelector(sel);
      if (!el) return;
      const should = !!(val === true || val === 'true' || val === 1 || val === '1' || val === 'on');
      if (el.checked !== should) {
        el.checked = should;
        el.dispatchEvent(new Event('change', { bubbles: true }));
      }
    }, selector, truthy);
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * Main: fill the form for each excel row
 * @param {puppeteer.Page} page - Puppeteer page already open on the form URL
 * @param {string} excelPath - path to xlsx file
 * @param {object} opts - optional object
 *    opts.fieldMap: { excelHeader: formFieldId } // default DEFAULT_FIELD_MAP
 *    opts.submitAfterRow: boolean - whether to click "Save" (btnSave) after filling each row (default false)
 *    opts.clickPrepareInsteadOfPost: boolean - to click hidden prepare input (btnPost) or to click visible Save (btnSave) (default false)
 *    opts.redirectUrlAfterSubmit: string (optional) - if set, page.goto(redirectUrl) after submission
 */
async function fillLoanDisbursementFromExcel(page, excelPath, opts = {}) {
  const fieldMap = opts.fieldMap || DEFAULT_FIELD_MAP;
  const submitAfterRow = !!opts.submitAfterRow;
  const clickPrepareInsteadOfPost = !!opts.clickPrepareInsteadOfPost;
  const redirectUrl = opts.redirectUrlAfterSubmit || null;

  if (!fs.existsSync(excelPath)) throw new Error('Excel file not found: ' + excelPath);
  const rows = readExcelRows(excelPath);
  if (!rows || rows.length === 0) {
    return { ok: false, message: 'No rows in excel' };
  }

  const results = [];

  // Helper to map excel header -> field id
  const mapHeaderToId = header => {
    if (!header) return null;
    // exact match to mapping keys (case-sensitive)
    if (fieldMap[header]) return fieldMap[header];
    // if header already matches an id present in DOM, return header
    return header;
  };

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const rowResult = { rowIndex: i, filled: [], errors: [] };
    try {
      // For safety, wait that page still has formid = 40004 (if present)
      await page.waitForTimeout(200); // small throttle

      // Iterate fields in row
      for (const colHeader of Object.keys(row)) {
        const rawVal = row[colHeader];
        if (rawVal === '' || rawVal === null || rawVal === undefined) continue;
        const fieldId = mapHeaderToId(colHeader);
        if (!fieldId) {
          rowResult.errors.push(`No mapping for Excel header "${colHeader}"`);
          continue;
        }
        const selectorById = `#${fieldId}`;
        // prefer selects when element is select
        const elementType = await page.evaluate((sel) => {
          const el = document.querySelector(sel);
          if (!el) return null;
          return el.tagName.toLowerCase();
        }, selectorById).catch(() => null);

        const val = rawVal;

        if (elementType === 'select') {
          const ok = await selectOption(page, selectorById, val);
          if (ok) rowResult.filled.push(fieldId);
          else rowResult.errors.push(`Could not select ${fieldId} -> ${val}`);
        } else if (elementType === 'input' || elementType === 'textarea' || elementType === 'null') {
          // handle checkboxes specially if id corresponds to known checkbox names
          if (fieldId.toLowerCase().includes('is') || fieldId.toLowerCase().includes('id') || fieldId.toLowerCase().includes('chk')) {
            // don't assume — check element type
            const tag = await page.evaluate((sel) => {
              const el = document.querySelector(sel);
              if (!el) return null;
              return el.type || el.tagName.toLowerCase();
            }, selectorById).catch(() => null);

            if (tag === 'checkbox') {
              const ok = await setCheckbox(page, selectorById, val);
              if (ok) rowResult.filled.push(fieldId);
              else rowResult.errors.push(`Could not set checkbox ${fieldId}`);
              continue;
            }
          }
          // default: set input value
          const ok = await setInputValue(page, selectorById, val);
          if (ok) rowResult.filled.push(fieldId);
          else {
            // maybe element id doesn't exist — try name selector (some inputs use name attr)
            const nameSel = `[name="${fieldId}"]`;
            const ok2 = await setInputValue(page, nameSel, val);
            if (ok2) rowResult.filled.push(`${fieldId}(byName)`);
            else {
              rowResult.errors.push(`Could not set value for ${fieldId} (tried id & name)`);
            }
          }
        } else {
          // unknown element, fallback: try id input value set
          const ok = await setInputValue(page, `#${fieldId}`, val);
          if (ok) rowResult.filled.push(fieldId);
          else rowResult.errors.push(`Unknown element type or not found for ${fieldId}`);
        }

        // small delay between fields to let page JS run
        await page.waitForTimeout(120);
      } // end fields loop

      // Extra: if MoreDisbursementDetaos_DisbursmentAmount exists use it to set TotalAmount if not present
      const maybeDisbAmount = row['MoreDisbursementDetaos_DisbursmentAmount'] || row['MoreDisbursementDetaos_DisbursmentAmount'.toLowerCase()];
      if (maybeDisbAmount && !(row['TotalAmount'] || row['TotalAmount'.toLowerCase()])) {
        // try set TotalAmount and AmountInWords
        await setInputValue(page, '#TotalAmount', maybeDisbAmount).catch(() => {});
        // not calculating words; server may compute AmountInWords. leave it.
      }

      // Optionally open More Disbursement modal and click Save there if fields set
      // There is a "SaveDisbursementDetails" button which is triggered by clicking #btnsaveDisbursments
      // If the row includes MoreDisbursementDetaos_... fields we'll try to click SaveDisbursementDetails
      const hasMoreDisbFields = Object.keys(row).some(h => h.startsWith('MoreDisbursementDetaos'));
      if (hasMoreDisbFields) {
        // ensure modal is visible (Some inputs are not inside modal in this HTML - they are present as inputs)
        // Try to click SaveDisbursementDetails button if present
        const savedisb = await page.$('#btnsaveDisbursments');
        if (savedisb) {
          // call the SaveDisbursementDetails JS method by clicking the button
          await page.evaluate(() => {
            const el = document.querySelector('#btnsaveDisbursments');
            if (el) el.click();
          });
          await page.waitForTimeout(600); // allow internal processing
        }
      }

      // Submit (Prepare/Save) if requested
      if (submitAfterRow) {
        // two ways: click visible Save (btnSave) OR click hidden submit input btnPost or btnSaveDisTemp
        if (clickPrepareInsteadOfPost) {
          // Try hidden post input
          const btnPostExists = await page.$('#btnPost');
          if (btnPostExists) {
            await page.evaluate(() => {
              const btn = document.querySelector('#btnPost');
              if (btn) btn.click();
            }).catch(() => {});
          } else {
            // fallback click Save button
            await safeClick(page, '#btnSave');
          }
        } else {
          // click Save visible button (calls resetPostFields())
          const ok = await safeClick(page, '#btnSave');
          if (!ok) {
            // fallback to form submit
            await page.evaluate(() => {
              const forms = document.querySelectorAll('form[action*="TransactionPayment"]');
              if (forms && forms[0]) forms[0].submit();
            }).catch(() => {});
          }
        }
        // wait for possible navigation or result
        await page.waitForTimeout(1000);
        if (redirectUrl) {
          await page.goto(redirectUrl, { waitUntil: 'domcontentloaded', timeout: 15000 }).catch(() => {});
          await page.waitForTimeout(600);
        }
      }

      rowResult.ok = true;
    } catch (err) {
      rowResult.ok = false;
      rowResult.exception = String(err && err.message ? err.message : err);
    }
    results.push(rowResult);
    // small delay between rows
    await page.waitForTimeout(500);
  } // end rows

  return { ok: true, rows: rows.length, results };
}

// Example usage (uncomment and run from a script that launches puppeteer and signs-in):
/*
const puppeteer = require('puppeteer');

(async () => {
  const browser = await puppeteer.launch({ headless: false, defaultViewport: null });
  const page = await browser.newPage();

  // navigate and authenticate as your automation does...
  await page.goto('https://up6.uniteerp.in', { waitUntil: 'networkidle2' });

  // After login, navigate to Loans -> Loan Disbursement (or call your helper)
  // await openLoansAndOpenLoanDisbursement(page); // your earlier helper

  // Wait for form to appear (formid hidden input)
  await page.waitForSelector('#formid');

  // Fill from excel
  const excelPath = path.resolve(__dirname, 'loan_disbursements.xlsx');
  const result = await fillLoanDisbursementFromExcel(page, excelPath, {
    fieldMap: DEFAULT_FIELD_MAP,
    submitAfterRow: false, // set true to auto-click Save per row
    clickPrepareInsteadOfPost: false,
    redirectUrlAfterSubmit: null
  });

  console.log(JSON.stringify(result, null, 2));

  // await browser.close();
})();
*/



module.exports = { fillLoanDisbursementFromExcel, DEFAULT_FIELD_MAP };

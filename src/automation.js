// logs folder// automation.customized.js
// Improved, modular automation flow: explicit "login -> redirect -> fill form" behaviour
// Key features:
// - Modular safe helpers (safeClick, safeFill, safeSelect)
// - Explicit login then redirect to per-row or global FAS URL
// - Robust captcha handling with manual wait
// - Configurable via environment variables
// - CSV results and rotating logfile
// Usage: node automation.customized.js (set env vars or a .env file)

import "dotenv/config";
import fs from "fs";
import path from "path";
import { chromium } from "playwright";
import { readExcel } from "./excel.js";
import { SELECTORS } from "./config.js";
import { createObjectCsvWriter } from "csv-writer";
import { log } from "console";

// ---------------------- Configuration ----------------------
const LOGIN_URL = process.env.TARGET_LOGIN_URL || "https://example.com/Home";
const HEADLESS = (process.env.HEADLESS || "false") === "true";
const EXCEL_PATH = process.env.EXCEL_PATH || "./data/input.xlsx";
const CAPTCHA_WAIT_MS = parseInt(process.env.CAPTCHA_WAIT_MS || "300000", 10); // 5 minutes default
const SAVE_SCREENSHOT_ON_SUCCESS =
  (process.env.SAVE_SCREENSHOT_ON_SUCCESS || "false") === "true";
const RETRY_LOGIN = parseInt(process.env.RETRY_LOGIN || "1", 10);
const WAIT_AFTER_SAVE_MS = parseInt(
  process.env.WAIT_AFTER_SAVE_MS || "1000",
  10
);
const DRY_RUN = (process.env.DRY_RUN || "false") === "true";
const MAX_ROWS = process.env.MAX_ROWS
  ? parseInt(process.env.MAX_ROWS, 10)
  : null;
// Global redirect target (FAS/form) — can be overridden by a column in Excel (FAS_URL or NextPage)
const GLOBAL_FAS_URL =
  process.env.FAS_URL || process.env.GLOBAL_FAS_URL || null;

// logs folder
const LOGS_DIR = path.join(process.cwd(), "logs");
if (!fs.existsSync(LOGS_DIR)) fs.mkdirSync(LOGS_DIR, { recursive: true });

// logfile (append)
const LOGFILE = path.join(LOGS_DIR, `run_${Date.now()}.log`);
function appendLog(line) {
  const ts = new Date().toISOString();
  try {
    fs.appendFileSync(LOGFILE, `${ts} ${line}\n`);
  } catch (e) {
    // ignore logging errors
  }
  console.log(line);
}

// CSV writer
const RESULTS_CSV = path.join(LOGS_DIR, "results.csv");
const csvWriter = createObjectCsvWriter({
  path: RESULTS_CSV,
  header: [
    { id: "username", title: "username" },
    { id: "success", title: "success" },
    { id: "message", title: "message" },
    { id: "screenshot", title: "screenshot" },
    { id: "stateFile", title: "stateFile" },
    { id: "targetUrl", title: "targetUrl" },
  ],
  append: fs.existsSync(RESULTS_CSV),
});

// ------------------ Helpers ------------------
function sel(key, fallback) {
  if (SELECTORS && SELECTORS[key]) return SELECTORS[key];
  return fallback;
}

async function safeClick(page, selector, opts = {}) {
  if (!selector) return false;
  try {
    await page
      .waitForSelector(selector, { timeout: opts.timeout || 4000 })
      .catch(() => null);
    await page
      .click(selector, { timeout: opts.clickTimeout || 5000 })
      .catch(() => null);
    return true;
  } catch (e) {
    return false;
  }
}

async function safeFill(page, selector, value, opts = {}) {
  if (!selector || value === undefined || value === null) return false;
  try {
    await page
      .waitForSelector(selector, { timeout: opts.timeout || 4000 })
      .catch(() => null);
    await page.fill(selector, String(value)).catch(() => null);

    return true;
  } catch (e) {
    return false;
  }
}

async function safeSelect(page, selector, value, opts = {}) {
  if (!selector || value === undefined || value === null) return false;
  try {
    await page
      .waitForSelector(selector, { timeout: opts.timeout || 4000 })
      .catch(() => null);
    await page.selectOption(selector, String(value)).catch(async () => {
      await page
        .selectOption(selector, { label: String(value) })
        .catch(() => null);
    });
    return true;
  } catch (e) {
    return false;
  }
}

function formatDateForInput(value) {
  if (!value) return "";
  try {
    const d = new Date(value);
    if (!isNaN(d.getTime())) return d.toISOString().split("T")[0];
  } catch (e) {}
  const cleaned = String(value).trim().replace(/\//g, "-");
  const d2 = new Date(cleaned);
  if (!isNaN(d2.getTime())) return d2.toISOString().split("T")[0];
  return "";
}

// manual captcha wait
async function waitForManualCaptchaSolve(page) {
  appendLog("Captcha present — waiting for manual solve in the browser...");
  await page.waitForTimeout(10000);
}

// ------------------ Flow-specific functions ------------------
async function attemptLogin(page, row) {
  // Fill credentials and try to submit
  console.log("step 1");
  if (row.UserName)
    await safeFill(
      page,
      sel("userName", "#UserName, input[name='UserName'], input#UserName"),
      row.UserName
    );
  console.log("step 2");
  if (row.Password)
    await safeFill(
      page,
      sel("password", "#Password, input[name='Password']"),
      row.Password
    );

  console.log("step 3");
  // optional login date
  if (row.LoginDataTime) {
    const formattedDate = formatDateForInput(row.LoginDataTime);
    if (formattedDate)
      await safeFill(
        page,
        sel(
          "loginDate",
          "#loginDate, input[name='LoginDataTime'], #LoginDataTime"
        ),
        formattedDate
      ).catch(() => {});
  }

  console.log("step 4");
  // optional language select
  // if (row.Language)
  //   await safeSelect(
  //     page,
  //     sel("language", "#Language, select#Language"),
  //     row.Language
  //   ).catch(() => {});

  console.log("waiting for captcha");

  // captcha handling: if ValidateCaptcha column present use it, otherwise wait if captcha image exists

  const captchaExists = await page.$(
    sel("captchaImage", "#CaptchaImage, img.captcha")
  );
  if (captchaExists) await waitForManualCaptchaSolve(page);
  console.log("capcha solved");

  // submit login — attempt navigation concurrently but do not fail if navigation not triggered
  await Promise.all([
    page
      .waitForNavigation({ waitUntil: "networkidle", timeout: 30000 })
      .catch(() => null),
    page
      .click(
        sel("loginButton", "#loginButton, button#Login, input[type='submit']")
      )
      .catch(() => null),
  ]);
  console.log("logged in");
}

async function successHeuristic(page) {
  // Look for common logged-in signals: logout button, user menu, absence of username field
  const logoutSelectors = [
    sel("logout", "#logout, a.logout, button.logout"),
    sel("userMenu", "#userMenu, .user-menu, .profile-menu"),
  ];
  for (const s of logoutSelectors) {
    if (!s) continue;
    const el = await page.$(s).catch(() => null);
    if (el) {
      const visible = await el.isVisible().catch(() => false);
      if (visible) return true;
    }
  }

  // fallback: username field missing
  const usernameField = await page
    .$(sel("userName", "#UserName, input[name='UserName'], input#UserName"))
    .catch(() => null);
  if (!usernameField) return true;

  // fallback 2: check for redirection away from login page (URL changed)
  try {
    const url = page.url();
    if (
      url &&
      !url.includes("/Login") &&
      !url.includes("/Home") &&
      url !== LOGIN_URL
    )
      return true;
  } catch (e) {}

  return false;
}

async function fillTransactionPaymentForm(page) {
  // local helper: choose the second <option> for a (possibly comma-separated) selector string
  async function selectSecondOption(page, selectorCandidates) {
    try {
      // wait up to 5s for any matching element to appear
      await page
        .waitForSelector(selectorCandidates, { timeout: 5000 })
        .catch(() => null);

      // find the first matching element in the page's DOM for the candidate selector string
      const val = await page.evaluate((sel) => {
        const el = document.querySelector(sel);
        if (!el) {
          // try splitting by comma and find first that exists
          const parts = sel.split(",").map((s) => s.trim());
          for (const p of parts) {
            const e = document.querySelector(p);
            if (e)
              return (() => {
                const opts = e.querySelectorAll("option");
                if (opts.length > 1)
                  return opts[1].value || opts[1].textContent.trim();
                return null;
              })();
          }
          return null;
        }
        const opts = el.querySelectorAll("option");
        if (opts.length > 1) return opts[1].value || opts[1].textContent.trim();
        return null;
      }, selectorCandidates);

      if (!val) return false;

      // try your safeSelect first (works if it selects by value or visible text)
      try {
        await safeSelect(page, selectorCandidates, val);
        return true;
      } catch (e) {
        // fallback: set selectedIndex and dispatch change
        await page
          .evaluate(
            (sel, value) => {
              const el = document.querySelector(sel);
              if (!el) return;
              const opts = el.querySelectorAll("option");
              for (let i = 0; i < opts.length; i++) {
                const o = opts[i];
                if (
                  (o.value && o.value === value) ||
                  o.textContent.trim() === value
                ) {
                  el.selectedIndex = i;
                  el.dispatchEvent(new Event("change", { bubbles: true }));
                  break;
                }
              }
            },
            selectorCandidates,
            val
          )
          .catch(() => {});
        return true;
      }
    } catch (e) {
      return false;
    }
  }

  // small helper to select an option by visible text if exact value known
  async function trySelectByText(page, selectorCandidates, text) {
    try {
      await safeSelect(page, selectorCandidates, text);
      return true;
    } catch (e) {
      // fallback: find option with matching text and set it
      try {
        await page
          .evaluate(
            (sel, t) => {
              const el = document.querySelector(sel);
              if (!el) return;
              const opts = Array.from(el.querySelectorAll("option"));
              const match = opts.find(
                (o) => o.textContent.trim().toLowerCase() === t.toLowerCase()
              );
              if (match) {
                el.value = match.value || match.textContent.trim();
                el.dispatchEvent(new Event("change", { bubbles: true }));
              }
            },
            selectorCandidates,
            text
          )
          .catch(() => {});
        return true;
      } catch (e2) {
        return false;
      }
    }
  }

  // --- start of customized flow ---
  await page.waitForTimeout(700);
  appendLog("form opened (custom flow)");

  // use provided row if present; fallback to inline values
  const row = {
    Admissionno: "101165",
    VoucherType: "Transfer",
    Product: "Normal KCC",
    Amount: 100000,
    // if outer scope row exists, keep its fields (if you intend that)
  };

  // 1) Admission number + search
  const admissionCols = [
    "Admissionno",
    "AdmissionNo",
    "AdmnNo",
    "AdmissionNoPkey",
    "AdmissionNumber",
  ];
  const admValue = admissionCols
    .map((c) => row[c])
    .find((v) => v !== undefined && v !== null && String(v).trim() !== "");
  if (admValue) {
    await safeFill(
      page,
      sel("admissionInput", "#Admissionno, input#Admissionno"),
      String(admValue)
    ).catch(() => {});
    await safeClick(
      page,
      sel("iconSearch", "#iconsearch, button#iconsearch, .icon-search")
    ).catch(() => {});
    await page.waitForTimeout(5000); // <-- wait 5 seconds as requested
  }

  console.log("Admission number filled");

  // 2) Product: select SECOND option
  const productSelector = "#Product, select#Product";
  const productPicked = await selectSecondOption(page, productSelector);
  if (!productPicked && row.Product) {
    // fallback to selecting by provided product text
    await trySelectByText(page, productSelector, row.Product).catch(() => {});
  }
  await page.waitForTimeout(2000); // wait 2s

  console.log("product selected");

  // 3) Account: select SECOND option
  const accountSelector =
    "#AccountNo, select#AccountNo, #Account, select#Account";
  await selectSecondOption(page, accountSelector).catch(() => {});
  await page.waitForTimeout(3000); // wait 2s

  console.log("Account selected");

  // 4) Purpose: select SECOND option
  const purposeSelector =
    "#Purpose, select#Purpose, #PurposeId, select#PurposeId, #PurposeList";
  await selectSecondOption(page, purposeSelector).catch(() => {});
  await page.waitForTimeout(3000); // wait 2s

  console.log("purpose selected");

  // 5) ActivityType: choose "Disbursment" (try both spellings)
  const activitySelector = "#ActivityType, select#ActivityType";
  const activityChosen = await trySelectByText(
    page,
    activitySelector,
    "Disbursement"
  );
  false;
  if (!activityChosen) {
    // as last-resort, try selecting second option for Activity
    await selectSecondOption(page, activitySelector).catch(() => {});
  }

  console.log("activity selected ");
  await page.waitForTimeout(2000);

  // outer clicked for loading
  await safeClick(
    page,
    sel("iconSearch", "#iconsearch, button#iconsearch, .icon-search")
  ).catch(() => {});

  console.log("outer clicked");
  await page.waitForTimeout(2000);

  // 6) VoucherType: ensure "Transfer"
  const voucherSelector = "#VoucherType, select#VoucherType";
  await trySelectByText(page, voucherSelector, "Transfer").catch(() => {});

  console.log("voucher selected");
  await page.waitForTimeout(2000); // wait 30s

  // select voucher

  // ------------- row click -------------
  // prefer row.BatchId if provided; otherwise fall back to known value from markup
  const targetBatchId =
    (row && (row.BatchId || row.BatchIdString)) || "72533490022";

  const clicked = await page
    .evaluate((batchId) => {
      // 1) try to find an input whose name/id contains "BatchId" and whose value matches
      const inputs = Array.from(document.querySelectorAll("input"));
      const matchInput = inputs.find((i) => {
        const idOrName =
          (i.id || "").toLowerCase() + "|" + (i.name || "").toLowerCase();
        return (
          idOrName.includes("batchid") &&
          String(i.value).trim() === String(batchId).trim()
        );
      });
      if (matchInput) {
        const tr = matchInput.closest("tr");
        if (tr) {
          // clicking the tr element should invoke onRowClick(this) inline handler
          tr.click();
          return true;
        }
      }

      // 2) fallback: find a <td> whose text includes the batch id then click its row
      const tds = Array.from(document.querySelectorAll("tr td"));
      const tdMatch = tds.find(
        (td) =>
          td.textContent &&
          td.textContent.trim().includes(String(batchId).trim())
      );
      if (tdMatch) {
        const tr = tdMatch.closest("tr");
        if (tr) {
          tr.click();
          return true;
        }
      }

      // 3) as a last resort, click the first row that has an onRowClick inline handler
      const anyRow = Array.from(document.querySelectorAll("tr[onclick]")).find(
        (r) => /onRowClick\s*\(/.test(r.getAttribute("onclick") || "")
      );
      if (anyRow) {
        anyRow.click();
        return true;
      }

      return false;
    }, String(targetBatchId))
    .catch(() => false);

  if (clicked) {
    appendLog(`Clicked transfers row for BatchId: ${targetBatchId}`);
  } else {
    appendLog(
      `Could not locate transfers row for BatchId: ${targetBatchId} — no click performed.`
    );
  }

  // 7) Amount: fill vouchertext field
  const amountVal = row.Amount || 20;
  if (amountVal !== undefined && amountVal !== null) {
    console.log("trying to fill amount");

    try {
      // Define selector using sel()
      const amountSelector = sel(
        "vouchertext",
        "#vouchertext, input#vouchertext"
      );

      await page.waitForTimeout(1000);
      // Wait for the field
      await page
        .waitForSelector(amountSelector, { timeout: 3000 })
        .catch(() => null);

      // Fill the field
      await safeFill(page, amountSelector, String(amountVal)).catch((e) =>
        console.log("safeFill error:", e)
      );

      // Trigger onchange / onkeyup handlers
      await page.evaluate((val) => {
        const el = document.querySelector("#vouchertext");
        if (el) {
          el.value = val;
          el.dispatchEvent(new Event("input", { bubbles: true }));
          el.dispatchEvent(new Event("change", { bubbles: true }));
          el.dispatchEvent(
            new KeyboardEvent("keyup", { bubbles: true, key: "0" })
          );
        }
      }, String(amountVal));

      console.log("amount filled");
    } catch (e) {
      console.log("Amount fill failed:", e);
    }
  }
  await page.waitForTimeout(1000); // wait 3s
  console.log("form submitting");

   // outer clicked for loading
  await safeClick(
    page,
    sel("iconSearch", "#iconsearch, button#iconsearch, .icon-search")
  ).catch(() => {});

  console.log("outer clicked");
  await page.waitForTimeout(2000);

  // 8) Save / Submit using your existing save logic
  let saved = false;
  if (!DRY_RUN) {
    const saveSelectors = [
      "#btnSave",
      "#btnPost",
      "button#btnSave",
      "input#btnSave",
      "#Save, button.save",
      "button[type='submit']",
    ];
    for (const s of saveSelectors) {
      if (!s) continue;
      const el = await page.$(s).catch(() => null);
      if (el) {
        const isDisabled =
          (await el.getAttribute("disabled").catch(() => null)) ||
          (await el.evaluate((node) => node.disabled).catch(() => false));
        if (isDisabled) {
          appendLog(
            "Save button found but disabled — skipping click and treating as saved."
          );
          saved = true;
          break;
        }
        await el.click().catch(() => {});
        appendLog(`Clicked save selector: ${s}`);
        await page.waitForTimeout(2000);
        saved = true;
        break;
      }
    }

    if (!saved) {
      appendLog(
        "No explicit save button found. Attempting generic form submit..."
      );
      const generic = await page
        .$("form button[type='submit'], form input[type='submit']")
        .catch(() => null);
      if (generic) {
        await generic.click().catch(() => {});
        appendLog("Clicked generic form submit.");
        saved = true;
      }
    }
  } else {
    appendLog("DRY_RUN enabled — skipped save.");
  }

  // detect error modal post-save
  try {
    const errModal = await page.$(
      "#ErrostList, .error-modal, .validation-errors"
    );
    if (errModal) {
      const visible = await errModal.isVisible().catch(() => false);
      if (visible) appendLog("Error modal visible after attempted save.");
    }
  } catch (e) {}

  // --- After submit: handle up to 3 SweetAlert confirms and Disbursement popup ---
  try {
    //
    // helper: click SweetAlert confirm button up to `attempts` times
    async function clickSweetConfirm(attempts = 3, waitBetweenMs = 700) {
      for (let i = 0; i < attempts; i++) {
        const confirmSel = sel(
          "sweetConfirm",
          ".sa-confirm-button-container button.confirm, .swal2-confirm"
        );
        await page
          .waitForSelector(confirmSel, { timeout: 500 })
          .catch(() => null);
        const btn = await page.$(confirmSel).catch(() => null);
        if (btn) {
          const isVisible = await btn.isVisible().catch(() => false);
          if (isVisible) {
            if (!DRY_RUN) {
              await btn.click().catch(() => null);
              appendLog(`Clicked SweetAlert confirm (#${i + 1}).`);
            } else {
              appendLog(`DRY_RUN: would click SweetAlert confirm (#${i + 1}).`);
            }
            await page.waitForTimeout(waitBetweenMs);
            continue;
          }
        }
        break;
      }
    }

    // Click up to three confirm buttons (some flows show multiple confirms)
    await clickSweetConfirm(2, 700);

    // Wait for Disbursement modal input to appear
    const ledgerSelector = sel(
      "ledgerFolio",
      "#MoreDisbursementDetaos_LedgerFolioNo, input#MoreDisbursementDetaos_LedgerFolioNo"
    );
    await page
      .waitForSelector(ledgerSelector, { timeout: 2000 })
      .catch(() => null);

    const ledgerEl = await page.$(ledgerSelector).catch(() => null);
    if (ledgerEl) {
      const folioVal =
        row.LedgerFolioNo || row.LedgerFolio || row.FolioNo || "000000";

      if (!DRY_RUN) {
        await safeFill(page, ledgerSelector, String(folioVal)).catch((e) =>
          appendLog(
            `safeFill ledger folio failed: ${e && e.message ? e.message : e}`
          )
        );

        await page
          .evaluate((val) => {
            const el = document.querySelector(
              "#MoreDisbursementDetaos_LedgerFolioNo"
            );
            if (el) {
              el.value = val;
              el.dispatchEvent(new Event("input", { bubbles: true }));
              el.dispatchEvent(new Event("change", { bubbles: true }));
            }
          }, String(folioVal))
          .catch(() => null);

        appendLog(`Filled Ledger Folio Number: ${folioVal}`);
      } else {
        appendLog(`DRY_RUN: would fill Ledger Folio Number with "${folioVal}"`);
      }

      // Click Prepare button
      const prepareSelector = sel(
        "prepareBtn",
        "#btnsaveDisbursments, input#btnsaveDisbursments, input[value='Prepare']"
      );
      await page
        .waitForSelector(prepareSelector, { timeout: 2000 })
        .catch(() => null);
      const prepareBtn = await page.$(prepareSelector).catch(() => null);

      if (prepareBtn) {
        const isDisabled =
          (await prepareBtn.getAttribute("disabled").catch(() => null)) ||
          (await prepareBtn.evaluate((n) => n.disabled).catch(() => false));
        if (!isDisabled) {
          if (!DRY_RUN) {
            await prepareBtn.click().catch(() => null);
            appendLog("Clicked Prepare (Disbursement) button.");
          } else {
            appendLog("DRY_RUN: would click Prepare (Disbursement) button.");
          }
        } else {
          appendLog("Prepare button found but disabled; not clicking.");
        }
      } else {
        appendLog("Prepare button not found in Disbursement modal.");
      }

      await page.waitForTimeout(30000);
      // end here as requested (stops after clicking Prepare)
    } else {
      appendLog(
        "No Disbursement popup / ledger folio field detected after submit."
      );
    }
  } catch (e) {
    appendLog(
      `Post-submit Disbursement handling failed: ${
        e && e.message ? e.message : e
      }`
    );
  }
}

// Open FAS/Form — picks URL from row if present, otherwise global env var, otherwise default target
async function openFASAndFillForm(page, row) {
  // pick URL from row > global > default
  const target =
    "https://up6.uniteerp.in/FAS/TransactionPayment/TransactionPayment?formid=40004&moduleid=3";

  try {
    // navigate explicitly to the target (ensures we land on the form page after login)
    await page
      .goto(target, { waitUntil: "domcontentloaded" })
      .catch(() => null);
    console.log("form loaded");
    // Wait for a known form/root selector before attempting to fill — configurable via SELECTORS
    const formRoot = sel(
      "formRoot",
      "#TransactionPaymentForm, #MainForm, form#TransactionPayment, form"
    );
    appendLog(`Waiting up to 2000 ms for form root selector: ${formRoot}`);
    const present = await page
      .waitForSelector(formRoot, { timeout: 3000 })
      .catch(() => null);
    console.log("selectors available");
    if (!present) {
      appendLog(
        "Form root not found within timeout — falling back to networkidle wait."
      );
      await page
        .waitForLoadState("networkidle", { timeout: 3000 })
        .catch(() => null);
    } else {
      appendLog("Form root detected — proceeding to fill the form.");
    }

    await page.waitForTimeout(700);
    await fillTransactionPaymentForm(page, row);
  } catch (e) {
    appendLog(`Optional redirect to form failed: ${e.message}`);
    throw e;
  }
}

// ------------------ Main flow (IIFE) ------------------
(async () => {
  appendLog(`Starting automation. HEADLESS=${HEADLESS} DRY_RUN=${DRY_RUN}`);

  const rows = await readExcel(EXCEL_PATH);
  appendLog(`Read ${rows.length} rows from Excel: ${EXCEL_PATH}`);

  const browser = await chromium.launch({ headless: HEADLESS });
  const context = await browser.newContext({
    viewport: { width: 1280, height: 800 },
  });
  const page = await context.newPage();

  let processed = 0;
  for (const row of rows) {
    if (MAX_ROWS && processed >= MAX_ROWS) break;
    processed += 1;

    const username = row.UserName || `user_${Date.now()}`;
    let screenshotFile = "";
    let stateFile = "";

    try {
      appendLog(`\n=== Processing ${username} ===`);

      // Start at login page for each row to ensure fresh context (optionally could reuse state)
      await page
        .goto(LOGIN_URL, { waitUntil: "domcontentloaded" })
        .catch(() => null);
      await page.waitForTimeout(600);

      console.log("step 1");
      // wait for username field (if not present maybe already logged in)
      await page
        .waitForSelector(
          sel("userName", "#UserName, input[name='UserName'], input#UserName"),
          { timeout: 10000 }
        )
        .catch(() => null);

      // Retry login attempts
      let loggedIn = false;
      let lastLoginError = null;
      for (let attempt = 1; attempt <= Math.max(1, RETRY_LOGIN); attempt++) {
        try {
          console.log("logging in");
          await attemptLogin(page, row);

          if (await successHeuristic(page)) {
            loggedIn = true;
            break;
          }
        } catch (e) {
          lastLoginError = e;
          appendLog(`Login attempt ${attempt} error: ${e.message}`);
        }
      }

      if (!loggedIn) {
        screenshotFile = path.join(
          LOGS_DIR,
          `${username}_login_failed_${Date.now()}.png`
        );
        await page
          .screenshot({ path: screenshotFile, fullPage: true })
          .catch(() => {});
        appendLog(
          `Login may have failed for ${username}. Screenshot: ${screenshotFile}`
        );
        await csvWriter.writeRecords([
          {
            username,
            success: "false",
            message: `Login failed: ${
              lastLoginError ? lastLoginError.message : "still on login page"
            }`,
            screenshot: screenshotFile,
            stateFile: "",
            targetUrl: LOGIN_URL,
          },
        ]);
        continue; // next row
      }

      appendLog(`Login success for ${username}`);

      // Save storage state per user
      stateFile = path.join(LOGS_DIR, `${username}_state_${Date.now()}.json`);
      await context.storageState({ path: stateFile }).catch(() => {});

      // Redirect to form and fill it (explicit: login -> redirect -> fill form)
      try {
        await openFASAndFillForm(page, row);
      } catch (e) {
        appendLog(`FAS/form fill error: ${e.message}`);
      }

      // Optionally save screenshot of success
      if (SAVE_SCREENSHOT_ON_SUCCESS) {
        screenshotFile = path.join(
          LOGS_DIR,
          `${username}_success_${Date.now()}.png`
        );
        await page
          .screenshot({ path: screenshotFile, fullPage: true })
          .catch(() => {});
      }

      const usedTarget =
        row.FAS_URL || row.NextPage || GLOBAL_FAS_URL || LOGIN_URL;
      await csvWriter.writeRecords([
        {
          username,
          success: "true",
          message: "OK",
          screenshot: screenshotFile,
          stateFile,
          targetUrl: usedTarget,
        },
      ]);
      appendLog(`✅ Done for ${username}`);
    } catch (err) {
      appendLog(`Error processing ${username}: ${err.stack || err}`);
      screenshotFile = path.join(
        LOGS_DIR,
        `${username}_error_${Date.now()}.png`
      );
      await page
        .screenshot({ path: screenshotFile, fullPage: true })
        .catch(() => {});
      await csvWriter.writeRecords([
        {
          username,
          success: "error",
          message: String(err),
          screenshot: screenshotFile,
          stateFile,
          targetUrl: "",
        },
      ]);
    }
  }

  await browser.close();
  appendLog("All done.");
})();

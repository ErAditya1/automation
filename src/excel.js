// src/excel.js
import fs from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';
import XLSX from 'xlsx';

/**
 * Reads the first worksheet from the given Excel file and returns an array of normalized rows.
 * Expected header names (case-insensitive): UserName, Password, Language, LoginDataTime, ValidateCaptcha, NextPage, Field1, Field2
 * Supports: .xlsx (exceljs), and falls back to sheetjs (xlsx) for other formats (.xls, .csv) or corrupted xlsx attempts.
 *
 * @param {string} filePath
 * @returns {Promise<Array<Object>>}
 */
export async function readExcel(filePath) {
  const resolved = path.resolve(filePath);

  if (!fs.existsSync(resolved)) {
    throw new Error(`Excel file not found: ${resolved}`);
  }

  const stats = fs.statSync(resolved);
  if (!stats.isFile() || stats.size === 0) {
    throw new Error(`Excel file is empty or not a regular file: ${resolved}`);
  }

  // Try exceljs (good for well-formed .xlsx)
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(resolved);

    const worksheet = workbook.worksheets[0];
    if (!worksheet) return [];

    // Build headers from first row
    const headerRow = worksheet.getRow(1);
    
    const headers = [];
    headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      headers[colNumber] = (cell.value ?? '').toString().trim();
    });

    const rows = [];
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const obj = {};
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const header = (headers[colNumber] || '').trim();
        if (header) {
          obj[header] = cell.value !== null && cell.value !== undefined ? cell.value.toString().trim() : '';
        }
      });

      // Normalize keys
      rows.push(normalizeRow(obj));
    });

    return rows;
  } catch (err) {
    console.warn('exceljs read failed (trying fallback). Reason:', err.message || err);
    // fallback to sheetjs (xlsx)
  }

  // Fallback: use sheetjs (xlsx)
  try {
    const workbook = XLSX.readFile(resolved, { cellDates: true });
    const sheetName = workbook.SheetNames[0];
    if (!sheetName) return [];

    const sheet = workbook.Sheets[sheetName];
    const raw = XLSX.utils.sheet_to_json(sheet, { defval: '' }); // array of objects, keys from header

    // Normalize each row
    const rows = raw.map((r) => {
      // convert keys to simple object where header names are used
      const obj = {};
      Object.keys(r).forEach((k) => {
        const val = r[k] === null || r[k] === undefined ? '' : String(r[k]).trim();
        obj[k] = val;
      });
      return normalizeRow(obj);
    });

    return rows;
  } catch (err) {
    throw new Error(`Failed to read Excel with both exceljs and xlsx: ${err.message || err}`);
  }
}

/** normalize header names to expected fields */
function normalizeRow(obj) {
  const get = (names) => {
    for (const n of names) {
      if (obj.hasOwnProperty(n)) return obj[n];
      const lower = Object.keys(obj).find((k) => k.toLowerCase() === n.toLowerCase());
      if (lower) return obj[lower];
    }
    return '';
  };

  return {
    UserName: get(['UserName', 'username', 'user']) || '',
    Password: get(['Password', 'password']) || '',
    Language: get(['Language', 'language']) || '',
    LoginDataTime: get(['LoginDataTime', 'loginDate', 'LoginDate']) || '',
    ValidateCaptcha: get(['ValidateCaptcha', 'captcha', 'ValidateCaptcha']) || '',
    NextPage: get(['NextPage', 'nextpage']) || '',
    Field1: get(['Field1', 'field1']) || '',
    Field2: get(['Field2', 'field2']) || ''
  };
}

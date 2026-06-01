"use strict";

const fs = require("fs");
const path = require("path");

const DEFAULT_MAX_EXCEL_MB = 50;
const MAX_EXCEL_BYTES = Math.max(1, Number(process.env.MAX_EXCEL_MB || DEFAULT_MAX_EXCEL_MB)) * 1024 * 1024;

function assertExcelReadable(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext !== ".xlsx" && ext !== ".xlsm") {
    throw new Error(`Only .xlsx/.xlsm Excel files are supported: ${filePath}`);
  }

  const stat = fs.statSync(filePath);
  if (!stat.isFile()) throw new Error(`Excel path is not a file: ${filePath}`);
  if (stat.size > MAX_EXCEL_BYTES) {
    throw new Error(`Excel file is too large (${Math.ceil(stat.size / 1024 / 1024)} MB). Limit: ${Math.ceil(MAX_EXCEL_BYTES / 1024 / 1024)} MB.`);
  }

  return stat;
}

module.exports = { assertExcelReadable, MAX_EXCEL_BYTES };

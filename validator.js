/* global Office, Excel */

Office.onReady(() => {});

/**
 * eBay UK Validate: red (ERROR) / yellow (WARNING) + comments + Validation Report
 * - Validates the ACTIVE worksheet
 * - Assumes headers are in row 1
 */
async function validateEbayUK() {
  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const used = sheet.getUsedRangeOrNullObject();
    used.load(["rowCount", "values"]);
    await ctx.sync();

    if (used.isNullObject || used.rowCount < 2) {
      await ensureReport(ctx, "No data found (need header + at least 1 row).");
      return;
    }

    const values = used.values;
    const headers = values[0].map(h => String(h ?? "").trim());

    const colIndex = (name) =>
      headers.findIndex(h => h.toLowerCase() === name.toLowerCase());

    const cSKU = colIndex("SKU");
    const cTitle = colIndex("Title");
    const cLoc = colIndex("Localized For");
    const cEAN = colIndex("EAN");
    const cPic1 = colIndex("Picture URL 1");

    const missing = [];
    if (cTitle < 0) missing.push("Title");
    if (cLoc < 0) missing.push("Localized For");
    if (missing.length) {
      await ensureReport(ctx, `Missing required columns: ${missing.join(", ")}`);
      return;
    }

    const report = await getOrCreateReportSheet(ctx);
    clearSheet(report);
    writeReportHeader(report);

    // Special characters mentioned in checklist (customize anytime)
    const forbiddenCharsRe = /[“”‘’•§™®©✓★→|~^]/g;
    // Emoji detection (broad)
    const emojiRe = /[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}]/gu;

    let outRow = 2;

    for (let r = 1; r < values.length; r++) {
      const excelRow = r + 1;
      const sku = cSKU >= 0 ? String(values[r][cSKU] ?? "").trim() : "";

      // Localized For must be en_GB (ERROR)
      const locVal = String(values[r][cLoc] ?? "").trim();
      if (!locVal || locVal.toLowerCase() !== "en_gb") {
        outRow = await flag(ctx, sheet, report, outRow, excelRow, cLoc, sku, "Localized For",
          "ERROR", `Localized For must be en_GB (found: ${locVal || "blank"})`, locVal);
      }

      // Title checks
      const title = String(values[r][cTitle] ?? "").trim();
      if (!title) {
        outRow = await flag(ctx, sheet, report, outRow, excelRow, cTitle, sku, "Title",
          "ERROR", "Title is blank", "");
      } else {
        if (title.length > 80) {
          outRow = await flag(ctx, sheet, report, outRow, excelRow, cTitle, sku, "Title",
            "ERROR", `Title exceeds 80 characters (len=${title.length})`, title);
        }

        const badChars = title.match(forbiddenCharsRe);
        if (badChars && badChars.length) {
          const uniq = Array.from(new Set(badChars)).join(" ");
          outRow = await flag(ctx, sheet, report, outRow, excelRow, cTitle, sku, "Title",
            "WARNING", `Unnecessary special characters detected: ${uniq}`, title);
        }

        const emojis = title.match(emojiRe);
        if (emojis && emojis.length) {
          const uniqE = Array.from(new Set(emojis)).join(" ");
          outRow = await flag(ctx, sheet, report, outRow, excelRow, cTitle, sku, "Title",
            "WARNING", `Emojis detected: ${uniqE}`, title);
        }
      }

      // EAN must be 13 digits (ERROR when non-blank but invalid)
      if (cEAN >= 0) {
        const ean = String(values[r][cEAN] ?? "").trim();
        if (ean && !/^\d{13}$/.test(ean)) {
          outRow = await flag(ctx, sheet, report, outRow, excelRow, cEAN, sku, "EAN",
            "ERROR", "EAN must be 13 digits", ean);
        }
      }

      // Picture URL 1 required (ERROR)
      if (cPic1 >= 0) {
        const pic1 = String(values[r][cPic1] ?? "").trim();
        if (!pic1) {
          outRow = await flag(ctx, sheet, report, outRow, excelRow, cPic1, sku, "Picture URL 1",
            "ERROR", "Picture URL 1 is required", "");
        }
      }
    }

    autofit(report);
    report.activate();
    await ctx.sync();
  });
}

/** Clear fills + delete comments (best-effort) */
async function clearHighlightsAndNotes() {
  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const used = sheet.getUsedRangeOrNullObject();
    used.load("rowCount");
    await ctx.sync();
    if (!used.isNullObject && used.rowCount > 0) used.format.fill.clear();

    const comments = sheet.comments;
    comments.load("items");
    await ctx.sync();
    comments.items.forEach(c => c.delete());

    await ctx.sync();
  });
}

/** Open / create Validation Report */
async function openValidationReport() {
  await Excel.run(async (ctx) => {
    const report = await getOrCreateReportSheet(ctx);
    report.activate();
    await ctx.sync();
  });
}

/* ----- Helpers ----- */
async function getOrCreateReportSheet(ctx) {
  const name = "Validation Report";
  const existing = ctx.workbook.worksheets.getItemOrNullObject(name);
  existing.load("name");
  await ctx.sync();
  return existing.isNullObject ? ctx.workbook.worksheets.add(name) : existing;
}

function clearSheet(sheet) {
  const r = sheet.getUsedRangeOrNullObject();
  try { r.clear(); } catch (e) {}
}

function writeReportHeader(report) {
  report.getRange("A1:F1").values = [[
    "Row#", "SKU", "Column", "Severity", "Message", "Value"
  ]];
  report.getRange("A1:F1").format.font.bold = true;
}

function autofit(report) {
  report.getUsedRangeOrNullObject().format.autofitColumns();
}

async function ensureReport(ctx, message) {
  const report = await getOrCreateReportSheet(ctx);
  clearSheet(report);
  report.getRange("A1").values = [[message]];
  report.activate();
  await ctx.sync();
}

async function flag(ctx, sheet, report, outRow, excelRow, zeroBasedCol, sku, colName, severity, message, value) {
  const cell = sheet.getCell(excelRow - 1, zeroBasedCol);

  // Red / Yellow
  cell.format.fill.color = severity === "ERROR" ? "#FFC7CE" : "#FFEB9C";

  // Comment (best-effort)
  try {
    const existing = cell.getComments();
    existing.load("items");
    await ctx.sync();
    const line = `${severity}: ${message}`;
    if (!existing.items.length) {
      sheet.comments.add(cell, line);
    } else {
      const c = existing.items[0];
      c.load("content");
      await ctx.sync();
      c.content = (c.content ? (c.content + "\n") : "") + line;
    }
  } catch (e) {
    // If tenant blocks comments, report still captures everything.
  }

  // Report row
  report.getRange(`A${outRow}:F${outRow}`).values = [[
    excelRow, sku, colName, severity, message, value
  ]];

  return outRow + 1;
}

// Expose functions to Office ribbon
globalThis.validateEbayUK = validateEbayUK;
globalThis.clearHighlightsAndNotes = clearHighlightsAndNotes;
globalThis.openValidationReport = openValidationReport;

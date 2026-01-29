/* global Office, Excel */

Office.onReady(() => {});

/**
 * eBay UK Validate
 * - Red = ERROR (must change)
 * - Yellow = WARNING (review)
 * - Adds comments + Validation Report
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

    const col = (name) =>
      headers.findIndex(h => h.toLowerCase() === name.toLowerCase());

    const cSKU = col("SKU");
    const cTitle = col("Title");
    const cLoc = col("Localized For");
    const cEAN = col("EAN");
    const cPic1 = col("Picture URL 1");

    if (cTitle < 0 || cLoc < 0) {
      await ensureReport(ctx, "Missing required columns: Title / Localized For");
      return;
    }

    const report = await getOrCreateReport(ctx);
    clearSheet(report);
    writeHeader(report);

    const forbiddenChars = /[“”‘’•§™®©✓★→|~^]/g;
    const emojiRe = /[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}]/gu;

    let outRow = 2;

    for (let r = 1; r < values.length; r++) {
      const excelRow = r + 1;
      const sku = cSKU >= 0 ? String(values[r][cSKU] ?? "") : "";

      // Localized For (ERROR)
      const loc = String(values[r][cLoc] ?? "").trim();
      if (loc.toLowerCase() !== "en_gb") {
        outRow = await flag(ctx, sheet, report, outRow, excelRow, cLoc, sku,
          "Localized For", "ERROR", "Must be en_GB", loc);
      }

      // Title checks
      const title = String(values[r][cTitle] ?? "").trim();
      if (!title) {
        outRow = await flag(ctx, sheet, report, outRow, excelRow, cTitle, sku,
          "Title", "ERROR", "Title is blank", "");
      } else {
        if (title.length > 80) {
          outRow = await flag(ctx, sheet, report, outRow, excelRow, cTitle, sku,
            "Title", "ERROR", `Exceeds 80 characters (${title.length})`, title);
        }

        const bad = title.match(forbiddenChars);
        if (bad) {
          outRow = await flag(ctx, sheet, report, outRow, excelRow, cTitle, sku,
            "Title", "WARNING", `Special characters: ${[...new Set(bad)].join(" ")}`, title);
        }

        const emojis = title.match(emojiRe);
        if (emojis) {
          outRow = await flag(ctx, sheet, report, outRow, excelRow, cTitle, sku,
            "Title", "WARNING", `Emojis detected: ${[...new Set(emojis)].join(" ")}`, title);
        }
      }

      // EAN (ERROR)
      if (cEAN >= 0) {
        const ean = String(values[r][cEAN] ?? "").trim();
        if (ean && !/^\d{13}$/.test(ean)) {
          outRow = await flag(ctx, sheet, report, outRow, excelRow, cEAN, sku,
            "EAN", "ERROR", "EAN must be 13 digits", ean);
        }
      }

      // Image 1 required (ERROR)
      if (cPic1 >= 0) {
        const pic1 = String(values[r][cPic1] ?? "").trim();
        if (!pic1) {
          outRow = await flag(ctx, sheet, report, outRow, excelRow, cPic1, sku,
            "Picture URL 1", "ERROR", "Image 1 is mandatory", "");
        }
      }
    }

    report.getUsedRangeOrNullObject().format.autofitColumns();
    report.activate();
    await ctx.sync();
  });
}

async function clearHighlightsAndNotes() {
  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    sheet.getUsedRangeOrNullObject().format.fill.clear();

    const comments = sheet.comments;
    comments.load("items");
    await ctx.sync();
    comments.items.forEach(c => c.delete());
    await ctx.sync();
  });
}

async function openValidationReport() {
  await Excel.run(async (ctx) => {
    const r = await getOrCreateReport(ctx);
    r.activate();
    await ctx.sync();
  });
}

/* Helpers */
async function getOrCreateReport(ctx) {
  const name = "Validation Report";
  const s = ctx.workbook.worksheets.getItemOrNullObject(name);
  s.load("name");
  await ctx.sync();
  return s.isNullObject ? ctx.workbook.worksheets.add(name) : s;
}

function clearSheet(s) {
  try { s.getUsedRangeOrNullObject().clear(); } catch {}
}

function writeHeader(r) {
  r.getRange("A1:F1").values = [[
    "Row#", "SKU", "Column", "Severity", "Message", "Value"
  ]];
  r.getRange("A1:F1").format.font.bold = true;
}

async function ensureReport(ctx, msg) {
  const r = await getOrCreateReport(ctx);
  clearSheet(r);
  r.getRange("A1").values = [[msg]];
  r.activate();
  await ctx.sync();
}

async function flag(ctx, sheet, report, outRow, row, col, sku, colName, sev, msg, val) {
  const cell = sheet.getCell(row - 1, col);
  cell.format.fill.color = sev === "ERROR" ? "#FFC7CE" : "#FFEB9C";

  try {
    const ex = cell.getComments();
    ex.load("items");
    await ctx.sync();
    const line = `${sev}: ${msg}`;
    if (!ex.items.length) sheet.comments.add(cell, line);
    else ex.items[0].content += `\n${line}`;
  } catch {}

  report.getRange(`A${outRow}:F${outRow}`).values = [[
    row, sku, colName, sev, msg, val
  ]];

  return outRow + 1;
}

globalThis.validateEbayUK = validateEbayUK;
globalThis.clearHighlightsAndNotes = clearHighlightsAndNotes;
globalThis.openValidationReport = openValidationReport;

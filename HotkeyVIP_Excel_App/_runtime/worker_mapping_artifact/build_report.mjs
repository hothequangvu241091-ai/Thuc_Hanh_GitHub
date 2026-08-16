import fs from "node:fs/promises";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const root = "D:/HotkeyVIP_Excel_App";
const inputPath = `${root}/_runtime/worker_mapping_artifact/mapping.json`;
const outputDir = `${root}/outputs/worker_profile_word_errors_20260816`;
const outputPath = `${outputDir}/word_error_worker_mapping.xlsx`;
const previewPath = `${root}/_runtime/worker_mapping_artifact/preview.png`;

const payload = JSON.parse(await fs.readFile(inputPath, "utf8"));
const wb = Workbook.create();
const ws = wb.worksheets.add("WORD_ERROR theo Worker");
ws.showGridLines = false;

ws.getRange("A1:M1").merge();
ws.getRange("A1").values = [["DANH SÁCH WORD_ERROR THEO WORKER PROFILE"]];
ws.getRange("A1:M1").format = {
  fill: "#17365D",
  font: { bold: true, color: "#FFFFFF", size: 16 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
ws.getRange("A1:M1").format.rowHeight = 32;

ws.getRange("A2:M2").merge();
ws.getRange("A2").values = [[`Nguồn: ${payload.source} | Sheet: ${payload.sheet} | Đối chiếu URL chính xác với Edge History`]];
ws.getRange("A2:M2").format = {
  fill: "#DCE6F1",
  font: { italic: true, color: "#44546A", size: 10 },
  horizontalAlignment: "left",
  verticalAlignment: "center",
};
ws.getRange("A2:M2").format.rowHeight = 24;

ws.getRange("A4:F4").values = [["Tổng lỗi", payload.total, "worker_1", payload.worker_1, "worker_3", payload.worker_3]];
for (const range of ["A4:B4", "C4:D4", "E4:F4"]) {
  ws.getRange(range).format = {
    fill: range === "A4:B4" ? "#E2F0D9" : range === "C4:D4" ? "#DDEBF7" : "#FCE4D6",
    font: { bold: true, color: "#1F1F1F" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    borders: { preset: "all", style: "thin", color: "#A6A6A6" },
  };
}

const headers = [
  "STT", "Dòng Excel", "Tên Miền", "Tiêu đề SEO", "H1", "Từ khóa",
  "URL ChatGPT", "Worker profile", "Trạng thái viết", "Lỗi viết",
  "Đường dẫn Word", "Có đường dẫn ảnh 1", "Có đường dẫn ảnh 2",
];
const rows = payload.records.map((r) => [
  r.stt, r.excel_row, r.domain ?? "", r.seo_title ?? "", r.h1 ?? "", r.keyword ?? "",
  r.url ?? "", r.worker, r.status ?? "", r.error ?? "", r.word_path ?? "",
  r.image_1_exists ? "Có" : "Không", r.image_2_exists ? "Có" : "Không",
]);
ws.getRange("A6:M6").values = [headers];
ws.getRange(`A7:M${6 + rows.length}`).values = rows;

const allRange = ws.getRange(`A6:M${6 + rows.length}`);
allRange.format = {
  font: { size: 10, color: "#222222" },
  verticalAlignment: "top",
  borders: { preset: "all", style: "thin", color: "#D9E2F3" },
};
ws.getRange("A6:M6").format = {
  fill: "#2F75B5",
  font: { bold: true, color: "#FFFFFF", size: 10 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
  borders: { preset: "all", style: "thin", color: "#FFFFFF" },
};
ws.getRange("A6:M6").format.rowHeight = 32;
ws.getRange(`A7:B${6 + rows.length}`).format.horizontalAlignment = "center";
ws.getRange(`H7:I${6 + rows.length}`).format.horizontalAlignment = "center";
ws.getRange(`L7:M${6 + rows.length}`).format.horizontalAlignment = "center";
ws.getRange(`C7:M${6 + rows.length}`).format.wrapText = true;

ws.getRange(`H7:H${6 + rows.length}`).conditionalFormats.addCustom('=H7="worker_1"', {
  fill: "#DDEBF7", font: { bold: true, color: "#1F4E78" },
});
ws.getRange(`H7:H${6 + rows.length}`).conditionalFormats.addCustom('=H7="worker_3"', {
  fill: "#FCE4D6", font: { bold: true, color: "#9C5700" },
});

const table = ws.tables.add(`A6:M${6 + rows.length}`, true, "WordErrorWorkerMapping");
table.style = "TableStyleMedium2";
table.showFilterButton = true;
table.showBandedRows = true;

const widths = [7, 11, 20, 34, 34, 25, 48, 16, 17, 48, 42, 16, 16];
for (let i = 0; i < widths.length; i += 1) {
  ws.getRangeByIndexes(5, i, rows.length + 1, 1).format.columnWidth = widths[i];
}
ws.getRange(`A7:M${6 + rows.length}`).format.rowHeight = 36;
ws.freezePanes.freezeRows(6);
ws.freezePanes.freezeColumns(2);

await fs.mkdir(outputDir, { recursive: true });
const file = await SpreadsheetFile.exportXlsx(wb);
await file.save(outputPath);

const inspection = await wb.inspect({
  kind: "region,table,formula",
  sheetId: "WORD_ERROR theo Worker",
  range: "A1:M14",
  maxChars: 6000,
  tableMaxRows: 12,
  tableMaxCols: 13,
});
console.log(inspection.ndjson);

const errors = await wb.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
  maxChars: 3000,
});
console.log("FORMULA_ERROR_SCAN");
console.log(errors.ndjson);

const preview = await wb.render({ sheetName: "WORD_ERROR theo Worker", range: "A1:M18", scale: 1, format: "png" });
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));
console.log(JSON.stringify({ outputPath, previewPath, rows: rows.length }));

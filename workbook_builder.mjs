import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const sourcePath = "G:/project_calendar/캘린더(자동).xlsx";
const workDir = "G:/project_calendar/.work";
const outputDir = "G:/project_calendar/outputs/cellendar";
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(sourcePath));

const monthSheet = workbook.worksheets.getItem("26.8");
monthSheet.getRange("J1:Q1").merge();
monthSheet.getRange("J1").values = [["Cellendar 간편 일정 입력"]];
monthSheet.getRange("J1:Q1").format = {
  fill: "#172033",
  font: { bold: true, color: "#FFFFFF", size: 16 },
  rowHeight: 32,
  horizontalAlignment: "left",
  verticalAlignment: "center",
};
monthSheet.getRange("J2:Q2").merge();
monthSheet.getRange("J2").values = [["날짜·제목·메모를 입력하세요. ID·수정시간·삭제시간은 앱이 관리합니다."]];
monthSheet.getRange("J2:Q2").format = { fill: "#EAF2FF", font: { color: "#24466E" }, rowHeight: 26 };
monthSheet.getRange("J4:Q5").values = [
  ["날짜", "제목", "메모", "색상", "알림시각", "ID", "수정시간", "삭제시간"],
  [null, "", "", "#4f8ef7", "", "", "", ""],
];
monthSheet.getRange("J4:Q4").format = { fill: "#315D8A", font: { bold: true, color: "#FFFFFF" }, rowHeight: 24 };
monthSheet.getRange("J5").format.numberFormat = "yyyy-mm-dd";
monthSheet.getRange("N5").format.numberFormat = "hh:mm";
monthSheet.getRange("P5:Q5").format.numberFormat = "yyyy-mm-dd hh:mm:ss";
monthSheet.getRange("J:Q").format.autofitColumns();
monthSheet.getRange("J:J").format.columnWidth = 14;
monthSheet.getRange("K:K").format.columnWidth = 26;
monthSheet.getRange("L:L").format.columnWidth = 34;
monthSheet.getRange("M:N").format.columnWidth = 13;
monthSheet.getRange("O:O").format.columnWidth = 34;
monthSheet.getRange("P:Q").format.columnWidth = 23;
monthSheet.getRange("O4:Q5").format = { fill: "#F2F4F7", font: { color: "#667085" } };
const eventTable = monthSheet.tables.add("J4:Q5", true, "CalendarEvents");
eventTable.style = "TableStyleMedium2";
eventTable.showBandedRows = true;

for (const memoRow of [4, 6, 8, 10, 12, 14]) {
  for (const col of ["B", "C", "D", "E", "F", "G", "H"]) {
    const dateCell = `${col}${memoRow - 1}`;
    monthSheet.getRange(`${col}${memoRow}`).formulas = [[
      `=IFERROR(TEXTJOIN(CHAR(10),TRUE,IF(($J$5:$J$504=${dateCell})*($Q$5:$Q$504=""),$K$5:$K$504,"")),"")`,
    ]];
  }
}
monthSheet.getRange("B4:H14").format.wrapText = true;

await fs.mkdir(`${workDir}/final`, { recursive: true });
await fs.mkdir(outputDir, { recursive: true });
for (const name of ["공휴일", "26.8"]) {
  const preview = await workbook.render({ sheetName: name, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(`${workDir}/final/${name}.png`, new Uint8Array(await preview.arrayBuffer()));
}

const check = await workbook.inspect({ kind: "table", sheetId: "26.8", range: "J4:Q5", include: "values,formulas", tableMaxRows: 5, tableMaxCols: 8 });
console.log(check.ndjson);
const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#CALC!", options: { useRegex: true, maxResults: 100 }, summary: "formula error scan" });
console.log(errors.ndjson);
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(`${outputDir}/캘린더(자동).xlsx`);

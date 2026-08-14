import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const sourcePath = "G:/project_calendar/캘린더(자동).xlsx";
const workDir = "G:/project_calendar/.work";
const outputDir = "G:/project_calendar/outputs/cellendar";
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(sourcePath));

const dataSheet = workbook.worksheets.add("일정데이터");
dataSheet.showGridLines = false;
dataSheet.getRange("A1:H1").merge();
dataSheet.getRange("A1").values = [["Cellendar 일정 동기화 데이터"]];
dataSheet.getRange("A1:H1").format = {
  fill: "#172033",
  font: { bold: true, color: "#FFFFFF", size: 16 },
  rowHeight: 32,
  horizontalAlignment: "left",
  verticalAlignment: "center",
};
dataSheet.getRange("A2:H2").merge();
dataSheet.getRange("A2").values = [["날짜와 제목을 입력하세요. ID·수정시간·삭제시간은 앱이 관리하므로 수정하지 마세요."]];
dataSheet.getRange("A2:H2").format = { fill: "#EAF2FF", font: { color: "#24466E" }, rowHeight: 26 };
dataSheet.getRange("A4:H5").values = [
  ["ID", "날짜", "제목", "메모", "색상", "알림시각", "수정시간", "삭제시간"],
  ["", null, "", "", "#4f8ef7", "", "", ""],
];
dataSheet.getRange("A4:H4").format = { fill: "#315D8A", font: { bold: true, color: "#FFFFFF" }, rowHeight: 24 };
dataSheet.getRange("B5").format.numberFormat = "yyyy-mm-dd";
dataSheet.getRange("F5").format.numberFormat = "hh:mm";
dataSheet.getRange("G5:H5").format.numberFormat = "yyyy-mm-dd hh:mm:ss";
dataSheet.getRange("A4:H5").format.autofitColumns();
dataSheet.getRange("A:A").format.columnWidth = 34;
dataSheet.getRange("B:B").format.columnWidth = 14;
dataSheet.getRange("C:C").format.columnWidth = 28;
dataSheet.getRange("D:D").format.columnWidth = 36;
dataSheet.getRange("E:F").format.columnWidth = 13;
dataSheet.getRange("G:H").format.columnWidth = 23;
dataSheet.freezePanes.freezeRows(4);
const eventTable = dataSheet.tables.add("A4:H5", true, "CalendarEvents");
eventTable.style = "TableStyleMedium2";
eventTable.showBandedRows = true;

const monthSheet = workbook.worksheets.getItem("26.8");
for (const memoRow of [4, 6, 8, 10, 12, 14]) {
  for (const col of ["B", "C", "D", "E", "F", "G", "H"]) {
    const dateCell = `${col}${memoRow - 1}`;
    monthSheet.getRange(`${col}${memoRow}`).formulas = [[
      `=IFERROR(TEXTJOIN(CHAR(10),TRUE,IF(('일정데이터'!$B$5:$B$504=${dateCell})*('일정데이터'!$H$5:$H$504=""),'일정데이터'!$C$5:$C$504,"")),"")`,
    ]];
  }
}
monthSheet.getRange("B4:H14").format.wrapText = true;

await fs.mkdir(`${workDir}/final`, { recursive: true });
await fs.mkdir(outputDir, { recursive: true });
for (const name of ["공휴일", "26.8", "일정데이터"]) {
  const preview = await workbook.render({ sheetName: name, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(`${workDir}/final/${name}.png`, new Uint8Array(await preview.arrayBuffer()));
}

const check = await workbook.inspect({ kind: "table", sheetId: "일정데이터", range: "A4:H5", include: "values,formulas", tableMaxRows: 5, tableMaxCols: 8 });
console.log(check.ndjson);
const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#CALC!", options: { useRegex: true, maxResults: 100 }, summary: "formula error scan" });
console.log(errors.ndjson);
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(`${outputDir}/캘린더(자동).xlsx`);

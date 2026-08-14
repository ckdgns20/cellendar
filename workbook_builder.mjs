import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const sourcePath = "G:/project_calendar/캘린더(자동).xlsx";
const workDir = "G:/project_calendar/.work";
const outputDir = "G:/project_calendar/outputs/cellendar";
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(sourcePath));

const monthSheet = workbook.worksheets.getItem("26.8");
const firstInputRow = 5;
const lastInputRow = 34;

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
monthSheet.getRange("J2").values = [["노란 칸만 입력하세요. 날짜와 색상은 목록에서 선택할 수 있습니다."]];
monthSheet.getRange("J2:Q2").format = {
  fill: "#EAF2FF",
  font: { color: "#24466E" },
  rowHeight: 26,
};

const headers = ["날짜", "제목", "메모", "알림시각", "색상", "ID", "수정시간", "삭제시간"];
const rows = Array.from({ length: lastInputRow - firstInputRow + 1 }, () => Array(8).fill(""));
monthSheet.getRange(`J4:Q${lastInputRow}`).values = [headers, ...rows];
monthSheet.getRange("J4:Q4").format = {
  fill: "#315D8A",
  font: { bold: true, color: "#FFFFFF" },
  rowHeight: 24,
  horizontalAlignment: "center",
};
monthSheet.getRange(`J${firstInputRow}:N${lastInputRow}`).format = {
  fill: "#FFF7D6",
  font: { color: "#243447" },
  borders: { preset: "all", style: "thin", color: "#E7DCA8" },
};
monthSheet.getRange(`O${firstInputRow}:Q${lastInputRow}`).format = {
  fill: "#F2F4F7",
  font: { color: "#7A8492", size: 9 },
  borders: { preset: "all", style: "thin", color: "#D9DEE5" },
};

monthSheet.getRange(`J${firstInputRow}:J${lastInputRow}`).format.numberFormat = "yyyy-mm-dd";
monthSheet.getRange(`M${firstInputRow}:M${lastInputRow}`).format.numberFormat = "hh:mm";
monthSheet.getRange(`P${firstInputRow}:Q${lastInputRow}`).format.numberFormat = "yyyy-mm-dd hh:mm:ss";
monthSheet.getRange("J:J").format.columnWidth = 14;
monthSheet.getRange("K:K").format.columnWidth = 24;
monthSheet.getRange("L:L").format.columnWidth = 32;
monthSheet.getRange("M:M").format.columnWidth = 13;
monthSheet.getRange("N:N").format.columnWidth = 12;
monthSheet.getRange("O:O").format.columnWidth = 18;
monthSheet.getRange("P:Q").format.columnWidth = 18;
monthSheet.getRange(`J${firstInputRow}:Q${lastInputRow}`).format.rowHeight = 22;

// 현재 달력의 연도(G1)와 월(H1)에 따라 날짜 선택 목록을 자동으로 만듭니다.
monthSheet.getRange("S1:T1").values = [["자동 날짜 목록", "색상 목록"]];
monthSheet.getRange("S2:S32").formulas = Array.from({ length: 31 }, (_, index) => [
  `=IF(${index + 1}<=DAY(EOMONTH(DATE($G$1,$H$1,1),0)),DATE($G$1,$H$1,${index + 1}),"")`,
]);
monthSheet.getRange("S2:S32").format.numberFormat = "yyyy-mm-dd";
monthSheet.getRange("T2:T8").values = [["파랑"], ["빨강"], ["초록"], ["주황"], ["보라"], ["분홍"], ["회색"]];
monthSheet.getRange("S1:T32").format = { fill: "#F7F9FC", font: { color: "#98A2B3", size: 9 } };
monthSheet.getRange("S:S").format.columnWidth = 14;
monthSheet.getRange("T:T").format.columnWidth = 11;

// 같은 날짜의 여러 일정을 안정적으로 찾기 위한 내부 순번과 조회 키입니다.
monthSheet.getRange("R4").values = [["내부 순번"]];
monthSheet.getRange("U4").values = [["내부 조회 키"]];
monthSheet.getRange("R5").formulas = [[`=IF(OR(J5="",Q5<>""),"",COUNTIF($J$5:J5,J5))`]];
monthSheet.getRange("R5:R504").fillDown();
monthSheet.getRange("U5").formulas = [[`=IF(OR(J5="",Q5<>""),"",J5&"|"&R5)`]];
monthSheet.getRange("U5:U504").fillDown();
monthSheet.getRange("R4:R504").format = { fill: "#F7F9FC", font: { color: "#C2C8D0", size: 8 } };
monthSheet.getRange("U4:U504").format = { fill: "#F7F9FC", font: { color: "#C2C8D0", size: 8 } };
monthSheet.getRange("R:R").format.columnWidth = 9;
monthSheet.getRange("U:U").format.columnWidth = 14;

monthSheet.getRange(`J${firstInputRow}:J${lastInputRow}`).dataValidation = {
  rule: { type: "list", formula1: "=$S$2:$S$32" },
};
monthSheet.getRange(`N${firstInputRow}:N${lastInputRow}`).dataValidation = {
  rule: { type: "list", formula1: "=$T$2:$T$8" },
};

const eventTable = monthSheet.tables.add(`J4:Q${lastInputRow}`, true, "CalendarEvents");
eventTable.style = "TableStyleMedium2";
eventTable.showBandedRows = false;

for (const memoRow of [4, 6, 8, 10, 12, 14]) {
  for (const col of ["B", "C", "D", "E", "F", "G", "H"]) {
    const dateCell = `${col}${memoRow - 1}`;
    const eventAt = (position) =>
      `XLOOKUP(${dateCell}&"|"&${position},$U$5:$U$504,$K$5:$K$504,"")`;
    monthSheet.getRange(`${col}${memoRow}`).formulas = [[
      `=${eventAt(1)}&IF(${eventAt(2)}="","",CHAR(10)&${eventAt(2)})&IF(${eventAt(3)}="","",CHAR(10)&${eventAt(3)})&IF(${eventAt(4)}="","",CHAR(10)&${eventAt(4)})&IF(${eventAt(5)}="","",CHAR(10)&${eventAt(5)})`,
    ]];
  }
}
monthSheet.getRange("B:H").format.columnWidth = 24;
monthSheet.getRange("B3:H14").format.wrapText = true;
monthSheet.getRange("B3:H3").format.rowHeight = 22;
monthSheet.getRange("B5:H5").format.rowHeight = 22;
monthSheet.getRange("B7:H7").format.rowHeight = 22;
monthSheet.getRange("B9:H9").format.rowHeight = 22;
monthSheet.getRange("B11:H11").format.rowHeight = 22;
monthSheet.getRange("B13:H13").format.rowHeight = 22;
for (const eventRow of [4, 6, 8, 10, 12, 14]) {
  monthSheet.getRange(`B${eventRow}:H${eventRow}`).format.rowHeight = 72;
  monthSheet.getRange(`B${eventRow}:H${eventRow}`).format.verticalAlignment = "top";
}

// 달력 아래의 빈 공간에 사용자 안내를 배치합니다.
monthSheet.getRange("B17:I17").merge();
monthSheet.getRange("B17").values = [["Cellendar 사용 방법"]];
monthSheet.getRange("B17:I17").format = {
  fill: "#172033",
  font: { bold: true, color: "#FFFFFF", size: 15 },
  rowHeight: 30,
  horizontalAlignment: "left",
  verticalAlignment: "center",
};
const guideRows = [
  ["1. 일정 입력", "오른쪽 노란 표에서 날짜와 색상을 목록으로 선택하고 제목을 입력하세요."],
  ["2. 달력 표시", "같은 날짜의 일정은 달력 칸에 줄바꿈으로 최대 5개까지 표시됩니다."],
  ["3. 다음 달", "상단의 연도와 월 숫자만 바꾸면 달력과 날짜 선택 목록이 함께 변경됩니다."],
  ["4. 비고 사용", "초록색 비고 영역은 동기화와 관계없이 자유롭게 작성할 수 있습니다."],
  ["5. 앱 동기화", "엑셀을 OneDrive의 Cellendar 폴더에 저장한 뒤 앱의 동기화 버튼을 누르세요."],
  ["6. 수정·삭제", "휴대폰 앱에서 일정을 누르면 수정하거나 삭제할 수 있습니다."],
  ["주의", "회색 ID·수정시간·삭제시간 칸과 내부 조회 영역은 수정하지 마세요."],
];
guideRows.forEach(([title, description], index) => {
  const row = 18 + index;
  monthSheet.getRange(`B${row}:C${row}`).merge();
  monthSheet.getRange(`D${row}:I${row}`).merge();
  monthSheet.getRange(`B${row}`).values = [[title]];
  monthSheet.getRange(`D${row}`).values = [[description]];
  monthSheet.getRange(`B${row}:C${row}`).format = {
    fill: index === guideRows.length - 1 ? "#FFF0F0" : "#EAF2FF",
    font: { bold: true, color: index === guideRows.length - 1 ? "#B33D3D" : "#315D8A" },
    rowHeight: 27,
    verticalAlignment: "center",
  };
  monthSheet.getRange(`D${row}:I${row}`).format = {
    fill: index === guideRows.length - 1 ? "#FFF8F8" : "#F8FAFD",
    font: { color: "#3D4A5C" },
    rowHeight: 27,
    verticalAlignment: "center",
    wrapText: true,
  };
});

await fs.mkdir(`${workDir}/final`, { recursive: true });
await fs.mkdir(outputDir, { recursive: true });
for (const name of ["공휴일", "26.8"]) {
  const preview = await workbook.render({
    sheetName: name,
    ...(name === "26.8" ? { range: "B1:Q34" } : { autoCrop: "all" }),
    scale: 1,
    format: "png",
  });
  await fs.writeFile(`${workDir}/final/${name}.png`, new Uint8Array(await preview.arrayBuffer()));
}

const check = await workbook.inspect({
  kind: "table",
  sheetId: "26.8",
  range: "J4:T10",
  include: "values,formulas",
  tableMaxRows: 10,
  tableMaxCols: 11,
});
console.log(check.ndjson);
const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#CALC!",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan",
});
console.log(errors.ndjson);
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(`${outputDir}/캘린더(자동).xlsx`);

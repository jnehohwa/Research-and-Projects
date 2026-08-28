import fs from "node:fs/promises";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const outputDir = "/Users/joshuanehohwa/Documents/Research and Projects/outputs/01a01d0b-e6ae-7072-a0f7-515334d16c9d";
const outputPath = `${outputDir}/MCAT_Retake_Study_Schedule_2026-2027.xlsx`;
const workbook = Workbook.create();

const dashboard = workbook.worksheets.add("Dashboard");
const daily = workbook.worksheets.add("Daily Schedule");
const weekly = workbook.worksheets.add("Weekly Overview");
const fullLengths = workbook.worksheets.add("Full-Length Tracker");
const errors = workbook.worksheets.add("Error Log");
const inventory = workbook.worksheets.add("Resource Inventory");
const classAvailability = workbook.worksheets.add("Class & Availability");

const COLORS = {
  purple: "#5B3B88",
  purpleDark: "#3D255E",
  lavender: "#EDE7F6",
  lavender2: "#D9CCF2",
  orange: "#F4A261",
  orangeLight: "#FCE3D1",
  green: "#D9EAD3",
  greenDark: "#38761D",
  red: "#F4CCCC",
  redDark: "#990000",
  yellow: "#FFF2CC",
  blue: "#D9EAF7",
  gray: "#F3F4F6",
  grayDark: "#5F6368",
  white: "#FFFFFF",
  ink: "#25212B",
};

function dateUTC(year, month, day) {
  return new Date(Date.UTC(year, month - 1, day, 12, 0, 0));
}
function isoDate(date) {
  return date.toISOString().slice(0, 10);
}
function addDays(date, days) {
  return new Date(date.getTime() + days * 86400000);
}
function weekdaysBetween(start, end) {
  const out = [];
  for (let d = start; d <= end; d = addDays(d, 1)) {
    const day = d.getUTCDay();
    if (day >= 1 && day <= 5) out.push(new Date(d));
  }
  return out;
}
function excelCol(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
function setWidths(sheet, widths) {
  widths.forEach((width, i) => {
    sheet.getRange(`${excelCol(i + 1)}:${excelCol(i + 1)}`).format.columnWidth = width;
  });
}
function styleTitle(sheet, range, title, subtitle) {
  const titleRange = sheet.getRange(range);
  titleRange.merge();
  titleRange.values = [[title]];
  titleRange.format.fill = COLORS.purpleDark;
  titleRange.format.font = { bold: true, color: COLORS.white, size: 18 };
  titleRange.format.verticalAlignment = "center";
  titleRange.format.rowHeight = 34;
  if (subtitle) {
    const endCol = range.split(":")[1].replace(/[0-9]/g, "");
    const subtitleRange = sheet.getRange(`A2:${endCol}2`);
    subtitleRange.merge();
    subtitleRange.values = [[subtitle]];
    subtitleRange.format.fill = COLORS.lavender;
    subtitleRange.format.font = { color: COLORS.purpleDark, italic: true, size: 10 };
    subtitleRange.format.wrapText = true;
    subtitleRange.format.rowHeight = 30;
  }
}
function styleHeader(range) {
  range.format.fill = COLORS.purple;
  range.format.font = { bold: true, color: COLORS.white };
  range.format.verticalAlignment = "center";
  range.format.wrapText = true;
  range.format.rowHeight = 30;
  range.format.borders = { preset: "inside", style: "thin", color: "#8F79B2" };
}
function styleData(range) {
  range.format.font = { color: COLORS.ink, size: 10 };
  range.format.verticalAlignment = "top";
  range.format.borders = { insideHorizontal: { style: "thin", color: "#E5E2EA" } };
}
function chunkResource(resource, total, chunk) {
  const sessions = [];
  let start = 1;
  while (start <= total) {
    const end = Math.min(total, start + chunk - 1);
    sessions.push({ resource, start, end, count: end - start + 1 });
    start = end + 1;
  }
  return sessions;
}

const studyStart = dateUTC(2026, 8, 31);
const studyEnd = dateUTC(2027, 1, 8);
const examDate = dateUTC(2027, 1, 9);
const studyDates = weekdaysBetween(studyStart, studyEnd);

const flWeeks = new Map([
  ["2026-09-21", { exam: 1, review: "retake; track recognized questions" }],
  ["2026-10-19", { exam: 2, review: "retake; track recognized questions" }],
  ["2026-11-16", { exam: 3, review: "first fresh scored benchmark" }],
  ["2026-12-07", { exam: 4, review: "fresh scored benchmark" }],
  ["2026-12-21", { exam: 5, review: "place flexibly before holiday break" }],
  ["2027-01-04", { exam: 6, review: "final fresh scored benchmark" }],
]);

const phaseFor = (d) => {
  const s = isoDate(d);
  if (s <= "2026-09-18") return "Foundation rebuild";
  if (s <= "2026-10-23") return "Targeted repair + timing";
  if (s <= "2026-11-20") return "Mixed timed practice";
  if (s <= "2026-12-11") return "Exam-like integration";
  if (s <= "2026-12-24") return "Focused repair";
  if (s <= "2027-01-01") return "Consolidation";
  return "Final taper";
};

const weekMonday = (d) => {
  const day = d.getUTCDay();
  return addDays(d, -(day - 1));
};

const carsSessions = [
  ...chunkResource("CARS Diagnostic Tool", 179, 9),
  ...chunkResource("CARS Question Pack Vol. 1", 120, 10),
  ...chunkResource("CARS Question Pack Vol. 2", 120, 10),
];
const primarySessions = [
  ...chunkResource("Biology Question Pack Vol. 1", 120, 30),
  ...chunkResource("Chemistry Question Pack", 120, 30),
  ...chunkResource("Physics Question Pack", 120, 30),
  ...chunkResource("Biology Question Pack Vol. 2", 120, 30),
  ...chunkResource("Section Bank Vol. 1", 300, 30),
  ...chunkResource("Section Bank Vol. 2", 300, 30),
  ...chunkResource("Independent Question Bank", 150, 30),
  ...chunkResource("Content Outline Course", 30, 10),
];

let carsIndex = 0;
let primaryIndex = 0;
const dailyRows = [];
const addDaily = (date, task, resource, section, block, hours, questions, flexibility, notes = "") => {
  dailyRows.push([
    date, task, resource, section, block, hours, questions, "Not Started", null,
    "Medium", "No", flexibility, notes,
  ]);
};

for (const d of studyDates) {
  const s = isoDate(d);
  const monday = isoDate(weekMonday(d));
  const fl = flWeeks.get(monday);
  const weekday = d.getUTCDay();
  const phase = phaseFor(d);
  const isHolidayWeek = monday === "2026-12-21" || monday === "2026-12-28";
  const isFinalWeek = monday === "2027-01-04";

  if (s === "2026-12-25") {
    addDaily(d, "Protected holiday recovery day", "Rest / Recovery", "Wellbeing", "Recovery", 0, 0, "Do not move", "No required MCAT work; optional Anki does not affect completion.");
    continue;
  }

  if (fl) {
    if (weekday === 1) {
      const mondayHours = isFinalWeek ? 2.0 : (isHolidayWeek ? 3.5 : 7.5);
      addDaily(d, `Practice Exam ${fl.exam} readiness: Anki, equations, pacing, and familiar passages`, "Anki / Error Log", "All Sections", "Pre-exam consolidation", mondayHours, 0, "Keep on Monday", `No unfamiliar heavy content; prepare for Tuesday's exam. ${fl.review}.`);
    } else if (weekday === 2) {
      addDaily(d, `Practice Exam ${fl.exam} — Tuesday after Epidemiology`, `Practice Exam ${fl.exam}`, "All Sections", "Full length", 7.5, 230, "Fixed Tuesday exam day", `Begin around 09:45 after Epidemiology ends at 09:15; ${fl.review}.`);
    } else if (weekday === 3) {
      addDaily(d, `Practice Exam ${fl.exam} review`, `Practice Exam ${fl.exam}`, "All Sections", "Full-length review", isFinalWeek ? 6.0 : 6.5, 0, "Keep on Wednesday", "Work around Kinesiology 10:00–10:50 and Personal Health 12:00–12:50. Review every incorrect, guessed, low-confidence, or recognized question before new heavy practice.");
    } else {
      const taperHours = isFinalWeek ? 2.0 : (isHolidayWeek ? (weekday === 5 ? 2.0 : 3.5) : (weekday === 4 ? 6.5 : 2.0));
      addDaily(d, "Anki + equations + high-yield recall", "Anki", "Mixed", weekday === 5 ? "Friday maintenance" : "Active recall", isFinalWeek ? 0.75 : 1.0, 0, "Movable within week", phase);
      addDaily(d, weekday === 5 ? "Light CARS or error-log review; coursework and homework remain the priority" : "Targeted remediation from full-length error log", "Khan Academy / Error Log", weekday % 2 ? "Bio/Biochem" : "CARS", "Content repair", taperHours - (isFinalWeek ? 0.75 : 1.0), 0, "Movable within week", weekday === 5 ? "Stop after this block and protect the rest of Friday for classwork." : "Use Khan only for a diagnosed gap; otherwise use AAMC explanations and reasoning review.");
    }
    continue;
  }

  if (isHolidayWeek || isFinalWeek) {
    const hours = weekday === 5 ? 2.0 : (isFinalWeek ? 2.0 : 3.5);
    addDaily(d, "Light Anki, equations, and familiar concepts", "Anki / Error Log", "Mixed", "Consolidation", hours, 0, "Movable within week", `${phase}; no heavy new content.`);
    continue;
  }

  if (weekday === 5) {
    addDaily(d, "Protected Friday Anki maintenance", "Anki", "Mixed", "Friday maintenance", 1.0, 0, "Complete around the two Friday classes", "Friday is primarily reserved for Kinesiology, Personal Health, and weekly homework.");
    if (carsIndex < carsSessions.length) {
      const q = carsSessions[carsIndex++];
      addDaily(d, `${q.resource} questions ${q.start}–${q.end} + concise review`, q.resource, "CARS", "Friday CARS maintenance", 1.0, q.count, "Movable within Friday", "Stop after the light CARS block; protect remaining time for coursework and homework.");
    } else {
      addDaily(d, "Light CARS error-log remediation", "CARS Error Log", "CARS", "Friday CARS maintenance", 1.0, 0, "Movable within Friday", "Stop after the light CARS block; protect remaining time for coursework and homework.");
    }
    addDaily(d, "Coursework and weekly homework protected", "College Coursework", "School", "Protected coursework", 0, 0, "Do not replace with MCAT work", "Kinesiology 10:00–10:50 and Personal Health 12:00–12:50; use the remainder of Friday for homework and course responsibilities.");
    continue;
  }

  addDaily(d, "Anki review and active recall", "Anki", "Mixed", weekday === 3 ? "Protected flashcard session" : "Active recall", weekday === 3 ? 1.25 : 0.75, 0, "Movable within week", weekday === 3 ? "Protected weekday Anki session; weekends remain optional." : "Prioritize due cards and cards from the error log.");
  addDaily(d, "Targeted content repair", "Khan Academy / Error Log", weekday % 2 === 0 ? "CARS / Bio/Biochem" : "Chem/Phys / Psych/Soc", "Content repair", weekday === 3 ? 1.0 : (weekday === 4 ? 1.25 : 1.5), 0, "Movable within week", "Only watch/read lessons tied to a diagnosed gap or low-confidence topic.");

  if (primaryIndex < primarySessions.length) {
    const q = primarySessions[primaryIndex++];
    const section = q.resource.includes("Biology") ? "Bio/Biochem" : q.resource.includes("Chemistry") || q.resource.includes("Physics") ? "Chem/Phys" : q.resource.includes("Section Bank") || q.resource.includes("Independent") ? "Mixed" : "Mixed";
    addDaily(d, `${q.resource} questions ${q.start}–${q.end} (timed)`, q.resource, section, "Timed AAMC questions", weekday === 4 ? 1.5 : 1.75, q.count, "Movable within week", "Flag guesses and low-confidence answers before checking explanations.");
  } else {
    addDaily(d, "Mixed timed set from flagged AAMC questions", "AAMC flagged questions", "Mixed", "Timed AAMC questions", weekday === 4 ? 1.5 : 1.75, 0, "Movable within week", "Do not count repeated flagged questions toward inventory completion twice.");
  }
  addDaily(d, "Review timed set and update error log", "AAMC / Error Log", "Mixed", "Question review", weekday === 4 ? 1.75 : 2.25, 0, "Movable within week", "Record error type, correct reasoning, content repair, and whether an Anki card is warranted.");

  if (carsIndex < carsSessions.length) {
    const q = carsSessions[carsIndex++];
    addDaily(d, `${q.resource} questions ${q.start}–${q.end} + review`, q.resource, "CARS", "CARS practice", 1.25, q.count, "Movable within week", "Focus on passage mapping, main idea, evidence, timing, and why each wrong choice fails.");
  } else {
    addDaily(d, "CARS error-log remediation or familiar passage redo", "CARS Error Log", "CARS", "CARS practice", 1.25, 0, "Movable within week", "Use only reviewed passages; do not inflate fresh-question totals.");
  }
}

// Add any remaining official resource sessions before the final taper as explicit flexible catch-up rows.
const catchupDates = studyDates.filter(d => isoDate(d) >= "2026-12-14" && isoDate(d) <= "2026-12-31");
let catchupCursor = 0;
while (primaryIndex < primarySessions.length) {
  const q = primarySessions[primaryIndex++];
  const d = catchupDates[catchupCursor++ % catchupDates.length];
  addDaily(d, `${q.resource} questions ${q.start}–${q.end} — inventory reconciliation`, q.resource, "Mixed", "Flexible catch-up", 1.5, q.count, "Move within this or the prior week", "Complete only after full-length review; replaces an equivalent remediation block.");
}
while (carsIndex < carsSessions.length) {
  const q = carsSessions[carsIndex++];
  const d = catchupDates[catchupCursor++ % catchupDates.length];
  addDaily(d, `${q.resource} questions ${q.start}–${q.end} — CARS reconciliation`, q.resource, "CARS", "Flexible catch-up", 1.25, q.count, "Move within this or the prior week", "Replaces that day's CARS remediation block.");
}

dailyRows.sort((a, b) => a[0] - b[0] || String(a[4]).localeCompare(String(b[4])));

// Daily Schedule
styleTitle(daily, "A1:M1", "MCAT Retake Daily Schedule", "Assignments may move within their week, except full lengths: every full length is fixed for Tuesday after Epidemiology, with review on Wednesday.");
const dailyHeaders = [["Date", "Task", "Resource", "Section", "Study Block", "Estimated Hours", "Question Target", "Status", "Accuracy", "Confidence", "Review Needed", "Flexibility", "Notes"]];
daily.getRange("A3:M3").values = dailyHeaders;
styleHeader(daily.getRange("A3:M3"));
daily.getRange(`A4:M${dailyRows.length + 3}`).values = dailyRows;
styleData(daily.getRange(`A4:M${dailyRows.length + 3}`));
daily.getRange(`A4:A${dailyRows.length + 3}`).format.numberFormat = "ddd, dd mmm yyyy";
daily.getRange(`F4:F${dailyRows.length + 3}`).format.numberFormat = "0.00";
daily.getRange(`G4:G${dailyRows.length + 3}`).format.numberFormat = "0";
daily.getRange(`I4:I${dailyRows.length + 3}`).format.numberFormat = "0.0%";
daily.getRange(`H4:H${dailyRows.length + 3}`).dataValidation = { rule: { type: "list", values: ["Not Started", "In Progress", "Review Needed", "Complete"] } };
daily.getRange(`J4:J${dailyRows.length + 3}`).dataValidation = { rule: { type: "list", values: ["Low", "Medium", "High"] } };
daily.getRange(`K4:K${dailyRows.length + 3}`).dataValidation = { rule: { type: "list", values: ["No", "Yes", "Complete"] } };
daily.getRange(`H4:H${dailyRows.length + 3}`).conditionalFormats.add("containsText", { text: "Complete", format: { fill: COLORS.green, font: { color: COLORS.greenDark } } });
daily.getRange(`H4:H${dailyRows.length + 3}`).conditionalFormats.add("containsText", { text: "Review Needed", format: { fill: COLORS.yellow, font: { color: "#7F6000" } } });
daily.getRange(`H4:H${dailyRows.length + 3}`).conditionalFormats.add("containsText", { text: "Not Started", format: { fill: COLORS.red, font: { color: COLORS.redDark } } });
daily.getRange(`I4:I${dailyRows.length + 3}`).conditionalFormats.add("colorScale", { colors: [COLORS.red, COLORS.yellow, COLORS.green] });
daily.freezePanes.freezeRows(3);
daily.freezePanes.freezeColumns(1);
daily.showGridLines = false;
setWidths(daily, [15, 38, 27, 16, 22, 13, 13, 16, 11, 12, 15, 24, 48]);
daily.getRange(`B4:M${dailyRows.length + 3}`).format.wrapText = true;

// Resource Inventory
styleTitle(inventory, "A1:H1", "AAMC Resource Inventory", "All resources are reset for the retake. Completed counts update only when matching Daily Schedule rows are marked Complete.");
const resources = [
  ["Independent Question Bank", 150, "Mixed", "78%", "Reset and redo"],
  ["Section Bank Vol. 1", 300, "Mixed", "91%", "Reset and redo"],
  ["Section Bank Vol. 2", 300, "Mixed", "88%", "Reset and redo"],
  ["CARS Question Pack Vol. 1", 120, "CARS", "100%", "Reset and redo"],
  ["CARS Question Pack Vol. 2", 120, "CARS", "75%", "Reset and redo"],
  ["Biology Question Pack Vol. 1", 120, "Bio/Biochem", "8%", "Reset and redo"],
  ["Biology Question Pack Vol. 2", 120, "Bio/Biochem", "8%", "Reset and redo"],
  ["Chemistry Question Pack", 120, "Chem/Phys", "58%", "Reset and redo"],
  ["Physics Question Pack", 120, "Chem/Phys", "39%", "Reset and redo"],
  ["CARS Diagnostic Tool", 179, "CARS", "5%", "Reset and redo"],
  ["Content Outline Course", 30, "Mixed", "0%", "Reset and complete"],
];
inventory.getRange("A3:H3").values = [["Resource", "Total Questions", "Section", "Starting Completion", "Reset Plan", "Assigned", "Completed", "Remaining"]];
styleHeader(inventory.getRange("A3:H3"));
inventory.getRange(`A4:E${resources.length + 3}`).values = resources;
for (let i = 0; i < resources.length; i++) {
  const row = i + 4;
  inventory.getRange(`F${row}`).formulas = [[`=SUMIF('Daily Schedule'!$C$4:$C$${dailyRows.length + 3},A${row},'Daily Schedule'!$G$4:$G$${dailyRows.length + 3})`]];
  inventory.getRange(`G${row}`).formulas = [[`=SUMIFS('Daily Schedule'!$G$4:$G$${dailyRows.length + 3},'Daily Schedule'!$C$4:$C$${dailyRows.length + 3},A${row},'Daily Schedule'!$H$4:$H$${dailyRows.length + 3},"Complete")`]];
  inventory.getRange(`H${row}`).formulas = [[`=MAX(0,B${row}-G${row})`]];
}
const invTotalRow = resources.length + 5;
inventory.getRange(`A${invTotalRow}:H${invTotalRow}`).values = [["Question-bank total", null, "", "", "", null, null, null]];
inventory.getRange(`B${invTotalRow}`).formulas = [[`=SUM(B4:B${resources.length + 3})`]];
inventory.getRange(`F${invTotalRow}`).formulas = [[`=SUM(F4:F${resources.length + 3})`]];
inventory.getRange(`G${invTotalRow}`).formulas = [[`=SUM(G4:G${resources.length + 3})`]];
inventory.getRange(`H${invTotalRow}`).formulas = [[`=SUM(H4:H${resources.length + 3})`]];
inventory.getRange(`A${invTotalRow}:H${invTotalRow}`).format.fill = COLORS.lavender2;
inventory.getRange(`A${invTotalRow}:H${invTotalRow}`).format.font = { bold: true, color: COLORS.purpleDark };
styleData(inventory.getRange(`A4:H${invTotalRow}`));
inventory.getRange(`B4:B${invTotalRow}`).format.numberFormat = "0";
inventory.getRange(`F4:H${invTotalRow}`).format.numberFormat = "0";
inventory.freezePanes.freezeRows(3);
inventory.showGridLines = false;
setWidths(inventory, [34, 15, 16, 18, 22, 12, 12, 12]);

// Class schedule and realistic study availability
styleTitle(classAvailability, "A1:G1", "Class Schedule & MCAT Availability", "Corrected semester schedule: classes remain protected, Tuesday is the full-length day, and Friday stays primarily reserved for classes and weekly homework.");
classAvailability.getRange("A3:G3").values = [["Day", "Class Commitments", "Suggested MCAT Windows", "MCAT Hour Target", "Friday / Homework Rule", "Full-Length Handling", "Notes"]];
styleHeader(classAvailability.getRange("A3:G3"));
const availabilityRows = [
  ["Monday", "Kinesiology 10:00–10:50; Personal Health 12:00–12:50", "07:30–09:30; 11:00–11:50; 13:10–17:50", 7.5, "—", "Pre-exam consolidation on full-length weeks.", "Protect both classes and use the longer afternoon block."],
  ["Tuesday", "Epidemiology 08:00–09:15", "09:45–12:45; 13:30–17:00; 18:00–19:00", 7.5, "—", "All full lengths start around 09:45 after Epidemiology.", "The early class leaves enough time for a complete exam day."],
  ["Wednesday", "Kinesiology 10:00–10:50; Personal Health 12:00–12:50", "07:30–09:30; 11:00–11:50; 13:10–17:50", 7.5, "Protected weekday Anki session", "Full-length review occurs Wednesday around both classes and may finish by 16:50.", "Primary flashcard day outside exam-review weeks."],
  ["Thursday", "Epidemiology 08:00–09:15; NSC 1 15:00–16:50", "09:45–13:15; 13:30–14:30; 17:15–19:15", 6.5, "—", "Post-exam remediation may continue here.", "Slightly lighter because of two separated classes."],
  ["Friday", "Kinesiology 10:00–10:50; Personal Health 12:00–12:50; weekly homework", "07:45–09:15; 11:00–11:30", 2.0, "Stop after Anki + light CARS; use remaining time for homework", "Never a planned full-length day.", "Friday MCAT work is maintenance, not a spillover day."],
  ["Saturday", "No scheduled class", "No required MCAT block", 0, "Protected rest", "Available only if the learner deliberately chooses it.", "Optional Anki never affects tracker completion."],
  ["Sunday", "No scheduled class", "No required MCAT block", 0, "Protected rest", "Available only if the learner deliberately chooses it.", "Optional Anki never affects tracker completion."],
];
classAvailability.getRange("A4:G10").values = availabilityRows;
styleData(classAvailability.getRange("A4:G10"));
classAvailability.getRange("D4:D10").format.numberFormat = "0.0";
classAvailability.getRange("A4:A8").format.fill = COLORS.green;
classAvailability.getRange("A9:A10").format.fill = COLORS.lavender;
classAvailability.getRange("E8:E10").format.fill = COLORS.yellow;
classAvailability.getRange("B4:G10").format.wrapText = true;
classAvailability.freezePanes.freezeRows(3);
classAvailability.showGridLines = false;
setWidths(classAvailability, [15, 35, 48, 17, 45, 46, 38]);
classAvailability.getRange("A12:G12").merge();
classAvailability.getRange("A12:G12").values = [["Default non-exam week: 31.0 MCAT hours. Full-length weeks use Tuesday after Epidemiology for the exam and Wednesday for review. Friday classes and homework remain protected."]];
classAvailability.getRange("A12:G12").format.fill = COLORS.lavender2;
classAvailability.getRange("A12:G12").format.font = { bold: true, color: COLORS.purpleDark };
classAvailability.getRange("A12:G12").format.rowHeight = 28;

// Full-Length Tracker
styleTitle(fullLengths, "A1:N1", "Full-Length Exam Tracker", "All six exams are scheduled for Tuesday after Epidemiology. Exams 1–2 are retakes and should not be treated as clean predictive scores.");
fullLengths.getRange("A3:N3").values = [["Exam", "Scheduled Tuesday", "Actual Date", "Type", "C/P", "CARS", "B/B", "P/S", "Total", "Timing Problems", "Stamina Notes", "Recognized Questions", "Review Status", "Key Next Action"]];
styleHeader(fullLengths.getRange("A3:N3"));
const flRows = [...flWeeks.entries()].map(([week, data]) => [
  `Practice Exam ${data.exam}`, addDays(dateUTC(...week.split("-").map(Number)), 1), null,
  data.exam <= 2 ? "Retake — recognition risk" : "Fresh scored benchmark",
  null, null, null, null, null, "", "", null, "Not Started", "",
]);
fullLengths.getRange("A4:N9").values = flRows;
for (let row = 4; row <= 9; row++) {
  fullLengths.getRange(`I${row}`).formulas = [[`=IF(COUNT(E${row}:H${row})=4,SUM(E${row}:H${row}),"")`]];
}
fullLengths.getRange("B4:C9").format.numberFormat = "dd mmm yyyy";
fullLengths.getRange("E4:I9").format.numberFormat = "0";
fullLengths.getRange("L4:L9").format.numberFormat = "0";
fullLengths.getRange("M4:M9").dataValidation = { rule: { type: "list", values: ["Not Started", "In Progress", "Review Needed", "Complete"] } };
fullLengths.getRange("M4:M9").conditionalFormats.add("containsText", { text: "Complete", format: { fill: COLORS.green, font: { color: COLORS.greenDark } } });
fullLengths.getRange("M4:M9").conditionalFormats.add("containsText", { text: "Review Needed", format: { fill: COLORS.yellow } });
fullLengths.getRange("I4:I9").conditionalFormats.add("colorScale", { colors: [COLORS.red, COLORS.yellow, COLORS.green] });
styleData(fullLengths.getRange("A4:N9"));
fullLengths.freezePanes.freezeRows(3);
fullLengths.showGridLines = false;
setWidths(fullLengths, [19, 15, 15, 24, 9, 9, 9, 9, 10, 26, 26, 19, 16, 34]);
fullLengths.getRange("D4:N9").format.wrapText = true;

// Error Log
styleTitle(errors, "A1:N1", "MCAT Error Log", "Log every incorrect, guessed, low-confidence, or recognized question. Create Anki cards only for durable facts, equations, distinctions, or reasoning cues.");
errors.getRange("A3:N3").values = [["Date", "Source", "Question / Topic", "Section", "Error Type", "Confidence", "Recognized?", "Reasoning Failure", "Correct Approach", "Content Repair", "Anki Card?", "Revisit Date", "Revisit Result", "Resolved?"]];
styleHeader(errors.getRange("A3:N3"));
const emptyRows = Array.from({ length: 150 }, () => Array(14).fill(null));
errors.getRange("A4:N153").values = emptyRows;
styleData(errors.getRange("A4:N153"));
errors.getRange("A4:A153").format.numberFormat = "dd mmm yyyy";
errors.getRange("L4:L153").format.numberFormat = "dd mmm yyyy";
errors.getRange("D4:D153").dataValidation = { rule: { type: "list", values: ["Chem/Phys", "CARS", "Bio/Biochem", "Psych/Soc"] } };
errors.getRange("E4:E153").dataValidation = { rule: { type: "list", values: ["Content Gap", "Reasoning", "Passage Interpretation", "Timing", "Calculation", "Careless Error", "Memory / Recall"] } };
errors.getRange("F4:F153").dataValidation = { rule: { type: "list", values: ["Low", "Medium", "High"] } };
errors.getRange("G4:G153").dataValidation = { rule: { type: "list", values: ["No", "Yes"] } };
errors.getRange("K4:K153").dataValidation = { rule: { type: "list", values: ["No", "Yes", "Created"] } };
errors.getRange("N4:N153").dataValidation = { rule: { type: "list", values: ["No", "Yes"] } };
errors.getRange("N4:N153").conditionalFormats.add("containsText", { text: "Yes", format: { fill: COLORS.green, font: { color: COLORS.greenDark } } });
errors.getRange("N4:N153").conditionalFormats.add("containsText", { text: "No", format: { fill: COLORS.red, font: { color: COLORS.redDark } } });
errors.freezePanes.freezeRows(3);
errors.showGridLines = false;
setWidths(errors, [14, 25, 28, 16, 21, 12, 13, 34, 34, 28, 13, 15, 22, 12]);
errors.getRange("B4:N153").format.wrapText = true;

// Weekly Overview
styleTitle(weekly, "A1:J1", "Weekly Overview", "Weekly completion governs the plan. Daily assignments may move within the same week, especially around school demands and full-length exams.");
weekly.getRange("A3:J3").values = [["Week", "Phase", "Primary Priority", "Exact Deliverables", "Planned Hours", "Completed Hours", "Target Questions", "Completed Questions", "Full-Length", "Catch-up / Flex Rule"]];
styleHeader(weekly.getRange("A3:J3"));
const weekStarts = [];
for (let d = studyStart; d <= studyEnd; d = addDays(d, 7)) weekStarts.push(new Date(d));
const priorities = [
  "Reset systems; baseline weak-topic inventory", "Bio/Biochem foundations + CARS mapping", "Chem/Phys foundations + CARS evidence", "Practice Exam 1 + complete review",
  "Repair Exam 1 gaps", "Timed passage execution", "Mixed section timing", "Practice Exam 2 + complete review",
  "Repair Exam 2 gaps", "Mixed timed sets + stamina", "Recurring-error remediation", "Practice Exam 3 + complete review",
  "Fresh-score repair", "Exam-like mixed sets", "Practice Exam 4 + complete review", "Focused repair; reduce broad review",
  "Practice Exam 5 + holiday flexibility", "Consolidation and error-log revisits", "Practice Exam 6 + taper",
];
const weeklyRows = weekStarts.map((start, idx) => {
  const weekIso = isoDate(start);
  const fl = flWeeks.get(weekIso);
  const end = addDays(start, 4);
  const phase = phaseFor(start);
  const deliverable = fl
    ? `Take Practice Exam ${fl.exam} Tuesday after Epidemiology; review it Wednesday around classes; finish remediation before heavy new work.`
    : idx === 18
      ? "Complete final light review, equations, familiar CARS, and sleep stabilization; stop heavy work by 6 January."
      : "Complete all dated AAMC sessions, review every flagged response, update the error log, finish the protected weekday Anki session, and keep Friday to two light MCAT hours.";
  return [
    start, phase, priorities[idx] || phase, deliverable, null, null, null, null,
    fl ? `Practice Exam ${fl.exam} — Tuesday` : "—",
    fl ? "Tuesday is the fixed exam day; Wednesday is the review day." : "Move tasks within the week; unfinished work uses the flexible catch-up block before rolling forward.",
  ];
});
weekly.getRange(`A4:J${weeklyRows.length + 3}`).values = weeklyRows;
for (let i = 0; i < weeklyRows.length; i++) {
  const row = i + 4;
  const start = weekStarts[i];
  const end = addDays(start, 4);
  const startIso = `${start.getUTCFullYear()},${start.getUTCMonth() + 1},${start.getUTCDate()}`;
  const endIso = `${end.getUTCFullYear()},${end.getUTCMonth() + 1},${end.getUTCDate()}`;
  weekly.getRange(`E${row}`).formulas = [[`=SUMIFS('Daily Schedule'!$F$4:$F$${dailyRows.length + 3},'Daily Schedule'!$A$4:$A$${dailyRows.length + 3},">="&DATE(${startIso}),'Daily Schedule'!$A$4:$A$${dailyRows.length + 3},"<"&DATE(${endIso})+1)`]];
  weekly.getRange(`F${row}`).formulas = [[`=SUMIFS('Daily Schedule'!$F$4:$F$${dailyRows.length + 3},'Daily Schedule'!$A$4:$A$${dailyRows.length + 3},">="&DATE(${startIso}),'Daily Schedule'!$A$4:$A$${dailyRows.length + 3},"<"&DATE(${endIso})+1,'Daily Schedule'!$H$4:$H$${dailyRows.length + 3},"Complete")`]];
  weekly.getRange(`G${row}`).formulas = [[`=SUMIFS('Daily Schedule'!$G$4:$G$${dailyRows.length + 3},'Daily Schedule'!$A$4:$A$${dailyRows.length + 3},">="&DATE(${startIso}),'Daily Schedule'!$A$4:$A$${dailyRows.length + 3},"<"&DATE(${endIso})+1)`]];
  weekly.getRange(`H${row}`).formulas = [[`=SUMIFS('Daily Schedule'!$G$4:$G$${dailyRows.length + 3},'Daily Schedule'!$A$4:$A$${dailyRows.length + 3},">="&DATE(${startIso}),'Daily Schedule'!$A$4:$A$${dailyRows.length + 3},"<"&DATE(${endIso})+1,'Daily Schedule'!$H$4:$H$${dailyRows.length + 3},"Complete")`]];
}
weekly.getRange(`A4:A${weeklyRows.length + 3}`).format.numberFormat = '"Week of "dd mmm yyyy';
weekly.getRange(`E4:F${weeklyRows.length + 3}`).format.numberFormat = "0.00";
weekly.getRange(`G4:H${weeklyRows.length + 3}`).format.numberFormat = "0";
weekly.getRange(`E4:E${weeklyRows.length + 3}`).conditionalFormats.add("colorScale", { colors: [COLORS.blue, COLORS.yellow, COLORS.orangeLight] });
styleData(weekly.getRange(`A4:J${weeklyRows.length + 3}`));
weekly.getRange(`B4:J${weeklyRows.length + 3}`).format.wrapText = true;
weekly.freezePanes.freezeRows(3);
weekly.showGridLines = false;
setWidths(weekly, [18, 23, 31, 54, 14, 15, 16, 18, 27, 48]);

// Dashboard
styleTitle(dashboard, "A1:H1", "MCAT Retake Dashboard", "Goal: 515 | Baseline: 490 (C/P 123 • CARS 122 • B/B 122 • P/S 123) | Exam: 9 January 2027");
dashboard.getRange("A4:B4").values = [["Key Metric", "Current Value"]];
styleHeader(dashboard.getRange("A4:B4"));
const dashLabels = [
  ["Exam Date"], ["Days Remaining"], ["Goal Score"], ["Baseline Score"], ["Score Gap"],
  ["Planned Study Hours"], ["Completed Study Hours"], ["Study Completion"],
  ["Official Questions Assigned"], ["Official Questions Completed"], ["Question Completion"],
];
dashboard.getRange("A5:A15").values = dashLabels;
dashboard.getRange("B5").values = [[examDate]];
dashboard.getRange("B6").formulas = [["=MAX(0,INT(B5-TODAY()))"]];
dashboard.getRange("B7").values = [[515]];
dashboard.getRange("B8").values = [[490]];
dashboard.getRange("B9").formulas = [["=B7-B8"]];
dashboard.getRange("B10").formulas = [[`=SUM('Daily Schedule'!$F$4:$F$${dailyRows.length + 3})`]];
dashboard.getRange("B11").formulas = [[`=SUMIF('Daily Schedule'!$H$4:$H$${dailyRows.length + 3},"Complete",'Daily Schedule'!$F$4:$F$${dailyRows.length + 3})`]];
dashboard.getRange("B12").formulas = [["=IFERROR(B11/B10,0)"]];
dashboard.getRange("B13").formulas = [[`=SUM('Resource Inventory'!$B$4:$B$${resources.length + 3})+1380`]];
dashboard.getRange("B14").formulas = [[`=SUM('Resource Inventory'!$G$4:$G$${resources.length + 3})+SUMIFS('Daily Schedule'!$G$4:$G$${dailyRows.length + 3},'Daily Schedule'!$H$4:$H$${dailyRows.length + 3},"Complete",'Daily Schedule'!$C$4:$C$${dailyRows.length + 3},"Practice Exam*")`]];
dashboard.getRange("B15").formulas = [["=IFERROR(B14/B13,0)"]];
dashboard.getRange("B5").format.numberFormat = "dd mmm yyyy";
dashboard.getRange("B6:B11").format.numberFormat = "0";
dashboard.getRange("B12").format.numberFormat = "0%";
dashboard.getRange("B13:B14").format.numberFormat = "0";
dashboard.getRange("B15").format.numberFormat = "0%";
dashboard.getRange("A5:A15").format.fill = COLORS.gray;
dashboard.getRange("A5:A15").format.font = { bold: true, color: COLORS.purpleDark };
dashboard.getRange("B5:B15").format.fill = COLORS.white;
dashboard.getRange("A5:B15").format.borders = { preset: "inside", style: "thin", color: "#DDD6E8" };
dashboard.getRange("B12:B15").conditionalFormats.add("dataBar", { color: COLORS.purple, gradient: true });

dashboard.getRange("D4:H4").merge();
dashboard.getRange("D4:H4").values = [["Full-Length Score Trend"]];
styleHeader(dashboard.getRange("D4:H4"));
dashboard.getRange("D5:E11").values = [["Exam", "Total"], ...Array.from({ length: 6 }, (_, i) => [`Exam ${i + 1}`, null])];
for (let i = 0; i < 6; i++) dashboard.getRange(`E${i + 6}`).formulas = [[`=IF('Full-Length Tracker'!I${i + 4}>0,'Full-Length Tracker'!I${i + 4},"")`]];
dashboard.getRange("D5:E5").format.fill = COLORS.lavender2;
dashboard.getRange("D5:E5").format.font = { bold: true, color: COLORS.purpleDark };
dashboard.getRange("D6:E11").format.borders = { preset: "inside", style: "thin", color: "#E5E2EA" };
dashboard.getRange("E6:E11").conditionalFormats.add("colorScale", { colors: [COLORS.red, COLORS.yellow, COLORS.green] });

dashboard.getRange("A18:H18").merge();
dashboard.getRange("A18:H18").values = [["Section Accuracy from Completed Daily Sessions"]];
styleHeader(dashboard.getRange("A18:H18"));
dashboard.getRange("A19:B23").values = [["Section", "Average Accuracy"], ["Chem/Phys", null], ["CARS", null], ["Bio/Biochem", null], ["Psych/Soc", null]];
for (let row = 20; row <= 23; row++) {
  dashboard.getRange(`B${row}`).formulas = [[`=IFERROR(AVERAGEIFS('Daily Schedule'!$I$4:$I$${dailyRows.length + 3},'Daily Schedule'!$D$4:$D$${dailyRows.length + 3},A${row},'Daily Schedule'!$H$4:$H$${dailyRows.length + 3},"Complete"),"")`]];
}
dashboard.getRange("A19:B19").format.fill = COLORS.lavender2;
dashboard.getRange("A19:B19").format.font = { bold: true, color: COLORS.purpleDark };
dashboard.getRange("B20:B23").format.numberFormat = "0.0%";
dashboard.getRange("B20:B23").conditionalFormats.add("colorScale", { colors: [COLORS.red, COLORS.yellow, COLORS.green] });
dashboard.getRange("D18:H18").merge();
dashboard.getRange("D18:H18").values = [["Operating Rules"]];
styleHeader(dashboard.getRange("D18:H18"));
dashboard.getRange("D19:H23").merge();
dashboard.getRange("D19:H23").values = [["• Classes remain protected; Monday–Thursday carry the intensive MCAT workload.\n• All full lengths are Tuesday after Epidemiology; review is Wednesday around classes.\n• Friday includes Kinesiology, Personal Health, and homework, with only two light MCAT hours.\n• Weekends are protected; optional weekend Anki never affects completion.\n• Use Khan only for diagnosed gaps; use Anki only for durable recall prompts."]];
dashboard.getRange("D19:H23").format.fill = COLORS.lavender;
dashboard.getRange("D19:H23").format.font = { color: COLORS.purpleDark, size: 11 };
dashboard.getRange("D19:H23").format.wrapText = true;
dashboard.getRange("D19:H23").format.verticalAlignment = "top";

dashboard.showGridLines = false;
setWidths(dashboard, [29, 18, 4, 18, 14, 14, 14, 14]);
dashboard.getRange("1:1").format.rowHeight = 36;

try {
  const chart = dashboard.charts.add("line", dashboard.getRange("D5:E11"));
  chart.titleText = "Practice Exam Total Score";
  chart.hasLegend = false;
  chart.setPosition("D12", "H17");
} catch (error) {
  console.warn(`Chart creation skipped: ${error.message}`);
}

// Final table styling and filters
daily.tables.add(`A3:M${dailyRows.length + 3}`, true, "DailyScheduleTable").style = "TableStyleMedium4";
inventory.tables.add(`A3:H${invTotalRow - 1}`, true, "ResourceInventoryTable").style = "TableStyleMedium4";
fullLengths.tables.add("A3:N9", true, "FullLengthTable").style = "TableStyleMedium4";
errors.tables.add("A3:N153", true, "ErrorLogTable").style = "TableStyleMedium4";
weekly.tables.add(`A3:J${weeklyRows.length + 3}`, true, "WeeklyOverviewTable").style = "TableStyleMedium4";
classAvailability.tables.add("A3:G10", true, "ClassAvailabilityTable").style = "TableStyleMedium4";

// Export and compact inspection/render verification artifacts.
await fs.mkdir(outputDir, { recursive: true });
const exportBlob = await SpreadsheetFile.exportXlsx(workbook);
await exportBlob.save(outputPath);

const checks = [];
checks.push((await workbook.inspect({ kind: "table", range: "Dashboard!A1:H23", include: "values,formulas", tableMaxRows: 23, tableMaxCols: 8 })).ndjson);
checks.push((await workbook.inspect({ kind: "table", range: `Resource Inventory!A1:H${invTotalRow}`, include: "values,formulas", tableMaxRows: 20, tableMaxCols: 8 })).ndjson);
checks.push((await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 300 }, summary: "final formula error scan" })).ndjson);
await fs.writeFile(`${outputDir}/verification.txt`, checks.join("\n"));

for (const [sheetName, range] of [
  ["Dashboard", "A1:H23"],
  ["Daily Schedule", "A1:M24"],
  ["Weekly Overview", "A1:J22"],
  ["Full-Length Tracker", "A1:N9"],
  ["Error Log", "A1:N18"],
  ["Resource Inventory", `A1:H${invTotalRow}`],
  ["Class & Availability", "A1:G12"],
]) {
  const preview = await workbook.render({ sheetName, range, scale: 1, format: "png" });
  const bytes = new Uint8Array(await preview.arrayBuffer());
  await fs.writeFile(`${outputDir}/preview_${sheetName.replaceAll(" ", "_")}.png`, bytes);
}

console.log(JSON.stringify({ outputPath, dailyRows: dailyRows.length, studyWeekdays: studyDates.length, resourceTotal: resources.reduce((s, r) => s + r[1], 0), fullLengthQuestions: 1380 }, null, 2));

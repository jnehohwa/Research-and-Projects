import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = path.dirname(fileURLToPath(import.meta.url));
const outputDir = path.join(root, "outputs");
const previewDir = path.join(root, "tracker_previews");
await fs.mkdir(outputDir, { recursive: true });
await fs.mkdir(previewDir, { recursive: true });

const workbook = Workbook.create();
workbook.comments.setSelf({ displayName: "Joshua Nehohwa" });

const pipeline = workbook.worksheets.add("Application Pipeline");
const dashboard = workbook.worksheets.add("Dashboard");
const guidance = workbook.worksheets.add("Guidance");
const recruiters = workbook.worksheets.add("Recruiter Directory");

const navy = "#142437";
const teal = "#007575";
const paleTeal = "#E8F2F2";
const paleAmber = "#FFF4D6";
const paleRed = "#FDE8E7";
const paleGreen = "#E7F5EC";
const grid = "#D7DEE5";

pipeline.showGridLines = false;
pipeline.freezePanes.freezeRows(4);
pipeline.getRange("A1:T1").merge();
pipeline.getRange("A1").values = [["IRELAND SOFTWARE JOB APPLICATION PIPELINE"]];
pipeline.getRange("A1:T1").format = {
  fill: navy,
  font: { bold: true, color: "#FFFFFF", size: 16 },
  verticalAlignment: "center",
};
pipeline.getRange("A1:T1").format.rowHeight = 30;
pipeline.getRange("A2:T2").merge();
pipeline.getRange("A2").values = [["Use one row per vacancy. Confirm duties, salary, contract and work-permit feasibility before investing in a tailored application."]];
pipeline.getRange("A2:T2").format = { fill: paleTeal, font: { color: navy, italic: true }, wrapText: true };
pipeline.getRange("A2:T2").format.rowHeight = 30;

const headers = [
  "Priority", "Employer", "Role", "Location", "Level", "Status", "Date Found", "Closing Date",
  "Start", "Salary EUR", "Role Fit", "Permit Fit", "Sponsorship Evidence", "CV Variant",
  "Next Action", "Follow-up", "Job URL", "Source / Careers URL", "Notes", "Last Checked"
];
pipeline.getRange("A4:T4").values = [headers];
pipeline.getRange("A4:T4").format = {
  fill: teal,
  font: { bold: true, color: "#FFFFFF" },
  wrapText: true,
  verticalAlignment: "center",
  borders: { preset: "all", style: "thin", color: grid },
};
pipeline.getRange("A4:T4").format.rowHeight = 34;

const leads = [
  ["High", "EY Ireland", "Technology Consulting Graduate Programme 2027", "Dublin / Cork", "Graduate", "Tailoring", new Date("2026-08-26"), new Date("2026-10-21"), "Main intake Sep 2027", "", 5, "Strong potential", "EY explicitly welcomes non-EEA graduate applicants and lists Consulting and Technology among permit-supporting programmes", "Graduate software", "Tailor one EY application with Technology Consulting as first preference", new Date("2026-09-02"), "https://eyglobal.yello.co/jobs/gi8_dm-yLjEV3rAaBvar2A?job_board_id=c1riT--B2O-KySgYWsZO1Q", "https://www.ey.com/en_ie/careers/what-you-can-do-here/student-programmes", "Best immediate target. EY allows only one graduate application; choose AI & Data and Assurance Technology Risk as alternative preferences if the form permits.", new Date("2026-08-26")],
  ["High", "Susquehanna (SIG)", "Software Developer Graduate: September 2027", "Dublin", "Graduate", "Research", new Date("2026-08-21"), "", "Sep 2027", "", 5, "Needs verification", "Work-permit support and salary not stated", "Graduate software", "Confirm whether May 2027 ceremony satisfies 'graduating in 2027' and whether permit support is available", new Date("2026-09-02"), "https://careers.sig.com/recent-graduate/jobs/11337?lang=en-us", "https://careers.sig.com/recent-graduate/jobs/11337?lang=en-us", "Official role is live and requires availability in September 2027. Degree completion is Nov 2026 but ceremony is expected May 2027; eligibility needs confirmation.", new Date("2026-08-26")],
  ["High", "EirGrid", "2027 Graduate Programme - IT / Cyber", "Ireland", "Graduate", "Watchlist", new Date("2026-08-18"), new Date("2026-09-18"), "2027", "", 4, "Needs verification", "Applications open 18 Sep 2026; sponsorship not stated", "Graduate software", "Revisit on opening date and verify permit support, salary and start date", new Date("2026-09-18"), "https://www.eirgrid.ie/graduates", "https://www.eirgrid.ie/graduates", "Official page confirms IT and cyber streams and an 18 September 2026 opening.", new Date("2026-08-26")],
  ["Medium", "AIB", "Technology & Data Graduate Programme 2027", "Ireland", "Graduate", "Watchlist", new Date("2026-08-26"), "", "2027", "", 4, "Needs verification", "Official programme page does not currently state permit support or a 2027 opening date", "Graduate software", "Monitor official jobs page through September and assess the live vacancy when posted", new Date("2026-09-15"), "https://www.aib.ie/careers/future-talent-programmes", "https://jobs.aib.ie/content/Graduate/?locale=en_GB&res=1&sub_sector_id=14", "Technology & Data is a verified programme area; application timing and non-EEA eligibility remain unresolved.", new Date("2026-08-26")],
  ["Medium", "Bank of Ireland", "Data, Technology & Change Graduate Programme 2027", "Ireland", "Graduate", "Watchlist", new Date("2026-08-26"), "", "2027", "", 4, "Needs verification", "2027 interest sign-up is live; vacancy and permit support are not yet stated", "Graduate software", "Join the official interest list and monitor for the 2027 vacancy", new Date("2026-09-15"), "https://careers.bankofireland.com/graduates/bank-of-ireland-graduates-programmes-data-tech-change", "https://careers.bankofireland.com/graduates/bank-of-ireland-graduates-programmes-data-tech-change", "Relevant preference areas include Technology Services & Security and Data & Engineering.", new Date("2026-08-26")],
  ["Medium", "PwC Ireland", "2027 Graduate Programme - technology-related stream", "Ireland", "Graduate", "Watchlist", new Date("2026-08-26"), "", "2027", "", 3, "Needs verification", "2027 roles are not yet confirmed on the official graduate page", "Graduate software", "Monitor official graduate listings and check work-permit support in each role", new Date("2026-09-10"), "https://www.pwc.ie/careers-ie/students/graduate-programme/apply.html", "https://www.pwc.ie/careers-ie/students/frequently-asked-questions.html", "Do not use the remaining 2026 listings; wait for a clearly labelled 2027 technology role.", new Date("2026-08-26")],
  ["Medium", "KPMG Ireland", "2027 Graduate Programme - STEM / Technology", "Ireland", "Graduate", "Watchlist", new Date("2026-08-26"), "", "2027", "", 3, "Needs verification", "Official STEM page says graduate applications open each autumn; 2027 vacancy not yet verified", "Graduate software", "Monitor the official graduate page in September", new Date("2026-09-15"), "https://kpmg.com/ie/en/careers/stem-opportunities.html", "https://kpmg.com/ie/en/careers/graduate.html", "Potential technology-risk or consulting route; not yet a live verified 2027 vacancy.", new Date("2026-08-26")],
  ["Medium", "Deloitte Ireland", "NextGen Tech Academy / 2027 technology intake", "Dublin", "Graduate / academy", "Watchlist", new Date("2026-08-26"), "", "Not stated", "", 4, "Needs verification", "Programme page does not state the next intake date or permit support", "Graduate software", "Monitor for a 2027 intake; do not apply until timing and eligibility are confirmed", new Date("2026-09-15"), "https://www.deloitte.com/ie/en/careers/explore-your-fit/professional-careers/nextgen-ai-academy.html", "https://www.deloitte.com/ie/en/careers.html", "Strong AWS/cloud/AI alignment, but current timing is unresolved.", new Date("2026-08-26")],
  ["Medium", "Esri Ireland", "Graduate Programme", "Dublin", "Graduate", "Watchlist", new Date("2026-08-18"), new Date("2026-12-31"), "2027 intake", "", 3, "Likely ineligible", "Requires degree from an institution in Republic of Ireland or Northern Ireland", "Graduate software", "Do not apply unless eligibility wording changes", new Date("2026-09-01"), "https://www.esri-ireland.ie/en-ie/about/careers/graduate-program", "https://www.esri-ireland.ie/en-ie/about/careers/graduate-program", "Current published requirement appears to exclude overseas degrees.", new Date("2026-08-18")],
  ["Low", "Ryanair", "Software Development Graduate Programme", "Dublin", "Graduate", "Not eligible", new Date("2026-08-18"), "", "2026", "", 4, "Not eligible", "Requires unrestricted right to live and work in Ireland", "Graduate software", "Archive; do not spend tailoring time", "", "https://careers.ryanair.com/jobs/software-development-graduate-programme/", "https://careers.ryanair.com/jobs/software-development-graduate-programme/", "Technically relevant but fails stated work-right requirement.", new Date("2026-08-18")],
];

pipeline.getRange(`A5:T${4 + leads.length}`).values = leads;
pipeline.getRange(`A5:T${4 + leads.length}`).format = {
  wrapText: true,
  verticalAlignment: "top",
  borders: { preset: "all", style: "thin", color: grid },
};
pipeline.getRange(`G5:G${4 + leads.length}`).setNumberFormat("yyyy-mm-dd");
pipeline.getRange(`H5:H${4 + leads.length}`).setNumberFormat("yyyy-mm-dd");
pipeline.getRange(`P5:P${4 + leads.length}`).setNumberFormat("yyyy-mm-dd");
pipeline.getRange(`T5:T${4 + leads.length}`).setNumberFormat("yyyy-mm-dd");
pipeline.getRange(`J5:J${4 + leads.length}`).setNumberFormat('€#,##0');

pipeline.getRange("A5:A200").dataValidation = { rule: { type: "list", values: ["High", "Medium", "Low"] } };
pipeline.getRange("F5:F200").dataValidation = { rule: { type: "list", values: ["Watchlist", "Research", "Tailoring", "Ready for review", "Approved", "Submitted", "Interview", "Offer", "Rejected", "Withdrawn", "Not eligible", "Closed / timing mismatch"] } };
pipeline.getRange("K5:K200").dataValidation = { rule: { type: "whole", operator: "between", formula1: 1, formula2: 5 } };
pipeline.getRange("L5:L200").dataValidation = { rule: { type: "list", values: ["Strong potential", "Needs verification", "Likely ineligible", "Not eligible"] } };
pipeline.getRange("N5:N200").dataValidation = { rule: { type: "list", values: ["Master full-stack", "Full-stack", "Graduate software", "Frontend", "Backend"] } };

pipeline.getRange("A5:A200").conditionalFormats.add("containsText", { text: "High", format: { fill: paleGreen, font: { bold: true, color: "#176B3A" } } });
pipeline.getRange("F5:F200").conditionalFormats.add("containsText", { text: "Not eligible", format: { fill: paleRed, font: { color: "#9B1C1C" } } });
pipeline.getRange("F5:F200").conditionalFormats.add("containsText", { text: "Submitted", format: { fill: paleGreen, font: { color: "#176B3A", bold: true } } });
pipeline.getRange("L5:L200").conditionalFormats.add("containsText", { text: "Needs verification", format: { fill: paleAmber } });

const widths = [10, 19, 29, 15, 12, 18, 12, 12, 14, 12, 9, 17, 25, 18, 30, 12, 43, 38, 40, 12];
for (let i = 0; i < widths.length; i += 1) pipeline.getRangeByIndexes(0, i, 200, 1).format.columnWidth = widths[i];
pipeline.getRange(`A5:T${4 + leads.length}`).format.autofitRows();
pipeline.tables.add(`A4:T${4 + leads.length}`, true, "ApplicationsTable").style = "TableStyleMedium2";

dashboard.showGridLines = false;
dashboard.getRange("A1:H1").merge();
dashboard.getRange("A1").values = [["JOB SEARCH DASHBOARD"]];
dashboard.getRange("A1:H1").format = { fill: navy, font: { bold: true, color: "#FFFFFF", size: 16 } };
dashboard.getRange("A1:H1").format.rowHeight = 30;
dashboard.getRange("A3:B9").values = [
  ["Metric", "Count"],
  ["Total leads", ""],
  ["High priority", ""],
  ["Needs permit verification", ""],
  ["Ready for review", ""],
  ["Submitted", ""],
  ["Interviews", ""],
];
dashboard.getRange("B4").formulas = [["=COUNTA('Application Pipeline'!B5:B200)"]];
dashboard.getRange("B5").formulas = [["=COUNTIF('Application Pipeline'!A5:A200,\"High\")"]];
dashboard.getRange("B6").formulas = [["=COUNTIF('Application Pipeline'!L5:L200,\"Needs verification\")"]];
dashboard.getRange("B7").formulas = [["=COUNTIF('Application Pipeline'!F5:F200,\"Ready for review\")"]];
dashboard.getRange("B8").formulas = [["=COUNTIF('Application Pipeline'!F5:F200,\"Submitted\")"]];
dashboard.getRange("B9").formulas = [["=COUNTIF('Application Pipeline'!F5:F200,\"Interview\")"]];
dashboard.getRange("A3:B3").format = { fill: teal, font: { bold: true, color: "#FFFFFF" } };
dashboard.getRange("A3:B9").format.borders = { preset: "all", style: "thin", color: grid };
dashboard.getRange("A4:A9").format.font = { bold: true, color: navy };
dashboard.getRange("A3:B9").format.columnWidth = 24;
dashboard.getRange("B4:B9").format.numberFormat = "0";
dashboard.getRange("D3:H3").merge();
dashboard.getRange("D3").values = [["NEXT MILESTONES"]];
dashboard.getRange("D3:H3").format = { fill: teal, font: { bold: true, color: "#FFFFFF" } };
dashboard.getRange("D4:H8").merge(true);
dashboard.getRange("D4:H8").values = [
  ["1. Degree requirements complete 10 Nov 2026; earliest Ireland start Apr 2027."],
  ["2. Contact licensed tech recruiters and replace watchlist rows with live vacancies."],
  ["3. Prepare for EirGrid's 18 September 2026 opening."],
  ["4. Verify each role's duties and basic salary against current permit rules."],
  ["5. Tailor, review, approve, submit, and record confirmation."],
];
dashboard.getRange("D4:H8").format = { fill: paleTeal, wrapText: true, font: { color: navy }, borders: { preset: "all", style: "thin", color: grid } };
dashboard.getRange("D4:H8").format.rowHeight = 32;
dashboard.getRange("D1:H12").format.columnWidth = 17;

guidance.showGridLines = false;
guidance.freezePanes.freezeRows(3);
guidance.getRange("A1:D1").merge();
guidance.getRange("A1").values = [["APPLICATION RULES AND DEFINITIONS"]];
guidance.getRange("A1:D1").format = { fill: navy, font: { bold: true, color: "#FFFFFF", size: 16 } };
guidance.getRange("A3:D3").values = [["Area", "Rule", "Evidence required", "Official source"]];
guidance.getRange("A3:D3").format = { fill: teal, font: { bold: true, color: "#FFFFFF" }, wrapText: true };
const rules = [
  ["Permit assessment", "Assess actual duties and SOC classification; job title alone is insufficient.", "Full vacancy, contract duties, salary and qualifications.", "https://enterprise.gov.ie/en/what-we-do/workplace-and-skills/employment-permits/employment-permit-eligibility/classification-of-employments/"],
  ["Critical Skills list", "Programmers/software developers and web design/development professionals are listed as of August 2026.", "Current government list checked on application date.", "https://enterprise.gov.ie/en/what-we-do/workplace-and-skills/employment-permits/employment-permit-eligibility/highly-skilled-eligible-occupations-list/"],
  ["Salary", "Use basic annual salary only and verify the current threshold before applying.", "Published salary or recruiter confirmation.", "https://enterprise.gov.ie/en/news-and-events/department-news/2025/december/20251202.html"],
  ["Work authorization", "Never state that Joshua has unrestricted Irish work rights unless documentary status changes.", "Current immigration status confirmed by Joshua.", "https://enterprise.gov.ie/en/what-we-do/workplace-and-skills/employment-permits/"],
  ["Application approval", "No submission before Joshua reviews the final CV, letter and screening answers.", "Explicit approval and submission confirmation.", "Internal workflow"],
  ["Evidence integrity", "Do not convert prototypes, placeholders or planned features into production claims.", "Repository, deployment, document, certificate or stakeholder evidence.", "Internal workflow"],
];
guidance.getRange(`A4:D${3 + rules.length}`).values = rules;
guidance.getRange(`A3:D${3 + rules.length}`).format = { wrapText: true, verticalAlignment: "top", borders: { preset: "all", style: "thin", color: grid } };
guidance.getRange("A:A").format.columnWidth = 20;
guidance.getRange("B:B").format.columnWidth = 44;
guidance.getRange("C:C").format.columnWidth = 38;
guidance.getRange("D:D").format.columnWidth = 55;
guidance.getRange(`A4:D${3 + rules.length}`).format.autofitRows();

recruiters.showGridLines = false;
recruiters.freezePanes.freezeRows(4);
recruiters.getRange("A1:J1").merge();
recruiters.getRange("A1").values = [["VERIFIED IRISH TECHNOLOGY RECRUITER DIRECTORY"]];
recruiters.getRange("A1:J1").format = { fill: navy, font: { bold: true, color: "#FFFFFF", size: 16 } };
recruiters.getRange("A1:J1").format.rowHeight = 30;
recruiters.getRange("A2:J2").merge();
recruiters.getRange("A2").values = [["Contact only through the official site or listed domain. Ask whether the employer can consider a non-EEA graduate requiring an employment permit; no recruiter can guarantee sponsorship."]];
recruiters.getRange("A2:J2").format = { fill: paleAmber, font: { color: navy, italic: true }, wrapText: true };
recruiters.getRange("A2:J2").format.rowHeight = 34;
const recruiterHeaders = ["Priority", "Agency", "Technology focus", "Irish base", "Best first contact", "Official contact", "WRC legal entity", "Licence", "Outreach status", "Last checked"];
recruiters.getRange("A4:J4").values = [recruiterHeaders];
recruiters.getRange("A4:J4").format = { fill: teal, font: { bold: true, color: "#FFFFFF" }, wrapText: true, borders: { preset: "all", style: "thin", color: grid } };
recruiters.getRange("A4:J4").format.rowHeight = 32;
const recruiterRows = [
  ["High", "IT Search", "Software development, cloud, data and IT", "Dublin", "Graeme King, software-development recruiter: graeme.king@itsearch.ie | +353 87 4998421", "https://itsearch.ie/team/graeme-king/", "Vertical Markets Group Ltd", "EA 4262", "Contacted", new Date("2026-08-20")],
  ["High", "Stelfox", "Technology and digital recruitment", "Dublin, Cork, Galway", "dublin@stelfox.com | +353 1 679 3182", "https://www.stelfox.com/branches", "Stelfox Ltd", "EA 2059", "Contacted", new Date("2026-08-24")],
  ["High", "GemPool", "Technology recruitment", "Dublin", "info@gempool.ie | +353 1 913 6804", "https://www.gempool.ie/contact/", "GemPool Ltd", "EA 3546", "Contacted", new Date("2026-08-20")],
  ["High", "Archer Recruitment", "IT, data and change", "Dublin", "info@archer.ie | +353 1 649 8500", "https://www.archer.ie/contact/", "Recport Ltd", "EA 3658", "Contacted", new Date("2026-08-24")],
  ["Medium", "Morgan McKinley", "Technology recruitment", "Dublin", "Technology team | +353 1 432 1567", "https://www.morganmckinley.com/ie/about/contact-us/dublin", "Premier Recruitment International Ltd", "EA 1080", "Not contacted", new Date("2026-08-20")],
  ["Medium", "Hays", "Software development and technology", "Dublin, Cork, Galway, Limerick", "Dublin office | +353 1 571 0010", "https://www.hays.ie/contact/offices-near-me", "Hays Specialist Recruitment (Ireland) Ltd", "EA 1273", "Not contacted", new Date("2026-08-20")],
  ["Medium", "Reperio Human Capital", "IT recruitment", "Dublin", "Official contact form", "https://www.reperiohumancapital.com/contact-us", "Reperio Human Capital (Ireland) Ltd", "EA 4024", "Not contacted", new Date("2026-08-20")],
  ["Medium", "Berkley Group", "Software engineering, digital and IT", "Cork", "Technology recruitment page / consultant directory", "https://www.berkley-group.com/digital-technology-it-recruitment/", "Berkley Recruitment (Group) Ltd", "EA 1260", "Not contacted", new Date("2026-08-20")],
];
recruiters.getRange(`A5:J${4 + recruiterRows.length}`).values = recruiterRows;
recruiters.getRange(`A5:J${4 + recruiterRows.length}`).format = { wrapText: true, verticalAlignment: "top", borders: { preset: "all", style: "thin", color: grid } };
recruiters.getRange(`J5:J${4 + recruiterRows.length}`).setNumberFormat("yyyy-mm-dd");
recruiters.getRange("A5:A200").dataValidation = { rule: { type: "list", values: ["High", "Medium", "Low"] } };
recruiters.getRange("I5:I200").dataValidation = { rule: { type: "list", values: ["Not contacted", "Message drafted", "Contacted", "Replied", "Follow-up due", "Closed"] } };
const recruiterWidths = [10, 21, 29, 20, 45, 48, 36, 12, 18, 13];
for (let i = 0; i < recruiterWidths.length; i += 1) recruiters.getRangeByIndexes(0, i, 200, 1).format.columnWidth = recruiterWidths[i];
recruiters.getRange(`A5:J${4 + recruiterRows.length}`).format.autofitRows();
recruiters.tables.add(`A4:J${4 + recruiterRows.length}`, true, "RecruitersTable").style = "TableStyleMedium2";

const keyThread = workbook.comments.addThread({ cell: pipeline.getRange("L4") }, "Permit fit is a triage field, not legal advice. Re-check the current government rules against the complete vacancy before applying.");
keyThread.addReply("Use 'Needs verification' until duties, salary and contract are known.");

for (const [sheetName, range, filename] of [
  ["Dashboard", "A1:H12", "dashboard.png"],
  ["Application Pipeline", `A1:T${4 + leads.length}`, "pipeline.png"],
  ["Guidance", `A1:D${3 + rules.length}`, "guidance.png"],
  ["Recruiter Directory", `A1:J${4 + recruiterRows.length}`, "recruiters.png"],
]) {
  const preview = await workbook.render({ sheetName, range, scale: 1.25, format: "png" });
  await fs.writeFile(path.join(previewDir, filename), new Uint8Array(await preview.arrayBuffer()));
}

const check = await workbook.inspect({ kind: "table", range: "Dashboard!A1:H12", include: "values,formulas", tableMaxRows: 15, tableMaxCols: 10, maxChars: 5000 });
console.log(check.ndjson);
const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, summary: "final formula error scan", maxChars: 3000 });
console.log(errors.ndjson);

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(path.join(outputDir, "Joshua_Nehohwa_Ireland_Job_Tracker.xlsx"));
console.log(path.join(outputDir, "Joshua_Nehohwa_Ireland_Job_Tracker.xlsx"));

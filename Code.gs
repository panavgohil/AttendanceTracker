/* STATELINE BACKEND: Auth, Attendance, Schedule, Profiles, Weekly PDF Reports */

/* =========================================================================
 *  WHAT'S NEW vs your previous version:
 *  1. "send-report" doPost action + weekly time-trigger -> emails a PDF
 *     attendance summary to opted-in students (see REPORT section at bottom).
 *  2. updateUserProfile() now also stores a ReportOptIn flag in the
 *     previously-unused 7th column of "Users".
 *  3. handleLogin() now also returns email + reportOptIn so a fresh login
 *     on a new device shows the right toggle state immediately.
 *  4. getHistoryJSON() now includes "date" per entry — needed by the new
 *     dashboard's activity heatmap and streak counter, which group your
 *     history by day. This was previously omitted.
 *  Everything else is untouched from your original script.
 * ========================================================================= */

const REPORT_CONFIG = {
  TARGET_PCT: 75, // used only for the "0/1000" colour styling of the emailed PDF
};

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    let result = {};

    if (data.type === "login") result = handleLogin(data);
    else if (data.type === "save-schedule") result = saveUserSchedule(data);
    else if (data.type === "update-profile") result = updateUserProfile(data);
    else if (data.type === "reg-send-otp") result = sendRegistrationOtp(data);
    else if (data.type === "reg-complete") result = completeRegistration(data);
    else if (data.type === "forgot-send-otp") result = sendForgotOtp(data);
    else if (data.type === "forgot-complete") result = completeReset(data);
    else if (data.type === "send-report") result = handleSendReport(data);
    else result = saveAttendance(data);

    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  if (e.parameter.type === "history") return getHistoryJSON(e.parameter.roll);
  if (e.parameter.roll) return exportUserData(e.parameter.roll);
  return ContentService.createTextOutput("Stateline API Running");
}

/* === PROFILE MANAGEMENT === */
function updateUserProfile(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1].toString().toLowerCase() === String(data.roll).toLowerCase()) {
      if (data.password) sheet.getRange(i + 1, 4).setValue(data.password);
      if (data.branch) sheet.getRange(i + 1, 5).setValue(data.branch);
      if (typeof data.reportOptIn !== "undefined") sheet.getRange(i + 1, 7).setValue(data.reportOptIn ? "TRUE" : "FALSE");
      return { status: "success", message: "Profile Updated" };
    }
  }
  return { status: "error", message: "User not found" };
}

/* === SCHEDULES & AUTH === */
function saveUserSchedule(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Schedules");
  if (!sheet) { sheet = ss.insertSheet("Schedules"); sheet.appendRow(["Roll", "ScheduleJSON", "LastUpdated"]); }
  const rows = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0].toString().toLowerCase() === String(data.roll).toLowerCase()) { rowIndex = i + 1; break; }
  }
  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 2).setValue(data.schedule);
    sheet.getRange(rowIndex, 3).setValue(new Date());
  } else {
    sheet.appendRow([data.roll, data.schedule, new Date()]);
  }
  return { status: "success", message: "Schedule Synced" };
}

function getUserSchedule(roll) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Schedules");
  if (!sheet) return null;
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0].toString().toLowerCase() === String(roll).toLowerCase()) return rows[i][1];
  }
  return null;
}

function handleLogin(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const rows = sheet.getDataRange().getValues();
  const id = String(data.identifier).toLowerCase();

  for (let i = 1; i < rows.length; i++) {
    if ((rows[i][1].toString().toLowerCase() === id || rows[i][2].toString().toLowerCase() === id) && rows[i][3].toString() === data.password) {
      return {
        status: "success",
        name: rows[i][0],
        roll: rows[i][1],
        email: rows[i][2],
        branch: rows[i][4],
        reportOptIn: String(rows[i][6]).toUpperCase() === "TRUE",
        schedule: getUserSchedule(rows[i][1])
      };
    }
  }
  return { status: "error", message: "Invalid Credentials" };
}

/* === REGISTRATION & UTILS === */
function sendRegistrationOtp(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1].toString().toLowerCase() === String(data.roll).toLowerCase()) return { status: "error", message: "Roll exists" };
    if (rows[i][2].toString().toLowerCase() === String(data.email).toLowerCase()) return { status: "error", message: "Email exists" };
  }
  sendEmailOtp(data.email, "Verification Code");
  return { status: "otp_sent", message: "OTP Sent" };
}

function completeRegistration(data) {
  if (!verifyOtp(data.email, data.otp)) return { status: "error", message: "Invalid OTP" };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  // 7th column defaults to "FALSE" -> weekly report opt-in, off until the student turns it on
  sheet.appendRow([data.name, data.roll, data.email, data.password, data.branch, new Date(), "FALSE"]);
  return { status: "success", message: "Registered" };
}

function sendForgotOtp(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][2].toString().toLowerCase() === String(data.email).toLowerCase()) {
      sendEmailOtp(data.email, "Reset Password");
      return { status: "otp_sent", message: "OTP Sent" };
    }
  }
  return { status: "error", message: "Email not found" };
}

function completeReset(data) {
  if (!verifyOtp(data.email, data.otp)) return { status: "error", message: "Invalid OTP" };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][2].toString().toLowerCase() === String(data.email).toLowerCase()) {
      sheet.getRange(i + 1, 4).setValue(data.newPassword);
      return { status: "success", message: "Password updated" };
    }
  }
  return { status: "error", message: "User not found" };
}

function sendEmailOtp(email, subject) {
  const otp = Math.floor(100000 + Math.random() * 900000).toString();
  CacheService.getScriptCache().put("OTP_" + email.toLowerCase(), otp, 600);
  MailApp.sendEmail({ to: email, subject: `${subject} - Stateline`, body: `Your code is: ${otp}` });
}

function verifyOtp(email, inputOtp) {
  const cached = CacheService.getScriptCache().get("OTP_" + email.toLowerCase());
  if (cached && cached === inputOtp.toString()) { CacheService.getScriptCache().remove("OTP_" + email.toLowerCase()); return true; }
  return false;
}

/* === ATTENDANCE (defaults to Semester 2) === */
function saveAttendance(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance");
  sheet.appendRow([data.name, data.roll, data.date, data.day, data.subject, data.status, data.remark, data.weight || 1, new Date(), data.semester || "Semester 2"]);
  return { status: "success", message: "Saved" };
}

/** Sheet dates sometimes come back as JS Date objects once Sheets auto-parses the
 *  string date() sends; normalize both cases to the same "M/d/yyyy" shape used by
 *  the front-end's `new Date().toLocaleDateString()` so heatmap/streak keys match. */
function formatDateCell_(d) {
  if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), "M/d/yyyy");
  return d;
}

function getHistoryJSON(roll) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance");
  const rows = sheet.getDataRange().getValues();
  const history = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1].toString().toLowerCase() === String(roll).toLowerCase()) {
      history.push({
        date: formatDateCell_(rows[i][2]),
        subject: rows[i][4],
        status: rows[i][5],
        weight: rows[i][7],
        semester: rows[i][9] || "Semester 2"
      });
    }
  }
  return ContentService.createTextOutput(JSON.stringify(history)).setMimeType(ContentService.MimeType.JSON);
}

function exportUserData(roll) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance");
  const rows = sheet.getDataRange().getValues();
  let csv = "Date,Day,Subject,Status,Remark,Semester\n";
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1].toString().toLowerCase() === String(roll).toLowerCase()) {
      const d = formatDateCell_(rows[i][2]);
      const sem = rows[i][9] || "Semester 2";
      csv += `"${d}","${rows[i][3]}","${rows[i][4]}","${rows[i][5]}","${rows[i][6]}","${sem}"\n`;
    }
  }
  return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV).downloadAsFile(`Stateline_Logs_${String(roll).replace(/[\/\\]/g, "_")}.csv`);
}

/* =========================================================================
 *  WEEKLY PDF ATTENDANCE REPORTS
 * ========================================================================= */

/** Run this ONCE from the Apps Script editor (select it in the function
 *  dropdown, click Run) to install the Sunday 6pm weekly trigger. Re-running
 *  it is safe — it clears any previous trigger of the same name first. */
function createWeeklyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "sendWeeklyReports") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("sendWeeklyReports").timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(18).create();
}

/** Scheduled entry point: emails every opted-in student a PDF covering the
 *  last 7 days of attendance, across all their semesters. */
function sendWeeklyReports() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const rows = sheet.getDataRange().getValues();
  const attendance = getAllAttendanceRows_();

  for (let i = 1; i < rows.length; i++) {
    const optedIn = String(rows[i][6]).toUpperCase() === "TRUE";
    const email = rows[i][2];
    if (!optedIn || !email) continue;
    const student = { name: rows[i][0], roll: rows[i][1], email: email };
    try {
      const rollRows = attendance.filter(r => String(r[1]).toLowerCase() === String(student.roll).toLowerCase());
      const pdf = buildReportPdf_(student, rollRows, 7, null);
      emailReport_(student, pdf, "Your weekly attendance report");
    } catch (err) {
      console.error("Weekly report failed for " + student.roll + ": " + err);
    }
  }
}

/** On-demand version behind the "Email me this week's report" button.
 *  Sends regardless of opt-in status since the student explicitly asked. */
function handleSendReport(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1].toString().toLowerCase() === String(data.roll).toLowerCase()) {
      const student = { name: rows[i][0], roll: rows[i][1], email: rows[i][2] };
      if (!student.email) return { status: "error", message: "No email on file" };
      const attendance = getAllAttendanceRows_().filter(r => String(r[1]).toLowerCase() === String(student.roll).toLowerCase());
      const pdf = buildReportPdf_(student, attendance, 7, data.semester || null);
      emailReport_(student, pdf, "Your attendance report");
      return { status: "success", message: "Report emailed" };
    }
  }
  return { status: "error", message: "Student not found" };
}

function getAllAttendanceRows_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance");
  return sheet.getDataRange().getValues().slice(1); // skip header row
}

/** Builds a one-page PDF summary. Attendance row layout (0-indexed):
 *  0 Name, 1 Roll, 2 Date, 3 Day, 4 Subject, 5 Status, 6 Remark, 7 Weight, 8 Timestamp, 9 Semester */
function buildReportPdf_(student, rows, lastNDays, semesterFilter) {
  const cutoff = new Date(); cutoff.setDate(cutoff.getDate() - lastNDays);
  const recent = rows.filter(r => {
    if (semesterFilter && (r[9] || "Semester 2") !== semesterFilter) return false;
    const d = r[2] instanceof Date ? r[2] : new Date(r[2]);
    return !isNaN(d) && d >= cutoff;
  });

  const subjectStats = {};
  recent.forEach(r => {
    const sub = r[4], status = r[5], weight = parseInt(r[7], 10) || 1;
    if (!subjectStats[sub]) subjectStats[sub] = { attended: 0, total: 0 };
    if (status === "Attended") { subjectStats[sub].attended += weight; subjectStats[sub].total += weight; }
    else if (status === "Absent" || status === "Bunked") { subjectStats[sub].total += weight; }
  });

  let rowsHtml = "";
  Object.keys(subjectStats).forEach(sub => {
    const s = subjectStats[sub];
    const pct = s.total ? Math.round((s.attended / s.total) * 100) : 0;
    const color = pct >= REPORT_CONFIG.TARGET_PCT ? "#16a34a" : (pct >= 50 ? "#d97706" : "#dc2626");
    rowsHtml += `<tr><td style="padding:8px;border-bottom:1px solid #eee;">${sub}</td><td style="padding:8px;border-bottom:1px solid #eee;">${s.attended}/${s.total}</td><td style="padding:8px;border-bottom:1px solid #eee;color:${color};font-weight:bold;">${pct}%</td></tr>`;
  });
  if (!rowsHtml) rowsHtml = `<tr><td colspan="3" style="padding:8px;text-align:center;color:#888;">No attendance recorded this week.</td></tr>`;

  const html = `
    <html><body style="font-family:Arial, sans-serif; padding:24px; color:#111;">
      <h2 style="margin-bottom:0;">Stateline — Weekly Attendance Report</h2>
      <p style="color:#666; margin-top:4px;">${student.name} (${student.roll}) &middot; Week ending ${new Date().toLocaleDateString()}</p>
      <table style="width:100%; border-collapse:collapse; margin-top:16px;">
        <thead><tr style="background:#f1f5f9;">
          <th style="text-align:left; padding:8px; border-bottom:1px solid #ddd;">Subject</th>
          <th style="text-align:left; padding:8px; border-bottom:1px solid #ddd;">Attended/Total (pts)</th>
          <th style="text-align:left; padding:8px; border-bottom:1px solid #ddd;">%</th>
        </tr></thead>
        <tbody>${rowsHtml}</tbody>
      </table>
      <p style="color:#999; font-size:11px; margin-top:24px;">Generated automatically by Stateline. Targets assume a ${REPORT_CONFIG.TARGET_PCT}% minimum.</p>
    </body></html>`;

  const blob = HtmlService.createHtmlOutput(html).getAs("application/pdf");
  blob.setName(`Stateline_Report_${student.roll}_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd")}.pdf`);
  return blob;
}

function emailReport_(student, pdfBlob, subject) {
  MailApp.sendEmail({
    to: student.email,
    subject: subject,
    body: `Hi ${String(student.name).split(" ")[0]},\n\nYour attendance report is attached.\n\n— Stateline`,
    attachments: [pdfBlob],
  });
}

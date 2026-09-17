/**
 * FIESTA IX Student Portal — Version 1
 * Google Apps Script server code
 *
 * Features:
 * - Creates the Google Sheets database tabs
 * - Pairs one office tablet as the authorized kiosk
 * - Student PIN authentication
 * - Clock in / clock out
 * - Prevents duplicate clock-ins
 * - Calculates monthly hours
 * - Uses LockService to prevent simultaneous write collisions
 */

const SHEETS = {
  STUDENTS: 'Students',
  TIME_ENTRIES: 'Time Entries',
  DEVICES: 'Authorized Devices',
  SETTINGS: 'Settings'
};

function doGet(e) {
  const page =
    (e && e.parameter && e.parameter.page) || 'kiosk';

  let file = 'Index';
  let title = 'FIESTA IX Time Clock';

if (page === 'attendancepreview') {
  file = 'MonthlyAttendancePreview';
  title = 'FIESTA Hub - Monthly Attendance';
}

if (page === 'attendanceeditor') {
  file = 'AttendanceEditors';
  title = 'FIESTA Hub - Attendance Editor';
}

if (page === 'schedulerequest') {
  file = 'ScheduleRequestss';
  title = 'FIESTA Hub - Schedule Request';
}

if (page === 'shiftadjustment') {
  file = 'PayrollShiftAdjustmentForm';
  title = 'FIESTA Hub - Manual Shift Adjustment';
}

if (page === 'finalpayroll') {
  file = 'FinalPayrollReports';
  title = 'FIESTA Hub - Final Stipend Report';
}

if (page === 'payroll') {
  file = 'PayrollCenter';
  title = 'FIESTA Hub - Payroll Center';
}

if (page === 'missioncontrol') {
  file = 'DashboardHome';
  title = 'FIESTA Hub - Mission Control';
}

  if (page === 'admin') {
    file = 'Admin';
    title = 'FIESTA Hub';
  }

  if (page === 'studentwizard') {
    file = 'StudentWizards';
    title = 'FIESTA Hub - New Student';
  }

if (page === 'scheduleapprovals') {
  file = 'ScheduleApprovals';
  title = 'FIESTA Hub - Schedule Approval Center';
}

if (page === 'studentportal') {

  file = 'StudentPortal';

  title = 'FIESTA IX - Student Portal';

}

  return HtmlService
    .createTemplateFromFile(file)
    .evaluate()
    .setTitle(title)
    .addMetaTag(
      'viewport',
      'width=device-width, initial-scale=1'
    );
}

/**
 * Run this ONCE from the Apps Script editor.
 * It creates all required sheets and basic settings.
 */
function setupPortal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  createOrResetHeaders_(ss, SHEETS.STUDENTS, [
    'Student ID', 'Full Name', 'Email', 'PIN Salt', 'PIN Hash', 'Active'
  ]);

  createOrResetHeaders_(ss, SHEETS.TIME_ENTRIES, [
    'Entry ID', 'Student ID', 'Student Name', 'Clock In',
    'Clock Out', 'Total Hours', 'Device ID', 'Status', 'Notes'
  ]);

  createOrResetHeaders_(ss, SHEETS.DEVICES, [
    'Device ID', 'Device Name', 'Authorized', 'Paired On', 'Last Used'
  ]);

  createOrResetHeaders_(ss, SHEETS.SETTINGS, [
    'Setting', 'Value'
  ]);

  const settings = ss.getSheetByName(SHEETS.SETTINGS);
  const existing = getSettings_();

  if (!existing.TIMEZONE) {
    settings.appendRow(['TIMEZONE', Session.getScriptTimeZone()]);
  }
  if (!existing.ADMIN_PAIRING_CODE_HASH) {
    // Change this immediately by running setPairingCode('your-new-code')
    const defaultCode = 'CHANGE-ME-2026';
    settings.appendRow(['ADMIN_PAIRING_CODE_HASH', hashText_(defaultCode)]);
  }
  if (!existing.MAX_SHIFT_HOURS) {
    settings.appendRow(['MAX_SHIFT_HOURS', '12']);
  }

  formatSheets_();
  return 'Setup complete. Next, run setPairingCode(), then add students.';
}

/**
 * Run from Apps Script, for example:
 * setPairingCode('Fiesta-Tablet-8246')
 */
function setPairingCode(newCode) {
  if (!newCode || String(newCode).length < 8) {
    throw new Error('The pairing code must contain at least 8 characters.');
  }
  setSetting_('ADMIN_PAIRING_CODE_HASH', hashText_(String(newCode)));
  return 'Pairing code updated.';
}

/**
 * Run once per student, for example:
 * addStudent('S001', 'Natalie Cintrón', 'student@upr.edu', '4821')
 */
function addStudent(studentId, fullName, email, pin) {
  if (!studentId || !fullName || !pin) {
    throw new Error('Student ID, full name, and PIN are required.');
  }
  if (!/^\d{4,8}$/.test(String(pin))) {
    throw new Error('PIN must contain 4 to 8 numbers.');
  }

  const sheet = getSheet_(SHEETS.STUDENTS);
  const rows = sheet.getDataRange().getValues();
  const duplicate = rows.slice(1).some(row =>
    String(row[0]).trim().toLowerCase() === String(studentId).trim().toLowerCase()
  );
  if (duplicate) {
    throw new Error('That Student ID already exists.');
  }

  const salt = Utilities.getUuid();
  sheet.appendRow([
    String(studentId).trim(),
    String(fullName).trim(),
    String(email || '').trim(),
    salt,
    hashText_(salt + String(pin)),
    true
  ]);

  return `${fullName} was added successfully.`;
}

/**
 * Called from the tablet during one-time pairing.
 */
function pairDevice(pairingCode, deviceId, deviceName) {
  if (!pairingCode || !deviceId) {
    throw new Error('Pairing code and device ID are required.');
  }

  const expectedHash = getSettings_().ADMIN_PAIRING_CODE_HASH;
  if (!expectedHash || hashText_(String(pairingCode)) !== expectedHash) {
    throw new Error('Incorrect pairing code.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getSheet_(SHEETS.DEVICES);
    const values = sheet.getDataRange().getValues();
    const now = new Date();

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(deviceId)) {
        sheet.getRange(i + 1, 2, 1, 4).setValues([[
          deviceName || 'FIESTA IX Office Tablet', true,
          values[i][3] || now, now
        ]]);
        return { ok: true, message: 'This tablet is paired and ready.' };
      }
    }

    sheet.appendRow([
      deviceId,
      deviceName || 'FIESTA IX Office Tablet',
      true,
      now,
      now
    ]);

    return { ok: true, message: 'Tablet paired successfully.' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Returns active students for the kiosk dropdown.
 * It only returns names and IDs—never PIN information.
 */
function getActiveStudents(deviceId) {
  assertAuthorizedDevice_(deviceId);

  const sheet = getSheet_(SHEETS.STUDENTS);
  const values = sheet.getDataRange().getValues();

  return values.slice(1)
    .filter(row => row[5] === true || String(row[5]).toLowerCase() === 'true')
    .map(row => ({
      id: String(row[0]),
      name: String(row[1])
    }))
    .sort((a, b) => a.name.localeCompare(b.name));
}

/**
 * Handles both clock-in and clock-out.
 */
function submitPunch(deviceId, studentId, pin, action) {
  assertAuthorizedDevice_(deviceId);

  const normalizedAction = String(action || '').toUpperCase();
  if (!['IN', 'OUT'].includes(normalizedAction)) {
    throw new Error('Invalid clock action.');
  }

  const student = verifyStudentPin_(studentId, pin);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getSheet_(SHEETS.TIME_ENTRIES);
    const values = sheet.getDataRange().getValues();
    const openEntry = findOpenEntry_(values, student.id);
    const now = new Date();

    if (normalizedAction === 'IN') {
      if (openEntry) {
        throw new Error(`${student.name} is already clocked in.`);
      }
const regularSchedule =
  getTodayRegularSchedule_(student.id);

let lateClockIn = false;

if (regularSchedule && regularSchedule.start) {
  const latestAllowedIn =
    new Date(
      regularSchedule.start.getTime() +
      5 * 60000
    );

  if (now > latestAllowedIn) {
    lateClockIn = true;
  }
}
     const entryId = Utilities.getUuid();

const activityContext =
  getActivityPunchContext_(
    student.id,
    now
  );

let punchNotes = {};

if (activityContext) {
  try {
    punchNotes = JSON.parse(activityContext);
  } catch (e) {
    punchNotes = {};
  }
}

if (lateClockIn) {
  punchNotes.lateClockIn = true;
}

if (lateClockIn && regularSchedule) {
  createAttendanceAlert_(
    student.id,
    student.name,
    student.email || '',
    now,
    regularSchedule.start,
    'LATE'
  );
}

const notesText =
  Object.keys(punchNotes).length
    ? JSON.stringify(punchNotes)
    : '';

sheet.appendRow([
  entryId,
  student.id,
  student.name,
  now,
  '',
  '',
  deviceId,
  'OPEN',
  notesText
]);

      touchDevice_(deviceId, now);
    return {
  ok: true,
  action: 'IN',
  studentName: student.name,
  timestamp: formatDateTime_(now),
  lateClockIn: lateClockIn,
  message: lateClockIn
    ? `Clock-in recorded for ${student.name}. You are more than 5 minutes late, so this arrival was marked as late.`
    : `Welcome, ${student.name}. You are clocked in.`
};
    }

    if (!openEntry) {
      throw new Error(`${student.name} does not have an open clock-in.`);
    }

    const clockIn = new Date(openEntry.clockIn);
    const regularSchedule =
  getTodayRegularSchedule_(student.id);

let earlyClockOut = false;

if (regularSchedule && regularSchedule.end) {
  const earliestAllowedOut =
    new Date(
      regularSchedule.end.getTime() -
      5 * 60000
    );

  if (now < earliestAllowedOut) {
    earlyClockOut = true;
  }
}
if (earlyClockOut && regularSchedule) {
  createAttendanceAlert_(
    student.id,
    student.name,
    student.email || '',
    now,
    regularSchedule.end,
    'EARLY CLOCK OUT'
  );
}
    const totalHours = Math.round(((now - clockIn) / 3600000) * 100) / 100;
    const maxHours = Number(getSettings_().MAX_SHIFT_HOURS || 12);
    const status =
  totalHours > maxHours || earlyClockOut
    ? 'REVIEW'
    : 'COMPLETE';

    sheet.getRange(openEntry.rowNumber, 5, 1, 4).setValues([[
  now,
  totalHours,
  deviceId,
  status
]]);

const existingNotes =
  String(
    sheet.getRange(openEntry.rowNumber, 9).getValue() || ''
  ).trim();

if (!existingNotes) {
  const activityContext =
    getActivityPunchContext_(
      student.id,
      now
    );

  if (activityContext) {
    sheet
      .getRange(openEntry.rowNumber, 9)
      .setValue(activityContext);
  }
}

    touchDevice_(deviceId, now);
    return {
  ok: true,
  action: 'OUT',
  studentName: student.name,
  timestamp: formatDateTime_(now),
  totalHours: totalHours,
  earlyClockOut: earlyClockOut,
  message: earlyClockOut
    ? `Clock-out recorded for ${student.name}. This shift was marked for review because the clock-out was more than 5 minutes before the scheduled end time.`
    : `Goodbye, ${student.name}. Shift total: ${totalHours.toFixed(2)} hours.`
};
  } finally {
    lock.releaseLock();
  }
}

function getActivityPunchContext_(studentId, punchTime) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Scheduled Shifts');

  if (!sheet || sheet.getLastRow() < 2) {
    return '';
  }

  const tz = Session.getScriptTimeZone();
  const punchDateKey = Utilities.formatDate(
    punchTime,
    tz,
    'yyyy-MM-dd'
  );

  const rows = sheet
    .getDataRange()
    .getValues()
    .slice(1);

  for (let i = 0; i < rows.length; i++) {
    const rowStudentId =
      String(rows[i][3] || '').trim();

    const activityId =
      String(rows[i][1] || '').trim();

    const activityName =
      String(rows[i][2] || '').trim();

    const scheduledDate =
      rows[i][5];

    const scheduledStart =
      rows[i][6];

    const scheduledEnd =
      rows[i][7];

    const hoursTreatment =
      String(rows[i][10] || '')
        .trim()
        .toUpperCase();

    if (
      rowStudentId !== String(studentId).trim() ||
      !(scheduledDate instanceof Date) ||
      !(scheduledStart instanceof Date) ||
      !(scheduledEnd instanceof Date)
    ) {
      continue;
    }

    const scheduledDateKey =
      Utilities.formatDate(
        scheduledDate,
        tz,
        'yyyy-MM-dd'
      );

    if (scheduledDateKey !== punchDateKey) {
      continue;
    }

    if (
      punchTime.getTime() >= scheduledStart.getTime() &&
      punchTime.getTime() <= scheduledEnd.getTime()
    ) {
      return JSON.stringify({
        type: 'ACTIVITY',
        activityId: activityId,
        activityName: activityName,
        hoursTreatment: hoursTreatment
      });
    }
  }

  return '';
}

/**
 * Returns totals for a selected month.
 * monthText must be YYYY-MM, for example 2026-08.
 */
function getMonthlyTotals(monthText) {
  if (!/^\d{4}-\d{2}$/.test(String(monthText))) {
    throw new Error('Month must use YYYY-MM format.');
  }

  const [year, month] = monthText.split('-').map(Number);
  const sheet = getSheet_(SHEETS.TIME_ENTRIES);
  const values = sheet.getDataRange().getValues();
  const totals = {};

  values.slice(1).forEach(row => {
    const clockIn = row[3];
    const totalHours = Number(row[5]);

    if (!(clockIn instanceof Date) || !Number.isFinite(totalHours)) return;
    if (clockIn.getFullYear() !== year || clockIn.getMonth() !== month - 1) return;

    const id = String(row[1]);
    if (!totals[id]) {
      totals[id] = {
        studentId: id,
        studentName: String(row[2]),
        totalHours: 0,
        shifts: 0
      };
    }

    totals[id].totalHours += totalHours;
    totals[id].shifts += 1;
  });

  return Object.values(totals)
    .map(item => ({
      ...item,
      totalHours: Math.round(item.totalHours * 100) / 100
    }))
    .sort((a, b) => a.studentName.localeCompare(b.studentName));
}

/* -------------------- Private helper functions -------------------- */

function getSheet_(name) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(name);

  if (!sheet) {
    throw new Error(
      `Missing sheet: ${name}. Run setupPortal() first.`
    );
  }

  return sheet;
}

function getTodayRegularSchedule_(studentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Regular Schedules');

  if (!sheet || sheet.getLastRow() < 2) {
    return null;
  }

  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const weekday = Utilities.formatDate(now, tz, 'EEEE');

  const rows = sheet
    .getDataRange()
    .getValues()
    .slice(1);

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];

    const active =
      row[8] === true ||
      String(row[8]).toLowerCase() === 'true';

    if (!active) continue;

    if (String(row[1] || '') !== String(studentId)) {
      continue;
    }

    if (String(row[4] || '') !== weekday) {
      continue;
    }

    const start = row[5];
    const end = row[6];

    if (
      !(start instanceof Date) ||
      !(end instanceof Date)
    ) {
      continue;
    }

    return {
      start: new Date(
        now.getFullYear(),
        now.getMonth(),
        now.getDate(),
        start.getHours(),
        start.getMinutes(),
        0
      ),

      end: new Date(
        now.getFullYear(),
        now.getMonth(),
        now.getDate(),
        end.getHours(),
        end.getMinutes(),
        0
      )
    };
  }

  return null;
}

function getExpectedScheduleForDate_(
  studentId,
  dateValue
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const tz =
    Session.getScriptTimeZone();

  const targetDate =
    new Date(dateValue + 'T12:00:00');

  const weekday =
    Utilities.formatDate(
      targetDate,
      tz,
      'EEEE'
    );

  let expectedSchedule = null;


  // ==================================
  // 1. REGULAR SCHEDULE
  // ==================================

  const regularSheet =
    ss.getSheetByName(
      'Regular Schedules'
    );

  if (
    regularSheet &&
    regularSheet.getLastRow() >= 2
  ) {

    const rows =
      regularSheet
        .getDataRange()
        .getValues()
        .slice(1);

    for (
      let i = 0;
      i < rows.length;
      i++
    ) {

      const row = rows[i];

      const active =
        row[8] === true ||
        String(row[8] || '')
          .toLowerCase() === 'true';

      if (!active) {
        continue;
      }

      if (
        String(row[1] || '') !==
        String(studentId)
      ) {
        continue;
      }

      if (
        String(row[4] || '') !==
        weekday
      ) {
        continue;
      }
const regularStartDate =
  row[9];

if (
  regularStartDate instanceof Date
) {

  const regularStartDateKey =
    Utilities.formatDate(
      regularStartDate,
      tz,
      'yyyy-MM-dd'
    );

  if (
    dateValue <
    regularStartDateKey
  ) {
    continue;
  }
}
const regularEndDate =
  row[10];

if (
  regularEndDate instanceof Date
) {

  const regularEndDateKey =
    Utilities.formatDate(
      regularEndDate,
      tz,
      'yyyy-MM-dd'
    );

  if (
    dateValue >
    regularEndDateKey
  ) {
    continue;
  }
}
      const start =
        row[5];

      const end =
        row[6];

      if (
        !(start instanceof Date) ||
        !(end instanceof Date)
      ) {
        continue;
      }

      expectedSchedule = {
        start:
          new Date(
            targetDate.getFullYear(),
            targetDate.getMonth(),
            targetDate.getDate(),
            start.getHours(),
            start.getMinutes(),
            0
          ),

        end:
          new Date(
            targetDate.getFullYear(),
            targetDate.getMonth(),
            targetDate.getDate(),
            end.getHours(),
            end.getMinutes(),
            0
          ),

        source:
          'REGULAR'
      };

      break;
    }
  }

  // ==================================
  // 2. SCHEDULE EXCEPTIONS
  // ==================================

  const exceptionsSheet =
    ss.getSheetByName(
      'Schedule Exceptions'
    );

  if (
    exceptionsSheet &&
    exceptionsSheet.getLastRow() >= 2
  ) {

    const rows =
      exceptionsSheet
        .getDataRange()
        .getValues()
        .slice(1);

    rows.forEach(function(row) {

      const active =
        row[15] === true ||
        String(row[15] || '')
          .toLowerCase() === 'true';

      if (!active) {
        return;
      }

      const rowStudentId =
        String(row[1] || '').trim();

      if (
        rowStudentId !==
        String(studentId).trim()
      ) {
        return;
      }

      const exceptionDate =
        row[3];

      const regularStart =
        row[7];

      const regularEnd =
        row[8];

      const replacementStart =
        row[9];

      const replacementEnd =
        row[10];

      const replacedHours =
        Number(row[11] || 0);


      // --------------------------
      // ORIGINAL REGULAR DATE
      // --------------------------

      if (
        exceptionDate instanceof Date
      ) {

        const originalDateKey =
          Utilities.formatDate(
            exceptionDate,
            tz,
            'yyyy-MM-dd'
          );

        if (
          originalDateKey ===
          dateValue &&
          regularStart instanceof Date &&
          regularEnd instanceof Date
        ) {

          const regularHours =
            (
              regularEnd.getTime() -
              regularStart.getTime()
            ) / 3600000;

          if (
            replacedHours >=
            regularHours - 0.001
          ) {

            expectedSchedule = null;
          }
        }
      }


      // --------------------------
      // REPLACEMENT DATE
      // --------------------------

      if (
        replacementStart instanceof Date &&
        replacementEnd instanceof Date
      ) {

        const replacementDateKey =
          Utilities.formatDate(
            replacementStart,
            tz,
            'yyyy-MM-dd'
          );

        if (
          replacementDateKey ===
          dateValue
        ) {

          expectedSchedule = {
            start:
              replacementStart,

            end:
              replacementEnd,

            source:
              'REPLACEMENT'
          };
        }
      }
    });
  }

  return expectedSchedule;
}


function getExpectedSchedulesForDate_(
  studentId,
  dateValue
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const tz =
    Session.getScriptTimeZone();

  const targetDate =
    new Date(dateValue + 'T12:00:00');

  const weekday =
    Utilities.formatDate(
      targetDate,
      tz,
      'EEEE'
    );

  const expectedSchedules = [];


  // ==================================
  // 1. REGULAR SCHEDULES
  // ==================================

  const regularSheet =
    ss.getSheetByName(
      'Regular Schedules'
    );

  if (
    regularSheet &&
    regularSheet.getLastRow() >= 2
  ) {

    const rows =
      regularSheet
        .getDataRange()
        .getValues()
        .slice(1);

    rows.forEach(function(row) {

      const active =
        row[8] === true ||
        String(row[8] || '')
          .toLowerCase() === 'true';

      if (!active) {
        return;
      }

      if (
        String(row[1] || '').trim() !==
        String(studentId).trim()
      ) {
        return;
      }

      if (
        String(row[4] || '').trim() !==
        weekday
      ) {
        return;
      }


      // --------------------------
      // START DATE
      // --------------------------

      const regularStartDate =
        row[9];

      if (
        regularStartDate instanceof Date
      ) {

        const regularStartDateKey =
          Utilities.formatDate(
            regularStartDate,
            tz,
            'yyyy-MM-dd'
          );

        if (
          dateValue <
          regularStartDateKey
        ) {
          return;
        }
      }


      // --------------------------
      // END DATE
      // --------------------------

      const regularEndDate =
        row[10];

      if (
        regularEndDate instanceof Date
      ) {

        const regularEndDateKey =
          Utilities.formatDate(
            regularEndDate,
            tz,
            'yyyy-MM-dd'
          );

        if (
          dateValue >
          regularEndDateKey
        ) {
          return;
        }
      }


      const start =
        row[5];

      const end =
        row[6];

      if (
        !(start instanceof Date) ||
        !(end instanceof Date)
      ) {
        return;
      }


      expectedSchedules.push({

        scheduleId:
          String(row[0] || ''),

        start:
          new Date(
            targetDate.getFullYear(),
            targetDate.getMonth(),
            targetDate.getDate(),
            start.getHours(),
            start.getMinutes(),
            0
          ),

        end:
          new Date(
            targetDate.getFullYear(),
            targetDate.getMonth(),
            targetDate.getDate(),
            end.getHours(),
            end.getMinutes(),
            0
          ),

        source:
          'REGULAR'

      });

    });
  }


  // ==================================
  // SORT BY START TIME
  // ==================================

  expectedSchedules.sort(
    function(a, b) {
      return (
        a.start.getTime() -
        b.start.getTime()
      );
    }
  );


  return expectedSchedules;
}

function testMultipleSchedules() {

  const result =
    getExpectedSchedulesForDate_(
      'STU-123456',
      '2026-09-08'
    );

  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );
}

function reconcileHistoricalMissingAttendance() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const tz =
    Session.getScriptTimeZone();

  const today =
    new Date();

  const todayKey =
    Utilities.formatDate(
      today,
      tz,
      'yyyy-MM-dd'
    );

  const scheduleSheet =
    ss.getSheetByName(
      'Regular Schedules'
    );

  if (
    !scheduleSheet ||
    scheduleSheet.getLastRow() < 2
  ) {
    return {
      ok: true,
      created: 0
    };
  }

  const timeSheet =
    ss.getSheetByName(
      'Time Entries'
    );

  const manualSheet =
    ss.getSheetByName(
      'Attendance Manual Shifts'
    );

  const scheduleRows =
    scheduleSheet
      .getDataRange()
      .getValues()
      .slice(1);

  const timeRows =
    timeSheet &&
    timeSheet.getLastRow() >= 2
      ? timeSheet
          .getDataRange()
          .getValues()
          .slice(1)
      : [];

  const manualRows =
    manualSheet &&
    manualSheet.getLastRow() >= 2
      ? manualSheet
          .getDataRange()
          .getValues()
          .slice(1)
      : [];

  const studentsToCheck = {};

  scheduleRows.forEach(function(row) {

    const active =
      row[8] === true ||
      String(row[8] || '')
        .toLowerCase() === 'true';

    if (!active) {
      return;
    }

    const studentId =
      String(row[1] || '').trim();

    const startDate =
      row[9];

    if (
      !studentId ||
      !(startDate instanceof Date)
    ) {
      return;
    }

    if (!studentsToCheck[studentId]) {
      studentsToCheck[studentId] = {
        studentId:
          studentId,

        studentName:
          String(row[2] || ''),

        studentEmail:
          String(row[3] || ''),

        earliestStartDate:
          startDate
      };

      return;
    }

    if (
      startDate.getTime() <
      studentsToCheck[
        studentId
      ].earliestStartDate.getTime()
    ) {
      studentsToCheck[
        studentId
      ].earliestStartDate =
        startDate;
    }
  });


  let createdCount = 0;


  Object.keys(
    studentsToCheck
  ).forEach(function(studentId) {

    const student =
      studentsToCheck[studentId];

    const cursor =
      new Date(
        student.earliestStartDate
      );

    cursor.setHours(
      12,
      0,
      0,
      0
    );


    while (true) {

      const dateKey =
        Utilities.formatDate(
          cursor,
          tz,
          'yyyy-MM-dd'
        );

      if (
        dateKey >= todayKey
      ) {
        break;
      }


      const expectedSchedule =
        getExpectedScheduleForDate_(
          studentId,
          dateKey
        );

      if (expectedSchedule) {

        let hasAttendance = false;


        // --------------------------
        // TIME ENTRIES
        // --------------------------

        for (
          let i = 0;
          i < timeRows.length;
          i++
        ) {

          const row =
            timeRows[i];

          const rowStudentId =
            String(
              row[1] || ''
            ).trim();

          const clockIn =
            row[3];

          if (
            rowStudentId !==
              studentId ||
            !(clockIn instanceof Date)
          ) {
            continue;
          }

          const entryDateKey =
            Utilities.formatDate(
              clockIn,
              tz,
              'yyyy-MM-dd'
            );

          if (
            entryDateKey === dateKey
          ) {
            hasAttendance = true;
            break;
          }
        }


        // --------------------------
        // MANUAL SHIFTS
        // --------------------------

        if (!hasAttendance) {

          for (
            let i = 0;
            i < manualRows.length;
            i++
          ) {

            const row =
              manualRows[i];

            const rowStudentId =
              String(
                row[1] || ''
              ).trim();

            const workDate =
              row[3];

            const active =
              row[10] === true ||
              String(
                row[10] || ''
              ).toLowerCase() ===
                'true';

            if (
              !active ||
              rowStudentId !==
                studentId ||
              !(workDate instanceof Date)
            ) {
              continue;
            }

            const manualDateKey =
              Utilities.formatDate(
                workDate,
                tz,
                'yyyy-MM-dd'
              );

            if (
              manualDateKey ===
              dateKey
            ) {
              hasAttendance = true;
              break;
            }
          }
        }


        // --------------------------
        // CREATE MISSING
        // --------------------------

        if (!hasAttendance) {

          const result =
            createAttendanceAlert_(
              studentId,
              student.studentName,
              student.studentEmail,
              new Date(
                dateKey +
                'T12:00:00'
              ),
              expectedSchedule.start,
              'MISSING'
            );

          if (
            result &&
            !result.duplicate
          ) {
            createdCount++;
          }
        }
      }


      cursor.setDate(
        cursor.getDate() + 1
      );
    }
  });


  return {
    ok: true,
    created:
      createdCount
  };
}

function createOrResetHeaders_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  sheet.setFrozenRows(1);
}

function formatSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.values(SHEETS).forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    sheet.getRange(1, 1, 1, sheet.getLastColumn())
      .setFontWeight('bold')
      .setBackground('#1f4e78')
      .setFontColor('#ffffff');
    sheet.autoResizeColumns(1, sheet.getLastColumn());
  });

  const entries = ss.getSheetByName(SHEETS.TIME_ENTRIES);
  entries.getRange('D:E').setNumberFormat('m/d/yyyy h:mm AM/PM');
  entries.getRange('F:F').setNumberFormat('0.00');
}

function getSettings_() {
  const sheet = getSheet_(SHEETS.SETTINGS);
  const values = sheet.getDataRange().getValues();
  const output = {};

  values.slice(1).forEach(row => {
    if (row[0]) output[String(row[0])] = String(row[1]);
  });
  return output;
}

function setSetting_(key, value) {
  const sheet = getSheet_(SHEETS.SETTINGS);
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(key)) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function hashText_(text) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(text),
    Utilities.Charset.UTF_8
  );
  return digest.map(byte => {
    const value = byte < 0 ? byte + 256 : byte;
    return value.toString(16).padStart(2, '0');
  }).join('');
}
function verifyStudentNumberPin(studentNumber, pin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const profileSheet = ss.getSheetByName('Student Profiles');

  if (!profileSheet) {
    throw new Error('Student Profiles sheet was not found.');
  }

  const values = profileSheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    const studentId = String(values[i][0] || '').trim();
    const number = String(values[i][2] || '').trim();
    const active =
      values[i][7] === true ||
      String(values[i][7]).toLowerCase() === 'true';

    if (number === String(studentNumber).trim() && active) {
      return verifyStudentPin_(studentId, pin);
    }
  }

  throw new Error('Student number not found.');
}
function verifyStudentPin(studentId, pin) {
  return verifyStudentPin_(studentId, pin);
}

function verifyStudentPin_(studentId, pin) {
  const sheet = getSheet_(SHEETS.STUDENTS);
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const active = row[5] === true || String(row[5]).toLowerCase() === 'true';

    if (String(row[0]) === String(studentId) && active) {
      const expectedHash = String(row[4]);
      const actualHash = hashText_(String(row[3]) + String(pin));

      if (actualHash !== expectedHash) {
        throw new Error('Incorrect PIN.');
      }

      return {
        id: String(row[0]),
        name: String(row[1]),
        email: String(row[2] || '')
      };
    }
  }

  throw new Error('Student not found or inactive.');
}

function assertAuthorizedDevice_(deviceId) {
  if (!deviceId) {
    throw new Error('This device is not paired.');
  }

  const sheet = getSheet_(SHEETS.DEVICES);
  const values = sheet.getDataRange().getValues();

  const authorized = values.slice(1).some(row =>
    String(row[0]) === String(deviceId) &&
    (row[2] === true || String(row[2]).toLowerCase() === 'true')
  );

  if (!authorized) {
    throw new Error('Clock-in is only available on the authorized office tablet.');
  }
}

function findOpenEntry_(values, studentId) {
  for (let i = values.length - 1; i >= 1; i--) {
    const row = values[i];
    const sameStudent = String(row[1]) === String(studentId);
    const isOpen = !row[4] && String(row[7]).toUpperCase() === 'OPEN';

    if (sameStudent && isOpen) {
      return {
        rowNumber: i + 1,
        clockIn: row[3]
      };
    }
  }
  return null;
}

function touchDevice_(deviceId, now) {
  const sheet = getSheet_(SHEETS.DEVICES);
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(deviceId)) {
      sheet.getRange(i + 1, 5).setValue(now);
      return;
    }
  }
}

function formatDateTime_(date) {
  const timezone = getSettings_().TIMEZONE || Session.getScriptTimeZone();
  return Utilities.formatDate(date, timezone, 'MMM d, yyyy h:mm a');
}function createMyPairingCode() {
  setPairingCode('FiestaIX20267!');
}function addHeidy() {
  addStudent(
    'A001',
    'Heidy Orama',
    'heidy.orama@upr.edu.com',
    '1234'
  );
}function addHeidy() {
  addStudent(
    'ADMIN001',
    'Heidy Orama',
    'prheidy20@gmail.com',
    '2580'
  );
}function setupSchedulingModule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  createOrResetHeaders_(ss, 'Activities', [
    'Activity ID',
    'Activity Name',
    'Date',
    'Start Time',
    'End Time',
    'Location',
    'Description',
    'Students Needed',
    'Status',
    'Created On'
  ]);

  createOrResetHeaders_(ss, 'Scheduled Shifts', [
    'Shift ID',
    'Activity ID',
    'Activity Name',
    'Student ID',
    'Student Name',
    'Scheduled Date',
    'Scheduled Start',
    'Scheduled End',
    'Location',
    'Attendance Status',
    'Hours Status',
    'Approved Hours',
    'Approved By',
    'Approval Date',
    'Notes'
  ]);

  const activities = ss.getSheetByName('Activities');
  const shifts = ss.getSheetByName('Scheduled Shifts');

  activities.setFrozenRows(1);
  shifts.setFrozenRows(1);

  activities.getRange('C:C').setNumberFormat('m/d/yyyy');
  activities.getRange('D:E').setNumberFormat('h:mm AM/PM');
  activities.getRange('J:J').setNumberFormat('m/d/yyyy h:mm AM/PM');

  shifts.getRange('F:F').setNumberFormat('m/d/yyyy');
  shifts.getRange('G:H').setNumberFormat('h:mm AM/PM');
  shifts.getRange('L:L').setNumberFormat('0.00');
  shifts.getRange('N:N').setNumberFormat('m/d/yyyy h:mm AM/PM');

  activities.autoResizeColumns(1, activities.getLastColumn());
  shifts.autoResizeColumns(1, shifts.getLastColumn());

  return 'Scheduling module created successfully.';
}function addTestActivity() {
  const activities = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Activities');

  const activityId = Utilities.getUuid();

  activities.appendRow([
    activityId,
    'FIESTA IX Office Work',
    new Date(2026, 7, 7),
    new Date(2026, 7, 7, 9, 0),
    new Date(2026, 7, 7, 13, 0),
    'FIESTA IX Office',
    'Testing the scheduling and time verification system.',
    1,
    'Scheduled',
    new Date()
  ]);

  return activityId;
}function createActivity(formData) {
  if (!formData) {
    throw new Error('No activity information was received.');
  }

  const activityName = String(formData.activityName || '').trim();
  const dateText = String(formData.date || '').trim();
  const startText = String(formData.startTime || '').trim();
  const endText = String(formData.endTime || '').trim();
  const location = String(formData.location || '').trim();
  const description = String(formData.description || '').trim();
  const studentsNeeded = Number(formData.studentsNeeded || 0);

  if (!activityName || !dateText || !startText || !endText || !location) {
    throw new Error('Complete all required fields.');
  }

  const activityDate = parsePortalDateTime_(dateText, '00:00');
  const startDateTime = parsePortalDateTime_(dateText, startText);
  const endDateTime = parsePortalDateTime_(dateText, endText);

  if (endDateTime <= startDateTime) {
    throw new Error('The ending time must be later than the starting time.');
  }

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Activities');

  if (!sheet) {
    throw new Error(
      'The Activities sheet was not found. Run setupSchedulingModule first.'
    );
  }

  const activityId = Utilities.getUuid();

  sheet.appendRow([
    activityId,
    activityName,
    activityDate,
    startDateTime,
    endDateTime,
    location,
    description,
    studentsNeeded,
    'Scheduled',
    new Date()
  ]);

  return {
    ok: true,
    activityId: activityId,
    activityName: activityName,
    message: activityName + ' was created successfully.'
  };
}

function parsePortalDateTime_(dateText, timeText) {
  const dateParts = dateText.split('-').map(Number);
  const timeParts = timeText.split(':').map(Number);

  if (
    dateParts.length !== 3 ||
    timeParts.length < 2 ||
    dateParts.some(Number.isNaN) ||
    timeParts.some(Number.isNaN)
  ) {
    throw new Error('The date or time format is invalid.');
  }

  return new Date(
    dateParts[0],
    dateParts[1] - 1,
    dateParts[2],
    timeParts[0],
    timeParts[1],
    0
  );
}function setupAttendanceAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  createOrResetHeaders_(ss, 'Regular Schedules', [
    'Schedule ID',
    'Student ID',
    'Student Name',
    'Student Email',
    'Day of Week',
    'Start Time',
    'End Time',
    'Location',
    'Active'
  ]);

  createOrResetHeaders_(ss, 'Attendance Alerts', [
    'Alert ID',
    'Student ID',
    'Student Name',
    'Student Email',
    'Scheduled Date',
    'Scheduled Start',
    'Alert Type',
    'Alert Time',
    'Coordinator Email',
    'Student Notified',
    'Coordinator Notified',
    'Status'
  ]);

  const schedules = ss.getSheetByName('Regular Schedules');
  const alerts = ss.getSheetByName('Attendance Alerts');

  schedules.setFrozenRows(1);
  alerts.setFrozenRows(1);

  schedules.getRange('F:G').setNumberFormat('h:mm AM/PM');

  alerts.getRange('E:E').setNumberFormat('m/d/yyyy');
  alerts.getRange('F:F').setNumberFormat('h:mm AM/PM');
  alerts.getRange('H:H').setNumberFormat('m/d/yyyy h:mm AM/PM');

  schedules.autoResizeColumns(1, schedules.getLastColumn());
  alerts.autoResizeColumns(1, alerts.getLastColumn());

  setSetting_('COORDINATOR_EMAIL', 'prheidy20@gmail.com');
  setSetting_('LATE_GRACE_MINUTES', '5');
  setSetting_('ALERT_STUDENTS', 'TRUE');
  setSetting_('ALERT_COORDINATOR', 'TRUE');

  return 'Attendance alert system created successfully.';
  }
  function createAttendanceAlert_(
  studentId,
  studentName,
  studentEmail,
  scheduledDate,
  scheduledStart,
  alertType
) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Attendance Alerts');

  if (!sheet) {
    throw new Error('Attendance Alerts sheet was not found.');
  }

  const normalizedType =
    String(alertType || '')
      .trim()
      .toUpperCase();

  const allowedTypes = [
    'LATE',
    'MISSING',
    'EARLY CLOCK OUT'
  ];

  if (!allowedTypes.includes(normalizedType)) {
    throw new Error('Invalid attendance alert type.');
  }

  const coordinatorEmail =
    String(
      getSettings_().COORDINATOR_EMAIL || ''
    ).trim();

  const now = new Date();

const tz = Session.getScriptTimeZone();

const alertDateKey =
  Utilities.formatDate(
    scheduledDate instanceof Date
      ? scheduledDate
      : now,
    tz,
    'yyyy-MM-dd'
  );

const existingRows =
  sheet.getDataRange().getValues();

for (let i = 1; i < existingRows.length; i++) {
  const row = existingRows[i];

  const existingStudentId =
    String(row[1] || '').trim();

  const existingDate = row[4];

  const existingType =
    String(row[6] || '')
      .trim()
      .toUpperCase();

  const existingStatus =
    String(row[11] || '')
      .trim()
      .toUpperCase();

  if (!(existingDate instanceof Date)) {
    continue;
  }

  const existingDateKey =
    Utilities.formatDate(
      existingDate,
      tz,
      'yyyy-MM-dd'
    );

  if (
    existingStudentId === String(studentId).trim() &&
    existingDateKey === alertDateKey &&
    existingType === normalizedType &&
    existingStatus !== 'CANCELLED'
  ) {
    return {
      ok: true,
      duplicate: true,
      alertType: normalizedType
    };
  }
}

  sheet.appendRow([
    Utilities.getUuid(),
    String(studentId || '').trim(),
    String(studentName || '').trim(),
    String(studentEmail || '').trim(),
    scheduledDate || '',
    scheduledStart || '',
    normalizedType,
    now,
    coordinatorEmail,
    false,
    false,
    'OPEN'
  ]);

  SpreadsheetApp.flush();

  return {
    ok: true,
    alertType: normalizedType
  };
}

  function getStudentsForSchedule() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Students');

  if (!sheet) {
    throw new Error('The Students sheet was not found.');
  }

  const rows = sheet.getDataRange().getValues();

  return rows
    .slice(1)
    .filter(row =>
      row[5] === true ||
      String(row[5]).toLowerCase() === 'true'
    )
    .map(row => ({
      studentId: String(row[0]),
      studentName: String(row[1]),
      studentEmail: String(row[2] || '')
    }))
    .sort((a, b) =>
      a.studentName.localeCompare(b.studentName)
    );
}

function getAttendanceEditorStudents() {
  return getStudentsForSchedule();
}

function saveRegularSchedule(formData) {
  if (!formData) {
    throw new Error('No schedule information was received.');
  }

  const studentId = String(formData.studentId || '').trim();
  const dayOfWeek = String(formData.dayOfWeek || '').trim();
  const startTime = String(formData.startTime || '').trim();
  const endTime = String(formData.endTime || '').trim();
  const location = String(formData.location || '').trim();
  const startDate = String(formData.startDate || '').trim();
  const endDateText = String(formData.endDate || '').trim();

  if (
    !studentId ||
  !dayOfWeek ||
  !startDate ||
  !startTime ||
  !endTime ||
  !endDateText ||
  !location
  
  ) {
    throw new Error('Complete all required schedule fields.');
  }

  const student = getStudentForSchedule_(studentId);
  const scheduleStartDate =
  new Date(startDate + 'T12:00:00');
  const startTimeDate = timeTextToDate_(startTime);
  const scheduleEndDate =
  new Date(endDateText + 'T12:00:00');

if (
  scheduleEndDate <
  scheduleStartDate
) {
  throw new Error(
    'The end date cannot be earlier than the start date.'
  );
}
  const endDate = timeTextToDate_(endTime);

  if (endDate <= startTimeDate) {
    throw new Error(
      'The ending time must be later than the starting time.'
    );
  }

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Regular Schedules');

  if (!sheet) {
    throw new Error(
      'The Regular Schedules sheet was not found.'
    );
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const rows = sheet.getDataRange().getValues();

    // Update an existing schedule for the same student and day.
    for (let i = 1; i < rows.length; i++) {
      const sameStudent =
        String(rows[i][1]) === student.studentId;

      const sameDay =
        String(rows[i][4]).toLowerCase() ===
        dayOfWeek.toLowerCase();

      const active =
        rows[i][8] === true ||
        String(rows[i][8]).toLowerCase() === 'true';

      if (sameStudent && sameDay && active) {
        sheet.getRange(i + 1, 3, 1, 9).setValues([[
  student.studentName,
  student.studentEmail,
  dayOfWeek,
  startTimeDate,
  endDate,
  location,
  true,
  scheduleStartDate,
  scheduleEndDate
]]);

        return {
          ok: true,
          message:
            student.studentName +
            "'s " +
            dayOfWeek +
            ' schedule was updated.'
        };
      }
    }

  sheet.appendRow([
  Utilities.getUuid(),
  student.studentId,
  student.studentName,
  student.studentEmail,
  dayOfWeek,
  startTimeDate,
  endDate,
  location,
  true,
  scheduleStartDate,
  scheduleEndDate
]);
    return {
      ok: true,
      message:
        student.studentName +
        "'s " +
        dayOfWeek +
        ' schedule was saved.'
    };

  } finally {
    lock.releaseLock();
  }
}

function saveWeeklyRegularSchedule(formData) {

  if (!formData) {
    throw new Error(
      'No weekly schedule information was received.'
    );
  }

  const studentId =
    String(formData.studentId || '').trim();

  const startDateText =
    String(formData.startDate || '').trim();

  const endDateText =
    String(formData.endDate || '').trim();

  const location =
    String(formData.location || '').trim();

  const shifts =
    Array.isArray(formData.shifts)
      ? formData.shifts
      : [];

  if (
    !studentId ||
    !startDateText ||
    !endDateText ||
    !location ||
    !shifts.length
  ) {
    throw new Error(
      'Complete the student, schedule dates, location, and at least one work shift.'
    );
  }

  const student =
    getStudentForSchedule_(studentId);

  const scheduleStartDate =
    new Date(
      startDateText + 'T12:00:00'
    );

  const scheduleEndDate =
    new Date(
      endDateText + 'T12:00:00'
    );

  if (
    scheduleEndDate <
    scheduleStartDate
  ) {
    throw new Error(
      'The end date cannot be earlier than the start date.'
    );
  }


  // -------------------------
  // VALIDATE ALL SHIFTS
  // -------------------------

  const preparedShifts = [];

  shifts.forEach(function(shift) {

    const dayOfWeek =
      String(
        shift.dayOfWeek || ''
      ).trim();

    const startTime =
      String(
        shift.startTime || ''
      ).trim();

    const endTime =
      String(
        shift.endTime || ''
      ).trim();

    if (
      !dayOfWeek ||
      !startTime ||
      !endTime
    ) {
      throw new Error(
        'Every work shift must have a starting and ending time.'
      );
    }

    const startTimeDate =
      timeTextToDate_(startTime);

    const endTimeDate =
      timeTextToDate_(endTime);

    if (
      endTimeDate <=
      startTimeDate
    ) {
      throw new Error(
        dayOfWeek +
        ': the ending time must be later than the starting time.'
      );
    }

    preparedShifts.push({
      dayOfWeek:
        dayOfWeek,

      startTime:
        startTimeDate,

      endTime:
        endTimeDate
    });

  });


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        'Regular Schedules'
      );

  if (!sheet) {
    throw new Error(
      'The Regular Schedules sheet was not found.'
    );
  }


  const lock =
    LockService.getScriptLock();

  lock.waitLock(10000);

  try {

    preparedShifts.forEach(
      function(shift) {

        sheet.appendRow([
          Utilities.getUuid(),
          student.studentId,
          student.studentName,
          student.studentEmail,
          shift.dayOfWeek,
          shift.startTime,
          shift.endTime,
          location,
          true,
          scheduleStartDate,
          scheduleEndDate
        ]);

      }
    );

    return {
      ok: true,

      message:
        student.studentName +
        "'s weekly schedule was saved with " +
        preparedShifts.length +
        ' work shift' +
        (
          preparedShifts.length === 1
            ? '.'
            : 's.'
        )
    };

  } finally {

    lock.releaseLock();

  }
}

function getStudentForSchedule_(studentId) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Students');

  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const active =
      rows[i][5] === true ||
      String(rows[i][5]).toLowerCase() === 'true';

    if (
      String(rows[i][0]) === String(studentId) &&
      active
    ) {
      return {
        studentId: String(rows[i][0]),
        studentName: String(rows[i][1]),
        studentEmail: String(rows[i][2] || '')
      };
    }
  }

  throw new Error('The selected student was not found.');
}


function timeTextToDate_(timeText) {
  const parts = String(timeText).split(':').map(Number);

  if (
    parts.length !== 2 ||
    parts.some(Number.isNaN)
  ) {
    throw new Error('The time format is invalid.');
  }

  return new Date(
    1899,
    11,
    30,
    parts[0],
    parts[1],
    0
  );
}function setupScheduleExceptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  createOrResetHeaders_(ss, 'Schedule Exceptions', [
    'Exception ID',
    'Student ID',
    'Student Name',
    'Exception Date',
    'Exception Type',
    'Activity ID',
    'Activity Name',
    'Regular Start',
    'Regular End',
    'Replacement Start',
    'Replacement End',
    'Paid Hours',
    'Approved By',
    'Created On',
    'Notes',
    'Active'
  ]);

  const sheet = ss.getSheetByName('Schedule Exceptions');

  sheet.setFrozenRows(1);
  sheet.getRange('D:D').setNumberFormat('m/d/yyyy');
  sheet.getRange('H:K').setNumberFormat('h:mm AM/PM');
  sheet.getRange('L:L').setNumberFormat('0.00');
  sheet.getRange('N:N').setNumberFormat('m/d/yyyy h:mm AM/PM');
  sheet.autoResizeColumns(1, sheet.getLastColumn());

  return 'Schedule Exceptions module created successfully.';
}
/**
 * FIESTA Hub 2.1 — Student Profiles
 * Paste this entire file at the bottom of Code.gs.
 *
 * This module keeps sensitive profile information in a separate sheet
 * from the Students authentication sheet.
 */

function setupStudentProfilesModule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  createOrResetHeaders_(ss, 'Student Profiles', [
    'Student ID',
    'Full Name',
    'Student Number',
    'Last 4 SSN',
    'Email',
    'Phone',
    'Home Address',
    'Active',
    'Created On',
    'Updated On',
    'Internal Notes'
  ]);

  createOrResetHeaders_(ss, 'Student Achievements', [
    'Achievement ID',
    'Student ID',
    'Student Name',
    'Achievement Name',
    'Description',
    'Date Awarded',
    'Awarded By',
    'Active'
  ]);

  createOrResetHeaders_(ss, 'Student Certificates', [
    'Certificate ID',
    'Student ID',
    'Student Name',
    'Certificate Name',
    'Date Issued',
    'Expiration Date',
    'File Link',
    'Notes',
    'Active'
  ]);

  const profiles = ss.getSheetByName('Student Profiles');
  const achievements = ss.getSheetByName('Student Achievements');
  const certificates = ss.getSheetByName('Student Certificates');

  profiles.setFrozenRows(1);
  achievements.setFrozenRows(1);
  certificates.setFrozenRows(1);

  profiles.getRange('I:J').setNumberFormat('m/d/yyyy h:mm AM/PM');
  achievements.getRange('F:F').setNumberFormat('m/d/yyyy');
  certificates.getRange('E:F').setNumberFormat('m/d/yyyy');

  profiles.autoResizeColumns(1, profiles.getLastColumn());
  achievements.autoResizeColumns(1, achievements.getLastColumn());
  certificates.autoResizeColumns(1, certificates.getLastColumn());

  return 'Student Profiles module created successfully.';
}


function saveStudentProfile(formData) {
  if (!formData) {
    throw new Error('No student information was received.');
  }

  const studentId = String(formData.studentId || '').trim();
  const fullName = String(formData.fullName || '').trim();
  const studentNumber = String(formData.studentNumber || '').trim();
  const last4 = String(formData.last4 || '').trim();
  const email = String(formData.email || '').trim();
  const phone = String(formData.phone || '').trim();
  const address = String(formData.address || '').trim();
  const notes = String(formData.notes || '').trim();

  if (!studentId || !fullName || !studentNumber || !last4 || !email) {
    throw new Error(
      'Complete Student ID, full name, student number, last 4 SSN, and email.'
    );
  }

  if (!/^\d{4}$/.test(last4)) {
    throw new Error('Last 4 SSN must contain exactly four numbers.');
  }

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Student Profiles');

  if (!sheet) {
    throw new Error(
      'Student Profiles sheet was not found. Run setupStudentProfilesModule first.'
    );
  }

  const rows = sheet.getDataRange().getValues();
  const now = new Date();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === studentId) {
      sheet.getRange(i + 1, 2, 1, 10).setValues([[
        fullName,
        studentNumber,
        last4,
        email,
        phone,
        address,
        true,
        rows[i][8] || now,
        now,
        notes
      ]]);

      syncStudentAuthRecord_(studentId, fullName, email);

      return {
        ok: true,
        message: fullName + "'s profile was updated."
      };
    }
  }

  sheet.appendRow([
    studentId,
    fullName,
    studentNumber,
    last4,
    email,
    phone,
    address,
    true,
    now,
    now,
    notes
  ]);

  syncStudentAuthRecord_(studentId, fullName, email);

  return {
    ok: true,
    message: fullName + "'s profile was created."
  };
}


function syncStudentAuthRecord_(studentId, fullName, email) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Students');

  if (!sheet) return;

  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(studentId)) {
      sheet.getRange(i + 1, 2).setValue(fullName);
      sheet.getRange(i + 1, 3).setValue(email);
      return;
    }
  }
}


function getStudentProfiles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const profilesSheet = ss.getSheetByName('Student Profiles');
  const requestsSheet = ss.getSheetByName('Activity Participation Requests');

  if (!profilesSheet) {
    throw new Error(
      'Student Profiles sheet was not found. Run setupStudentProfilesModule first.'
    );
  }

  const rows = profilesSheet.getDataRange().getValues();
  const hoursByStudent = getTotalHoursByStudent_();

  return rows
    .slice(1)
    .filter(row =>
      row[7] === true ||
      String(row[7]).toLowerCase() === 'true'
    )
    .map(row => {
      const id = String(row[0]);

      return {
        studentId: id,
        fullName: String(row[1] || ''),
        studentNumber: String(row[2] || ''),
        last4Masked: row[3] ? '•••• ' + String(row[3]) : '',
        email: String(row[4] || ''),
        phone: String(row[5] || ''),
        address: String(row[6] || ''),
        totalHours: Number(hoursByStudent[id] || 0)
      };
    })
    .sort((a, b) => a.fullName.localeCompare(b.fullName));
}


function getStudentProfileDetail(studentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const profilesSheet = ss.getSheetByName('Student Profiles');

  if (!profilesSheet) {
    throw new Error('Student Profiles sheet was not found.');
  }

  const rows = profilesSheet.getDataRange().getValues();
  let profile = null;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(studentId)) {
      profile = {
        studentId: String(rows[i][0]),
        fullName: String(rows[i][1] || ''),
        studentNumber: String(rows[i][2] || ''),
        last4: String(rows[i][3] || ''),
        last4Masked: rows[i][3] ? '•••• ' + String(rows[i][3]) : '',
        email: String(rows[i][4] || ''),
        phone: String(rows[i][5] || ''),
        address: String(rows[i][6] || ''),
        notes: String(rows[i][10] || '')
      };
      break;
    }
  }

  if (!profile) {
    throw new Error('Student profile not found.');
  }

  profile.totalHours = Number(getTotalHoursByStudent_()[studentId] || 0);
  profile.achievements = getStudentAchievements_(studentId);
  profile.certificates = getStudentCertificates_(studentId);

  return profile;
}


function getTotalHoursByStudent_() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Time Entries');

  const totals = {};
  if (!sheet || sheet.getLastRow() < 2) return totals;

  const rows = sheet.getDataRange().getValues();

  rows.slice(1).forEach(row => {
    const studentId = String(row[1] || '');
    const hours = Number(row[5]);

    if (studentId && Number.isFinite(hours)) {
      totals[studentId] = (totals[studentId] || 0) + hours;
    }
  });

  Object.keys(totals).forEach(id => {
    totals[id] = Math.round(totals[id] * 100) / 100;
  });

  return totals;
}


function getStudentAchievements_(studentId) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Student Achievements');

  if (!sheet || sheet.getLastRow() < 2) return [];

  return sheet.getDataRange().getValues()
    .slice(1)
    .filter(row =>
      String(row[1]) === String(studentId) &&
      (row[7] === true || String(row[7]).toLowerCase() === 'true')
    )
    .map(row => ({
      name: String(row[3] || ''),
      description: String(row[4] || ''),
      date: row[5] instanceof Date
        ? Utilities.formatDate(
            row[5],
            Session.getScriptTimeZone(),
            'MMM d, yyyy'
          )
        : ''
    }));
}

function getStudentAchievements(studentId) {

  if (!studentId) {
    throw new Error(
      'Student ID is required.'
    );
  }

  const achievements =
    getStudentAchievements_(
      studentId
    );

  return {
    achievements: achievements || []
  };
}

function getFiestaAchievementPeriod_(date) {

  const targetDate =
    date instanceof Date
      ? date
      : new Date();

  const month =
    targetDate.getMonth() + 1;

  const year =
    targetDate.getFullYear();

  if (month >= 10 && month <= 12) {
    return {
      key: year + '-OCT-DEC',
      label: 'October–December ' + year,
      startMonth: 10,
      endMonth: 12,
      year: year
    };
  }

  if (month >= 1 && month <= 5) {
    return {
      key: year + '-JAN-MAY',
      label: 'January–May ' + year,
      startMonth: 1,
      endMonth: 5,
      year: year
    };
  }

  if (month >= 6 && month <= 7) {
    return {
      key: year + '-JUN-JUL',
      label: 'June–July ' + year,
      startMonth: 6,
      endMonth: 7,
      year: year
    };
  }

  return {
    key: year + '-AUG-SEP',
    label: 'August–September ' + year,
    startMonth: 8,
    endMonth: 9,
    year: year
  };
}

function getFiestaShieldLevel_(score) {

  const value =
    Number(score || 0);

  if (value >= 90) {
    return {
      level: 'GOLD',
      label: '🥇 Gold Shield'
    };
  }

  if (value >= 80) {
    return {
      level: 'SILVER',
      label: '🥈 Silver Shield'
    };
  }

  if (value >= 70) {
    return {
      level: 'BRONZE',
      label: '🥉 Bronze Shield'
    };
  }

  return {
    level: 'PROGRESS',
    label: '🛡️ Building Your Shield'
  };
}

function getFiestaPunctualityScore_(
  studentId,
  periodStart,
  periodEnd
) {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      'Attendance Alerts'
    );

  if (!sheet || sheet.getLastRow() < 2) {
    return 100;
  }

  const rows =
    sheet
      .getDataRange()
      .getValues()
      .slice(1);

  let lateCount = 0;

  rows.forEach(function(row) {

    const rowStudentId =
      String(row[1] || '').trim();

    const scheduledDate =
      row[4];

    const alertType =
      String(row[6] || '')
        .trim()
        .toUpperCase();

    const status =
      String(row[11] || '')
        .trim()
        .toUpperCase();

    if (
      rowStudentId !==
      String(studentId).trim()
    ) {
      return;
    }

    if (!(scheduledDate instanceof Date)) {
      return;
    }

    if (
      scheduledDate < periodStart ||
      scheduledDate > periodEnd
    ) {
      return;
    }

    if (
      alertType === 'LATE' &&
      status !== 'EXCUSED'
    ) {
      lateCount++;
    }
  });

  return Math.max(
    0,
    100 - lateCount * 10
  );
}

function awardPerfectPerformanceAchievement_(
  studentId,
  studentName,
  periodLabel,
  performanceResult
) {

  if (
    !performanceResult ||
    !performanceResult.perfectPerformance
  ) {
    return {
      ok: true,
      awarded: false
    };
  }

  const achievementName =
    '💎 Perfect Performance - ' +
    periodLabel;

  return awardStudentAchievement_(
    studentId,
    studentName,
    achievementName,
    'Achieved 100% punctuality and 100% attendance for ' +
      periodLabel +
      '.',
    'FIESTA Hub'
  );
}

function awardFiestaShieldAchievement_(
  studentId,
  studentName,
  periodLabel,
  performanceResult
) {

  if (
    !performanceResult ||
    !performanceResult.shield
  ) {
    return {
      ok: true,
      awarded: false
    };
  }

  const shield =
    performanceResult.shield;

  if (
    shield.level === 'PROGRESS'
  ) {
    return {
      ok: true,
      awarded: false
    };
  }

  const achievementName =
    shield.label +
    ' - ' +
    periodLabel;

  const description =
    'Earned a ' +
    shield.label +
    ' with a performance score of ' +
    performanceResult.finalScore +
    '% for ' +
    periodLabel +
    '.';

  return awardStudentAchievement_(
    studentId,
    studentName,
    achievementName,
    description,
    'FIESTA Hub'
  );
}

function processFiestaAchievementPeriod_(
  periodStart,
  periodEnd,
  periodLabel
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const profilesSheet =
    ss.getSheetByName(
      'Student Profiles'
    );

  if (!profilesSheet) {
    throw new Error(
      'Student Profiles sheet was not found.'
    );
  }

  const rows =
    profilesSheet
      .getDataRange()
      .getValues()
      .slice(1);

  const results = [];

  rows.forEach(function(row) {

    const studentId =
      String(row[0] || '').trim();

    const studentName =
      String(row[1] || '').trim();

    const active =
      row[7] === true ||
      String(row[7] || '')
        .toLowerCase() === 'true';

    if (!studentId || !active) {
      return;
    }

const scheduledDays =
  getScheduledDaysInRange_(
    studentId,
    periodStart,
    periodEnd
  );

if (scheduledDays === 0) {
  return;
}

    const performance =
      getFiestaPerformanceScore_(
        studentId,
        periodStart,
        periodEnd
      );

    const shieldResult =
      awardFiestaShieldAchievement_(
        studentId,
        studentName,
        periodLabel,
        performance
      );

    const perfectResult =
      awardPerfectPerformanceAchievement_(
        studentId,
        studentName,
        periodLabel,
        performance
      );

    results.push({
      studentId: studentId,
      studentName: studentName,
      score: performance.finalScore,
      shield: performance.shield,
      perfectPerformance:
        performance.perfectPerformance,
      shieldAward:
        shieldResult,
      perfectAward:
        perfectResult
    });
  });

  return {
    ok: true,
    period: periodLabel,
    studentsProcessed:
      results.length,
    results: results
  };
}

function getPreviousFiestaPeriod_() {

  const today = new Date();

  const month =
    today.getMonth() + 1;

  const year =
    today.getFullYear();

  if (month === 1) {
    return {
      start: new Date(
        year - 1,
        9,
        1
      ),
      end: new Date(
        year - 1,
        11,
        31,
        23,
        59,
        59
      ),
      label:
        'October–December ' +
        (year - 1)
    };
  }

  if (month === 6) {
    return {
      start: new Date(
        year,
        0,
        1
      ),
      end: new Date(
        year,
        4,
        31,
        23,
        59,
        59
      ),
      label:
        'January–May ' +
        year
    };
  }

  if (month === 8) {
    return {
      start: new Date(
        year,
        5,
        1
      ),
      end: new Date(
        year,
        6,
        31,
        23,
        59,
        59
      ),
      label:
        'June–July ' +
        year
    };
  }

  if (month === 10) {
    return {
      start: new Date(
        year,
        7,
        1
      ),
      end: new Date(
        year,
        8,
        30,
        23,
        59,
        59
      ),
      label:
        'August–September ' +
        year
    };
  }

  return null;
}

function runFiestaPeriodAchievementProcessing_() {

  const period =
    getPreviousFiestaPeriod_();

  if (!period) {
    return {
      ok: true,
      skipped: true,
      message:
        'Today is not a FIESTA period closing month.'
    };
  }

  return processFiestaAchievementPeriod_(
    period.start,
    period.end,
    period.label
  );
}

function setupFiestaAchievementTrigger() {

  const functionName =
    'runFiestaPeriodAchievementProcessing_';

  const existingTriggers =
    ScriptApp.getProjectTriggers();

  existingTriggers.forEach(
    function(trigger) {

      if (
        trigger.getHandlerFunction() ===
        functionName
      ) {
        ScriptApp.deleteTrigger(
          trigger
        );
      }
    }
  );

  ScriptApp
    .newTrigger(functionName)
    .timeBased()
    .onMonthDay(1)
    .atHour(8)
    .create();

  return {
    ok: true,
    message:
      'FIESTA achievement monthly trigger created successfully.'
  };
}

function getCurrentFiestaWeek_() {

  const now = new Date();

  const dayOfWeek =
    now.getDay();

  const daysSinceMonday =
    (dayOfWeek + 6) % 7;

  const monday =
    new Date(
      now.getFullYear(),
      now.getMonth(),
      now.getDate() - daysSinceMonday
    );

  monday.setHours(
    0, 0, 0, 0
  );

  const friday =
    new Date(monday);

  friday.setDate(
    friday.getDate() + 4
  );

  friday.setHours(
    23, 59, 59, 999
  );

  return {
    start: monday,
    end: friday
  };
}
function getScheduledDaysInRange_(
  studentId,
  rangeStart,
  rangeEnd
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      'Regular Schedules'
    );

  if (!sheet || sheet.getLastRow() < 2) {
    return 0;
  }

  const rows =
    sheet
      .getDataRange()
      .getValues()
      .slice(1);

  const scheduledDays = {};

  const weekdayNames = [
    'Sunday',
    'Monday',
    'Tuesday',
    'Wednesday',
    'Thursday',
    'Friday',
    'Saturday'
  ];

  const currentDate =
    new Date(rangeStart);

  while (currentDate <= rangeEnd) {

    const weekdayName =
      weekdayNames[
        currentDate.getDay()
      ];

    const hasSchedule =
      rows.some(function(row) {

        const rowStudentId =
          String(row[1] || '')
            .trim();

        const rowDay =
          String(row[4] || '')
            .trim();

        const active =
          row[8] === true ||
          String(row[8] || '')
            .toLowerCase() === 'true';

        return (
          active &&
          rowStudentId ===
            String(studentId).trim() &&
          rowDay === weekdayName
        );
      });

    if (
  hasSchedule &&
  !hasApprovedScheduleException_(
    studentId,
    currentDate
  )
) {
  scheduledDays[
    currentDate.toDateString()
  ] = true;
}

    currentDate.setDate(
      currentDate.getDate() + 1
    );
  }

  const regularScheduledDays =
  Object.keys(
    scheduledDays
  ).length;

const approvedExceptionDays =
  getApprovedScheduleExceptionDaysInRange_(
    studentId,
    rangeStart,
    rangeEnd
  );

return (
  regularScheduledDays +
  approvedExceptionDays
);
}


function getScheduledDaysForWeek_(
  studentId,
  weekStart,
  weekEnd
) {

  return getScheduledDaysInRange_(
    studentId,
    weekStart,
    weekEnd
  );
}
function getFiestaWeeklyPerformance_(
  studentId,
  weekStart,
  weekEnd
) {

  const scheduledDays =
    getScheduledDaysForWeek_(
      studentId,
      weekStart,
      weekEnd
    );

  if (scheduledDays === 0) {
    return {
      hasSchedule: false,
      scheduledDays: 0
    };
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      'Attendance Alerts'
    );

  const rows =
    sheet &&
    sheet.getLastRow() >= 2
      ? sheet
          .getDataRange()
          .getValues()
          .slice(1)
      : [];

  let lateCount = 0;
  let missingCount = 0;

  rows.forEach(function(row) {

    const rowStudentId =
      String(row[1] || '')
        .trim();

    const scheduledDate =
      row[4];

    const alertType =
      String(row[6] || '')
        .trim()
        .toUpperCase();

    const status =
      String(row[11] || '')
        .trim()
        .toUpperCase();

    if (
      rowStudentId !==
      String(studentId).trim()
    ) {
      return;
    }

    if (!(scheduledDate instanceof Date)) {
      return;
    }

    if (
      scheduledDate < weekStart ||
      scheduledDate > weekEnd
    ) {
      return;
    }

    if (
  status === 'EXCUSED' ||
  status === 'WORKED'
) {
  return;
}
    if (alertType === 'LATE') {
      lateCount++;
    }

    if (alertType === 'MISSING') {
      missingCount++;
    }
  });

  return {
    hasSchedule: true,
    scheduledDays: scheduledDays,
    lateCount: lateCount,
    missingCount: missingCount
  };
}

function getFiestaWeeklyMessage_(
  weeklyPerformance
) {

  if (
    !weeklyPerformance ||
    !weeklyPerformance.hasSchedule
  ) {
    return null;
  }

  const lateCount =
    Number(
      weeklyPerformance.lateCount || 0
    );

  const missingCount =
    Number(
      weeklyPerformance.missingCount || 0
    );

  if (
    lateCount === 0 &&
    missingCount === 0
  ) {
    return {
      level: 'EXCELLENT',
      message:
        '🌟 Excellent work this week! Your attendance and punctuality were outstanding. Keep it up!'
    };
  }

  if (
    missingCount === 0 &&
    lateCount === 1
  ) {
    return {
      level: 'GOOD',
      message:
        '👏 Good job this week! You had one late arrival, but you stayed consistent. Keep pushing forward!'
    };
  }

  if (
    missingCount === 0 &&
    lateCount > 1
  ) {
    return {
      level: 'IMPROVE',
      message:
        '💪 This week had a few challenges with punctuality. Keep going — next week is a new opportunity to improve!'
    };
  }

  return {
    level: 'ENCOURAGEMENT',
    message:
      '🦁 This week may not have gone as planned, but don’t give up. Every new week is another chance to move forward!'
  };
}

function ensureWeeklyFeedbackSheet_() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  let sheet =
    ss.getSheetByName(
      'Weekly Feedback'
    );

  if (!sheet) {
    sheet =
      ss.insertSheet(
        'Weekly Feedback'
      );
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Feedback ID',
      'Student ID',
      'Student Name',
      'Week Start',
      'Week End',
      'Level',
      'Message',
      'Created On',
      'Active'
    ]);

    sheet.setFrozenRows(1);
  }

  return sheet;
}

function saveWeeklyFeedback_(
  studentId,
  studentName,
  weekStart,
  weekEnd,
  feedback
) {

  if (!feedback) {
    return {
      ok: true,
      saved: false
    };
  }

  const sheet =
    ensureWeeklyFeedbackSheet_();

  const tz =
    Session.getScriptTimeZone();

  const weekStartKey =
    Utilities.formatDate(
      weekStart,
      tz,
      'yyyy-MM-dd'
    );

  const rows =
    sheet
      .getDataRange()
      .getValues();

  for (let i = 1; i < rows.length; i++) {

    const rowStudentId =
      String(rows[i][1] || '').trim();

    const rowWeekStart =
      rows[i][3];

    const active =
      rows[i][8] === true ||
      String(rows[i][8] || '')
        .toLowerCase() === 'true';

    if (
      rowStudentId !==
      String(studentId).trim() ||
      !(rowWeekStart instanceof Date) ||
      !active
    ) {
      continue;
    }

    const rowWeekStartKey =
      Utilities.formatDate(
        rowWeekStart,
        tz,
        'yyyy-MM-dd'
      );

    if (
      rowWeekStartKey ===
      weekStartKey
    ) {
      return {
        ok: true,
        saved: false,
        duplicate: true
      };
    }
  }

  sheet.appendRow([
    Utilities.getUuid(),
    String(studentId || '').trim(),
    String(studentName || '').trim(),
    weekStart,
    weekEnd,
    String(feedback.level || ''),
    String(feedback.message || ''),
    new Date(),
    true
  ]);

  SpreadsheetApp.flush();

  return {
    ok: true,
    saved: true
  };
}

function processWeeklyFiestaFeedback_() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const profilesSheet =
    ss.getSheetByName(
      'Student Profiles'
    );

  if (!profilesSheet) {
    throw new Error(
      'Student Profiles sheet was not found.'
    );
  }

 const week =
  getCurrentFiestaWeek_();

  const rows =
    profilesSheet
      .getDataRange()
      .getValues()
      .slice(1);

  const results = [];

  rows.forEach(function(row) {

    const studentId =
      String(row[0] || '').trim();

    const studentName =
      String(row[1] || '').trim();

    const active =
      row[7] === true ||
      String(row[7] || '')
        .toLowerCase() === 'true';

    if (!studentId || !active) {
      return;
    }

    const performance =
      getFiestaWeeklyPerformance_(
        studentId,
        week.start,
        week.end
      );

    if (!performance.hasSchedule) {
      return;
    }

    const feedback =
      getFiestaWeeklyMessage_(
        performance
      );

    const saveResult =
      saveWeeklyFeedback_(
        studentId,
        studentName,
        week.start,
        week.end,
        feedback
      );

    results.push({
      studentId: studentId,
      studentName: studentName,
      performance: performance,
      feedback: feedback,
      saved: saveResult
    });
  });

  return {
    ok: true,
    studentsProcessed:
      results.length,
    results: results
  };
}

function setupWeeklyFiestaFeedbackTrigger() {

  const functionName =
    'processWeeklyFiestaFeedback_';

  const existingTriggers =
    ScriptApp.getProjectTriggers();

  existingTriggers.forEach(
    function(trigger) {

      if (
        trigger.getHandlerFunction() ===
        functionName
      ) {
        ScriptApp.deleteTrigger(
          trigger
        );
      }
    }
  );

  ScriptApp
    .newTrigger(functionName)
    .timeBased()
    .onWeekDay(
      ScriptApp.WeekDay.FRIDAY
    )
    .atHour(17)
    .create();

  return {
    ok: true,
    message:
      'Weekly FIESTA feedback trigger created successfully.'
  };
}
function getLatestWeeklyFeedback_(
  studentId
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      'Weekly Feedback'
    );

  if (!sheet || sheet.getLastRow() < 2) {
    return null;
  }

  const rows =
    sheet
      .getDataRange()
      .getValues()
      .slice(1);

  const matches =
    rows
      .filter(function(row) {

        const rowStudentId =
          String(row[1] || '').trim();

        const active =
          row[8] === true ||
          String(row[8] || '')
            .toLowerCase() === 'true';

        return (
          active &&
          rowStudentId ===
            String(studentId).trim()
        );
      })
      .sort(function(a, b) {

        const aDate =
          a[7] instanceof Date
            ? a[7].getTime()
            : 0;

        const bDate =
          b[7] instanceof Date
            ? b[7].getTime()
            : 0;

        return bDate - aDate;
      });

  if (!matches.length) {
    return null;
  }

  const row =
    matches[0];

  return {
    level:
      String(row[5] || ''),
    message:
      String(row[6] || ''),
    weekStart:
      row[3],
    weekEnd:
      row[4]
  };
}

function getStudentWeeklyFeedback(studentId) {

  if (!studentId) {
    throw new Error(
      'Student ID is required.'
    );
  }

  const feedback =
    getLatestWeeklyFeedback_(
      studentId
    );

  if (!feedback) {
    return {
      hasFeedback: false
    };
  }

  const tz =
    Session.getScriptTimeZone();

  return {
    hasFeedback: true,
    level:
      String(feedback.level || ''),
    message:
      String(feedback.message || ''),
    weekStart:
      feedback.weekStart instanceof Date
        ? Utilities.formatDate(
            feedback.weekStart,
            tz,
            'MMM d, yyyy'
          )
        : '',
    weekEnd:
      feedback.weekEnd instanceof Date
        ? Utilities.formatDate(
            feedback.weekEnd,
            tz,
            'MMM d, yyyy'
          )
        : ''
  };
}

function getFiestaAttendanceScore_(
  studentId,
  periodStart,
  periodEnd
) {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      'Attendance Alerts'
    );

  if (!sheet || sheet.getLastRow() < 2) {
    return 100;
  }

  const rows =
    sheet
      .getDataRange()
      .getValues()
      .slice(1);

  let absenceCount = 0;

  rows.forEach(function(row) {

    const rowStudentId =
      String(row[1] || '').trim();

    const scheduledDate =
      row[4];

    const alertType =
      String(row[6] || '')
        .trim()
        .toUpperCase();

    const status =
      String(row[11] || '')
        .trim()
        .toUpperCase();

    if (
      rowStudentId !==
      String(studentId).trim()
    ) {
      return;
    }

    if (!(scheduledDate instanceof Date)) {
      return;
    }

    if (
      scheduledDate < periodStart ||
      scheduledDate > periodEnd
    ) {
      return;
    }

  if (
  alertType === 'MISSING' &&
  status !== 'EXCUSED' &&
  status !== 'WORKED'
) {
  absenceCount++;
}
  });

  return Math.max(
    0,
    100 - absenceCount * 10
  );
}

function getFiestaPerformanceScore_(
  studentId,
  periodStart,
  periodEnd
) {

  const punctualityScore =
    getFiestaPunctualityScore_(
      studentId,
      periodStart,
      periodEnd
    );

  const attendanceScore =
    getFiestaAttendanceScore_(
      studentId,
      periodStart,
      periodEnd
    );

  const finalScore =
    (
      punctualityScore * 0.60
    ) +
    (
      attendanceScore * 0.40
    );

  return {
    punctualityScore:
      Math.round(
        punctualityScore * 100
      ) / 100,

    attendanceScore:
      Math.round(
        attendanceScore * 100
      ) / 100,

    finalScore:
      Math.round(
        finalScore * 100
      ) / 100,

    shield:
  getFiestaShieldLevel_(
    finalScore
  ),

perfectPerformance:
  punctualityScore === 100 &&
  attendanceScore === 100
  };
}

function getStudentCurrentFiestaProgress(studentId) {

  if (!studentId) {
    throw new Error(
      'Student ID is required.'
    );
  }

  const period =
    getFiestaAchievementPeriod_(
      new Date()
    );

  const year =
    Number(period.year);

  const start =
    new Date(
      year,
      Number(period.startMonth) - 1,
      1
    );

  const end =
    new Date();

   const scheduledDays =
  getScheduledDaysInRange_(
    studentId,
    start,
    end
  );

if (scheduledDays === 0) {
  return {
    periodLabel:
      String(period.label || ''),

    punctualityScore: null,
    attendanceScore: null,
    finalScore: null,

    shield: {
      level: 'PROGRESS',
      label: '🛡️ Building Your Shield'
    },

    perfectPerformance: false,
    hasProgress: false
  };
}

  const performance =
    getFiestaPerformanceScore_(
      studentId,
      start,
      end
    );

  return {

    hasProgress: true,

    periodLabel:
      String(period.label || ''),

    punctualityScore:
      Number(
        performance.punctualityScore || 0
      ),

    attendanceScore:
      Number(
        performance.attendanceScore || 0
      ),

    finalScore:
      Number(
        performance.finalScore || 0
      ),

    shield:
      performance.shield || {
        level: 'PROGRESS',
        label: '🛡️ Building Your Shield'
      },

    perfectPerformance:
      Boolean(
        performance.perfectPerformance
      )
  };
}

function hasStudentAchievement_(
  studentId,
  achievementName
) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Student Achievements');

  if (!sheet || sheet.getLastRow() < 2) {
    return false;
  }

  const cleanStudentId =
    String(studentId || '').trim();

  const cleanAchievementName =
    String(achievementName || '')
      .trim()
      .toLowerCase();

  const rows =
    sheet
      .getDataRange()
      .getValues()
      .slice(1);

  return rows.some(function(row) {

    const rowStudentId =
      String(row[1] || '').trim();

    const rowAchievementName =
      String(row[3] || '')
        .trim()
        .toLowerCase();

    const active =
      row[7] === true ||
      String(row[7] || '')
        .toLowerCase() === 'true';

    return (
      active &&
      rowStudentId === cleanStudentId &&
      rowAchievementName === cleanAchievementName
    );
  });
}

function awardStudentAchievement_(
  studentId,
  studentName,
  achievementName,
  description,
  awardedBy
) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Student Achievements');

  if (!sheet) {
    throw new Error('Student Achievements sheet was not found.');
  }

  if (!studentId || !achievementName) {
    throw new Error(
      'Student ID and achievement name are required.'
    );
  }

if (
  hasStudentAchievement_(
    studentId,
    achievementName
  )
) {
  return {
    ok: true,
    duplicate: true,
    message:
      achievementName +
      ' is already active for this student.'
  };
}

  sheet.appendRow([
    Utilities.getUuid(),
    String(studentId).trim(),
    String(studentName || '').trim(),
    String(achievementName).trim(),
    String(description || '').trim(),
    new Date(),
    String(awardedBy || 'FIESTA Hub').trim(),
    true
  ]);

  SpreadsheetApp.flush();

  return {
    ok: true,
    message:
      achievementName +
      ' was awarded successfully.'
  };
}

function testAwardAchievement() {
  awardStudentAchievement_(
    'TEST-STUDENT',
    'Test Student',
    '🏆 First Achievement',
    'Test achievement created by FIESTA Hub.',
    'FIESTA Hub'
  );
}

function getStudentCertificates_(studentId) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Student Certificates');

  if (!sheet || sheet.getLastRow() < 2) return [];

  return sheet.getDataRange().getValues()
    .slice(1)
    .filter(row =>
      String(row[1]) === String(studentId) &&
      (row[8] === true || String(row[8]).toLowerCase() === 'true')
    )
    .map(row => ({
      name: String(row[3] || ''),
      issued: row[4] instanceof Date
        ? Utilities.formatDate(
            row[4],
            Session.getScriptTimeZone(),
            'MMM d, yyyy'
          )
        : '',
      link: String(row[6] || '')
    }));
}
function getRegularShiftForActivity(
  studentId,
  activityId
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const tz =
    Session.getScriptTimeZone();

  const activitiesSheet =
    ss.getSheetByName('Activities');

  const schedulesSheet =
    ss.getSheetByName('Regular Schedules');

  if (
    !activitiesSheet ||
    activitiesSheet.getLastRow() < 2
  ) {
    throw new Error(
      'Activities sheet was not found.'
    );
  }

  if (
    !schedulesSheet ||
    schedulesSheet.getLastRow() < 2
  ) {
    return {
      found: false,
      shifts: [],
      message:
        'No regular schedules were found.'
    };
  }

  // -------------------------
  // FIND ACTIVITY
  // -------------------------

  const activityRows =
    activitiesSheet
      .getDataRange()
      .getValues();

  let activity = null;

  for (
    let i = 1;
    i < activityRows.length;
    i++
  ) {

    if (
      String(
        activityRows[i][0] || ''
      ).trim() ===
        String(activityId).trim()
    ) {

      activity = {
        activityId:
          String(activityRows[i][0] || ''),

        activityName:
          String(activityRows[i][1] || ''),

        date:
          activityRows[i][2],

        startTime:
          activityRows[i][3],

        endTime:
          activityRows[i][4]
      };

      break;
    }
  }

  if (
    !activity ||
    !(activity.date instanceof Date) ||
    !(activity.startTime instanceof Date) ||
    !(activity.endTime instanceof Date)
  ) {
    throw new Error(
      'Selected activity was not found or has invalid times.'
    );
  }

  const activityHours =
    Math.round(
      (
        (
          activity.endTime.getTime() -
          activity.startTime.getTime()
        ) / 3600000
      ) * 100
    ) / 100;

  // -------------------------
  // GET STUDENT'S ACTIVE
  // REGULAR WEEKLY SCHEDULES
  // -------------------------

  const scheduleRows =
    schedulesSheet
      .getDataRange()
      .getValues();

  const weeklySchedules = [];

  for (
    let i = 1;
    i < scheduleRows.length;
    i++
  ) {

    const rowStudentId =
      String(
        scheduleRows[i][1] || ''
      ).trim();

    const active =
      scheduleRows[i][8] === true ||
      String(
        scheduleRows[i][8] || ''
      ).toLowerCase() === 'true';

    if (
      !active ||
      rowStudentId !==
        String(studentId).trim()
    ) {
      continue;
    }

    const dayOfWeek =
      String(
        scheduleRows[i][4] || ''
      ).trim();

    const start =
      scheduleRows[i][5];

    const end =
      scheduleRows[i][6];

    if (
      !dayOfWeek ||
      !(start instanceof Date) ||
      !(end instanceof Date)
    ) {
      continue;
    }

    weeklySchedules.push({
      dayOfWeek:
        dayOfWeek,
      start:
        start,
      end:
        end
    });
  }

  if (!weeklySchedules.length) {
    return {
      found: false,
      shifts: [],
      activityHours:
        activityHours,
      message:
        'No active regular schedule was found for this student.'
    };
  }

  // -------------------------
  // BUILD ACTUAL SHIFT DATES
  // FOR THE ACTIVITY MONTH
  // -------------------------

  const year =
    activity.date.getFullYear();

  const month =
    activity.date.getMonth();

  const lastDay =
    new Date(
      year,
      month + 1,
      0
    ).getDate();

  const shifts = [];

  for (
    let day = 1;
    day <= lastDay;
    day++
  ) {

    const date =
      new Date(
        year,
        month,
        day
      );

    const weekday =
      Utilities.formatDate(
        date,
        tz,
        'EEEE'
      );

    weeklySchedules.forEach(
      function(schedule) {

        if (
          schedule.dayOfWeek
            .toLowerCase() !==
          weekday.toLowerCase()
        ) {
          return;
        }

        const startDateTime =
          new Date(
            year,
            month,
            day,
            schedule.start.getHours(),
            schedule.start.getMinutes(),
            0,
            0
          );

        const endDateTime =
          new Date(
            year,
            month,
            day,
            schedule.end.getHours(),
            schedule.end.getMinutes(),
            0,
            0
          );

        const shiftHours =
          Math.round(
            (
              (
                endDateTime.getTime() -
                startDateTime.getTime()
              ) / 3600000
            ) * 100
          ) / 100;

        shifts.push({
          shiftKey:
            Utilities.formatDate(
              date,
              tz,
              'yyyy-MM-dd'
            ) +
            '|' +
            Utilities.formatDate(
              startDateTime,
              tz,
              'HH:mm'
            ) +
            '|' +
            Utilities.formatDate(
              endDateTime,
              tz,
              'HH:mm'
            ),

          dateValue:
            Utilities.formatDate(
              date,
              tz,
              'yyyy-MM-dd'
            ),

          displayDate:
            Utilities.formatDate(
              date,
              tz,
              'EEE, MMM d'
            ),

          weekday:
            weekday,

          regularStartValue:
            Utilities.formatDate(
              startDateTime,
              tz,
              'HH:mm'
            ),

          regularEndValue:
            Utilities.formatDate(
              endDateTime,
              tz,
              'HH:mm'
            ),

          regularStart:
            Utilities.formatDate(
              startDateTime,
              tz,
              'h:mm a'
            ),

          regularEnd:
            Utilities.formatDate(
              endDateTime,
              tz,
              'h:mm a'
            ),

          availableHours:
            shiftHours
        });
      }
    );
  }

  return {
    found:
      shifts.length > 0,

    studentId:
      String(studentId),

    activityId:
      activity.activityId,

    activityName:
      activity.activityName,

    activityDate:
      Utilities.formatDate(
        activity.date,
        tz,
        'MMM d, yyyy'
      ),

    activityStart:
      Utilities.formatDate(
        activity.startTime,
        tz,
        'h:mm a'
      ),

    activityEnd:
      Utilities.formatDate(
        activity.endTime,
        tz,
        'h:mm a'
      ),

    activityHours:
      activityHours,

    shifts:
      shifts,

    message:
      shifts.length
        ? ''
        : 'No regular shifts were found in this month.'
  };
}
function saveActivityAssignment(formData) {

  if (!formData) {
    throw new Error(
      'Assignment information is required.'
    );
  }

  const studentId =
    String(
      formData.studentId || ''
    ).trim();

  const activityId =
    String(
      formData.activityId || ''
    ).trim();

  const scheduleTreatment =
    String(
      formData.scheduleTreatment || ''
    ).trim();

  const replacementAllocations =
    Array.isArray(
      formData.replacementAllocations
    )
      ? formData.replacementAllocations
      : [];

  if (
    !studentId ||
    !activityId ||
    !scheduleTreatment
  ) {
    throw new Error(
      'Student, activity, and treatment are required.'
    );
  }


  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const tz =
    Session.getScriptTimeZone();


  // -------------------------
  // STUDENT
  // -------------------------

  const student =
    getStudentForSchedule_(
      studentId
    );


  // -------------------------
  // ACTIVITY
  // -------------------------

  const activitiesSheet =
    ss.getSheetByName(
      'Activities'
    );

  if (
    !activitiesSheet ||
    activitiesSheet.getLastRow() < 2
  ) {
    throw new Error(
      'Activities sheet was not found.'
    );
  }

  const activityRows =
    activitiesSheet
      .getDataRange()
      .getValues();

  let activity = null;

  for (
    let i = 1;
    i < activityRows.length;
    i++
  ) {

    if (
      String(
        activityRows[i][0] || ''
      ).trim() ===
      activityId
    ) {

      activity = {
        activityId:
          String(
            activityRows[i][0] || ''
          ),

        activityName:
          String(
            activityRows[i][1] || ''
          ),

        date:
          activityRows[i][2],

        start:
          activityRows[i][3],

        end:
          activityRows[i][4],

        location:
          String(
            activityRows[i][5] || ''
          )
      };

      break;
    }
  }

  if (
    !activity ||
    !(activity.date instanceof Date) ||
    !(activity.start instanceof Date) ||
    !(activity.end instanceof Date)
  ) {
    throw new Error(
      'Selected activity was not found or has invalid date/time.'
    );
  }


  const activityHours =
    Math.round(
      (
        (
          activity.end.getTime() -
          activity.start.getTime()
        ) / 3600000
      ) * 100
    ) / 100;


  // -------------------------
  // SCHEDULED SHIFT RECORD
  // -------------------------

  let scheduledShiftsSheet =
    ss.getSheetByName(
      'Scheduled Shifts'
    );

  if (!scheduledShiftsSheet) {
    throw new Error(
      'Scheduled Shifts sheet was not found.'
    );
  }

  const shiftId =
    Utilities.getUuid();

  let hoursStatus =
    scheduleTreatment;

  if (
    scheduleTreatment ===
    'REPLACES REGULAR SHIFT'
  ) {
    hoursStatus =
      'REPLACES REGULAR SHIFT';
  }

  scheduledShiftsSheet.appendRow([
    shiftId,
    activity.activityId,
    activity.activityName,
    student.studentId,
    student.studentName,
    activity.date,
    activity.start,
    activity.end,
    activity.location,
    'SCHEDULED',
    hoursStatus,
    '',
    '',
    '',
    ''
  ]);


  // -------------------------
  // PARTIAL REPLACEMENTS
  // -------------------------

  if (
    scheduleTreatment ===
    'REPLACES REGULAR SHIFT'
  ) {

    if (
      !replacementAllocations.length
    ) {
      throw new Error(
        'Select the regular hours this activity will replace.'
      );
    }

    let totalAllocated = 0;

    replacementAllocations.forEach(
      function(allocation) {

        const replaceHours =
          Number(
            allocation.replaceHours || 0
          );

        const availableHours =
          Number(
            allocation.availableHours || 0
          );

        if (
          !Number.isFinite(
            replaceHours
          ) ||
          replaceHours <= 0
        ) {
          throw new Error(
            'Replacement hours must be greater than zero.'
          );
        }

        if (
          replaceHours >
          availableHours + 0.001
        ) {
          throw new Error(
            'Replacement hours cannot exceed the regular shift hours.'
          );
        }

        totalAllocated +=
          replaceHours;
      }
    );

    totalAllocated =
      Math.round(
        totalAllocated * 100
      ) / 100;

    if (
      Math.abs(
        totalAllocated -
        activityHours
      ) > 0.001
    ) {
      throw new Error(
        'Replacement allocations must equal the activity duration.'
      );
    }


    const exceptionsSheet =
      ss.getSheetByName(
        'Schedule Exceptions'
      );

    if (!exceptionsSheet) {
      throw new Error(
        'Schedule Exceptions sheet was not found.'
      );
    }

ensureScheduleExceptionApprovalStatus_();

    replacementAllocations.forEach(
      function(allocation) {

        const replacementDate =
          parsePortalDateTime_(
            allocation.date,
            '00:00'
          );

        const regularStart =
          parsePortalDateTime_(
            allocation.date,
            allocation.regularStart
          );

        const regularEnd =
          parsePortalDateTime_(
            allocation.date,
            allocation.regularEnd
          );

        const replaceHours =
          Number(
            allocation.replaceHours || 0
          );

        exceptionsSheet.appendRow([
          Utilities.getUuid(),
          student.studentId,
          student.studentName,
          replacementDate,
          'ACTIVITY PARTIAL REPLACEMENT',
          activity.activityId,
          activity.activityName,
          regularStart,
          regularEnd,
          activity.start,
          activity.end,
          replaceHours,
          getSettings_()
            .COORDINATOR_EMAIL ||
            'Administrator',
          new Date(),
          'Activity replaced ' +
  replaceHours.toFixed(2) +
  ' regular hour(s).',
false,
'PENDING'
        ]);
      }
    );
  }


  SpreadsheetApp.flush();


  return {
    ok: true,

    message:
  student.studentName +
  ' was assigned to ' +
  activity.activityName +
  '. The schedule change is pending approval.',

    shiftId:
      shiftId,

    activityHours:
      activityHours
  };
}
function ensureScheduleExceptionApprovalStatus_() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        'Schedule Exceptions'
      );

  if (!sheet) {
    throw new Error(
      'Schedule Exceptions sheet was not found.'
    );
  }

  const header =
    String(
      sheet.getRange(1, 17).getValue() || ''
    ).trim();

  if (header !== 'Approval Status') {
    sheet
      .getRange(1, 17)
      .setValue('Approval Status');
  }
}

function saveStudentEditProfile(formData) {

  if (!formData) {
    throw new Error(
      'Student information is required.'
    );
  }

  const studentId =
    String(formData.studentId || '').trim();

  const fullName =
    String(formData.fullName || '').trim();

  const studentNumber =
    String(formData.studentNumber || '').trim();

  const last4 =
    String(formData.last4 || '').trim();

  const email =
    String(formData.email || '').trim();

  const phone =
    String(formData.phone || '').trim();

  const address =
    String(formData.address || '').trim();

  const notes =
    String(formData.notes || '').trim();

  const newPin =
    String(formData.newPin || '').trim();


  if (
    !studentId ||
    !fullName ||
    !studentNumber ||
    !last4 ||
    !email
  ) {
    throw new Error(
      'Complete name, student number, last 4 SSN, and email.'
    );
  }

  if (!/^\d{4}$/.test(last4)) {
    throw new Error(
      'Last 4 SSN must contain exactly four numbers.'
    );
  }

  if (
    newPin &&
    !/^\d{4,8}$/.test(newPin)
  ) {
    throw new Error(
      'PIN must contain 4 to 8 numbers.'
    );
  }


  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const profileSheet =
    ss.getSheetByName(
      'Student Profiles'
    );

  const authSheet =
    ss.getSheetByName(
      'Students'
    );


  if (!profileSheet) {
    throw new Error(
      'Student Profiles sheet was not found.'
    );
  }

  if (!authSheet) {
    throw new Error(
      'Students authentication sheet was not found.'
    );
  }


  const profileRows =
    profileSheet
      .getDataRange()
      .getValues();

  let profileFound = false;

  const now =
    new Date();


  for (
    let i = 1;
    i < profileRows.length;
    i++
  ) {

    if (
      String(
        profileRows[i][0]
      ) === studentId
    ) {

      profileSheet
        .getRange(
          i + 1,
          2,
          1,
          10
        )
        .setValues([[
          fullName,
          studentNumber,
          last4,
          email,
          phone,
          address,
          true,
          profileRows[i][8] || now,
          now,
          notes
        ]]);

      profileFound = true;

      break;
    }
  }


  if (!profileFound) {
    throw new Error(
      'Student profile was not found.'
    );
  }


  // ==================================
  // UPDATE AUTHENTICATION RECORD
  // ==================================

  const authRows =
    authSheet
      .getDataRange()
      .getValues();

  let authFound = false;


  for (
    let i = 1;
    i < authRows.length;
    i++
  ) {

    if (
      String(
        authRows[i][0]
      ) === studentId
    ) {

      // Update name + email
      authSheet
        .getRange(
          i + 1,
          2
        )
        .setValue(
          fullName
        );

      authSheet
        .getRange(
          i + 1,
          3
        )
        .setValue(
          email
        );


      // Only reset PIN if a new one was entered
      if (newPin) {

        const newSalt =
          Utilities.getUuid();

        const newHash =
          hashText_(
            newSalt +
            newPin
          );

        authSheet
          .getRange(
            i + 1,
            4
          )
          .setValue(
            newSalt
          );

        authSheet
          .getRange(
            i + 1,
            5
          )
          .setValue(
            newHash
          );
      }

      authFound = true;

      break;
    }
  }


  if (!authFound) {
    throw new Error(
      'Student authentication record was not found.'
    );
  }


  SpreadsheetApp.flush();


  return {
    ok: true,

    message:
      fullName +
      (
        newPin
          ? "'s profile and PIN were updated."
          : "'s profile was updated."
      )
  };
}
function deactivateStudentProfile(
  studentId
) {

  studentId =
    String(
      studentId || ''
    ).trim();

  if (!studentId) {
    throw new Error(
      'Student ID is required.'
    );
  }


  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const profileSheet =
    ss.getSheetByName(
      'Student Profiles'
    );

  const authSheet =
    ss.getSheetByName(
      'Students'
    );


  if (!profileSheet) {
    throw new Error(
      'Student Profiles sheet was not found.'
    );
  }

  if (!authSheet) {
    throw new Error(
      'Students authentication sheet was not found.'
    );
  }


  let studentName =
    '';

  let profileFound =
    false;


  // ==================================
  // DEACTIVATE PROFILE
  // ==================================

  const profileRows =
    profileSheet
      .getDataRange()
      .getValues();

  for (
    let i = 1;
    i < profileRows.length;
    i++
  ) {

    if (
      String(
        profileRows[i][0] || ''
      ).trim() !== studentId
    ) {
      continue;
    }

    studentName =
      String(
        profileRows[i][1] || ''
      );

    // Active column H
    profileSheet
      .getRange(
        i + 1,
        8
      )
      .setValue(
        false
      );

    // Updated On column J
    profileSheet
      .getRange(
        i + 1,
        10
      )
      .setValue(
        new Date()
      );

    profileFound =
      true;

    break;
  }


  if (!profileFound) {
    throw new Error(
      'Student profile was not found.'
    );
  }


  // ==================================
  // DISABLE CLOCK-IN AUTH RECORD
  // ==================================

  const authRows =
    authSheet
      .getDataRange()
      .getValues();

  for (
    let i = 1;
    i < authRows.length;
    i++
  ) {

    if (
      String(
        authRows[i][0] || ''
      ).trim() !== studentId
    ) {
      continue;
    }

    // Students sheet Active = column F
    authSheet
      .getRange(
        i + 1,
        6
      )
      .setValue(
        false
      );

    break;
  }


  SpreadsheetApp.flush();


  return {
    ok: true,

    message:
      studentName +
      ' was deactivated. Historical records were preserved.'
  };
}

function getInactiveStudentProfiles() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName('Student Profiles');

  if (
    !sheet ||
    sheet.getLastRow() < 2
  ) {
    return [];
  }

  const rows =
    sheet.getDataRange().getValues();

  const students = [];

  for (
    let i = 1;
    i < rows.length;
    i++
  ) {

    const row = rows[i];

    const studentId =
      String(row[0] || '').trim();

    if (!studentId) {
      continue;
    }

    // Student Profiles:
    // A Student ID
    // B Full Name
    // C Student Number
    // D Last 4
    // E Email
    // F Phone
    // G Address
    // H Active

    const active =
      row[7] === true ||
      String(row[7]).toUpperCase() === 'TRUE';

    // We only want INACTIVE students
    if (active) {
      continue;
    }

    students.push({
      studentId: studentId,
      fullName:
        String(row[1] || ''),
      studentNumber:
        String(row[2] || ''),
      email:
        String(row[4] || ''),
      phone:
        String(row[5] || '')
    });
  }

  students.sort(function(a, b) {
    return a.fullName.localeCompare(
      b.fullName
    );
  });

  return students;
}

function reactivateStudentProfile(
  studentId,
  newPin
) {

  studentId =
    String(
      studentId || ''
    ).trim();

  newPin =
    String(
      newPin || ''
    ).trim();

  if (!studentId) {
    throw new Error(
      'Student ID is required.'
    );
  }

  if (
    !/^\d{4,8}$/.test(newPin)
  ) {
    throw new Error(
      'PIN must contain 4 to 8 numbers.'
    );
  }

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const profileSheet =
    ss.getSheetByName(
      'Student Profiles'
    );

  let authSheet =
    ss.getSheetByName(
      'Students'
    );

  if (!profileSheet) {
    throw new Error(
      'Student Profiles sheet was not found.'
    );
  }

  if (!authSheet) {
    throw new Error(
      'Students authentication sheet was not found.'
    );
  }


  let studentName = '';
  let studentEmail = '';
  let profileFound = false;


  // ==================================
  // REACTIVATE PROFILE
  // ==================================

  const profileRows =
    profileSheet
      .getDataRange()
      .getValues();

  for (
    let i = 1;
    i < profileRows.length;
    i++
  ) {

    if (
      String(
        profileRows[i][0] || ''
      ).trim() !==
        studentId
    ) {
      continue;
    }

    studentName =
      String(
        profileRows[i][1] || ''
      );

    studentEmail =
      String(
        profileRows[i][4] || ''
      );

    // Active = H
    profileSheet
      .getRange(
        i + 1,
        8
      )
      .setValue(true);

    // Updated On = J
    profileSheet
      .getRange(
        i + 1,
        10
      )
      .setValue(
        new Date()
      );

    profileFound = true;

    break;
  }


  if (!profileFound) {
    throw new Error(
      'Student profile was not found.'
    );
  }


  // ==================================
  // AUTHENTICATION RECORD
  // ==================================

  const authRows =
    authSheet
      .getDataRange()
      .getValues();

  let authFound = false;


  for (
    let i = 1;
    i < authRows.length;
    i++
  ) {

    if (
      String(
        authRows[i][0] || ''
      ).trim() !==
        studentId
    ) {
      continue;
    }

    const newSalt =
      Utilities.getUuid();

    const newHash =
      hashText_(
        newSalt +
        newPin
      );

    // Full Name
    authSheet
      .getRange(
        i + 1,
        2
      )
      .setValue(
        studentName
      );

    // Email
    authSheet
      .getRange(
        i + 1,
        3
      )
      .setValue(
        studentEmail
      );

    // New PIN Salt
    authSheet
      .getRange(
        i + 1,
        4
      )
      .setValue(
        newSalt
      );

    // New PIN Hash
    authSheet
      .getRange(
        i + 1,
        5
      )
      .setValue(
        newHash
      );

    // Active
    authSheet
      .getRange(
        i + 1,
        6
      )
      .setValue(true);

    authFound = true;

    break;
  }


  // ==================================
  // CREATE AUTH RECORD IF MISSING
  // ==================================

  if (!authFound) {

    const newSalt =
      Utilities.getUuid();

    const newHash =
      hashText_(
        newSalt +
        newPin
      );

    authSheet.appendRow([
      studentId,
      studentName,
      studentEmail,
      newSalt,
      newHash,
      true
    ]);
  }


  SpreadsheetApp.flush();


  return {
    ok: true,

    message:
      studentName +
      ' was reactivated and a new PIN was set.'
  };
}
function getStudentPortalActivities(studentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();

  const activitiesSheet = ss.getSheetByName('Activities');
  const shiftsSheet = ss.getSheetByName('Scheduled Shifts');
  const requestsSheet = ss.getSheetByName('Activity Participation Requests');

  if (!activitiesSheet || activitiesSheet.getLastRow() < 2) {
    return [];
  }

  const assignedActivityIds = {};
  const pendingActivityIds = {};

  if (shiftsSheet && shiftsSheet.getLastRow() >= 2) {
    const shiftRows = shiftsSheet.getDataRange().getValues();

    for (let i = 1; i < shiftRows.length; i++) {
      const rowStudentId = String(shiftRows[i][3] || '').trim();
      const activityId = String(shiftRows[i][1] || '').trim();

      if (rowStudentId === String(studentId).trim() && activityId) {
        assignedActivityIds[activityId] = true;
      }
    }
  }
if (requestsSheet && requestsSheet.getLastRow() >= 2) {
  const requestRows = requestsSheet.getDataRange().getValues();

  for (let i = 1; i < requestRows.length; i++) {
    const activityId = String(requestRows[i][1] || '').trim();
    const rowStudentId = String(requestRows[i][2] || '').trim();
    const status = String(requestRows[i][4] || '').trim().toUpperCase();

    if (
      rowStudentId === String(studentId).trim() &&
      activityId &&
      status === 'PENDING'
    ) {
      pendingActivityIds[activityId] = true;
    }
  }
}
  const rows = activitiesSheet.getDataRange().getValues();

const now = new Date();
now.setHours(0, 0, 0, 0);

  return rows
    .slice(1)
    .filter(function(row) {
  const status = String(row[8] || '').trim().toUpperCase();
  const activityDate = row[2] instanceof Date
    ? new Date(row[2])
    : null;

  if (!activityDate) {
    return false;
  }

  const firstDayOfCurrentMonth = new Date(
    now.getFullYear(),
    now.getMonth(),
    1
  );

  return (
    status === 'SCHEDULED' &&
    activityDate >= firstDayOfCurrentMonth
  );
})
    .map(function(row) {
      const activityId = String(row[0] || '').trim();

      const activityDate = row[2] instanceof Date
  ? new Date(row[2])
  : null;

if (activityDate) {
  activityDate.setHours(0, 0, 0, 0);
}

      return {
        activityId: activityId,
        activityName: String(row[1] || ''),
        date: row[2] instanceof Date
          ? Utilities.formatDate(row[2], tz, 'MMM d, yyyy')
          : '',
        start: row[3] instanceof Date
          ? Utilities.formatDate(row[3], tz, 'h:mm a')
          : '',
        end: row[4] instanceof Date
          ? Utilities.formatDate(row[4], tz, 'h:mm a')
          : '',
        location: String(row[5] || ''),
        description: String(row[6] || ''),
        studentsNeeded: Number(row[7] || 0),
participating: !!assignedActivityIds[activityId],
pending: !!pendingActivityIds[activityId],
past: activityDate ? activityDate < now : false
      };
    });
}
function requestActivityParticipation(
  studentId,
  activityId,
  hoursTreatment,
  replacementDetails
) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName('Activity Participation Requests');

  if (!sheet) {
    sheet = ss.insertSheet('Activity Participation Requests');

  sheet.appendRow([
  'Request ID',
  'Activity ID',
  'Student ID',
  'Requested On',
  'Status',
  'Hours Treatment',
  'Replacement Details'
]);
  }

  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const existingActivityId = String(rows[i][1] || '').trim();
    const existingStudentId = String(rows[i][2] || '').trim();
    const status = String(rows[i][4] || '').trim().toUpperCase();

    if (
      existingActivityId === String(activityId).trim() &&
      existingStudentId === String(studentId).trim() &&
      (status === 'PENDING' || status === 'APPROVED')
    ) {
      return {
        ok: false,
        message: status === 'APPROVED'
          ? 'You are already participating in this activity.'
          : 'Your participation request is already pending.'
      };
    }
  }

sheet.appendRow([
  Utilities.getUuid(),
  String(activityId).trim(),
  String(studentId).trim(),
  new Date(),
  'PENDING',
  String(hoursTreatment || '').trim(),
  String(replacementDetails || '').trim()
]);

  SpreadsheetApp.flush();

  return {
    ok: true,
    message: 'Your participation request was sent for approval.'
  };
}

function hasApprovedScheduleException_(
  studentId,
  targetDate
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      'Schedule Exceptions'
    );

  if (!sheet || sheet.getLastRow() < 2) {
    return false;
  }

  const tz =
    Session.getScriptTimeZone();

  const targetDateKey =
    Utilities.formatDate(
      targetDate,
      tz,
      'yyyy-MM-dd'
    );

  const rows =
    sheet
      .getDataRange()
      .getValues()
      .slice(1);

  return rows.some(function(row) {

    const rowStudentId =
      String(row[1] || '').trim();

    const exceptionDate =
      row[3];

    const active =
      row[15] === true ||
      String(row[15] || '')
        .toLowerCase() === 'true';

    const approvalStatus =
      String(row[16] || '')
        .trim()
        .toUpperCase();

const regularStart =
  row[7];

const regularEnd =
  row[8];

const replacedHours =
  Number(row[11] || 0);

    if (
      rowStudentId !==
        String(studentId).trim() ||
      !(exceptionDate instanceof Date) ||
      !active ||
      approvalStatus !== 'APPROVED'
    ) {
      return false;
    }

    const exceptionDateKey =
      Utilities.formatDate(
        exceptionDate,
        tz,
        'yyyy-MM-dd'
      );

    const regularHours =
  (
    regularStart instanceof Date &&
    regularEnd instanceof Date
  )
    ? (
        regularEnd.getTime() -
        regularStart.getTime()
      ) / 3600000
    : 0;

const isFullReplacement =
  regularHours > 0 &&
  replacedHours >=
    regularHours - 0.001;

return (
  exceptionDateKey ===
    targetDateKey &&
  isFullReplacement
);
  });
}

function getApprovedScheduleExceptionDaysInRange_(
  studentId,
  rangeStart,
  rangeEnd
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      'Schedule Exceptions'
    );

  if (!sheet || sheet.getLastRow() < 2) {
    return 0;
  }

  const tz =
    Session.getScriptTimeZone();

  const rows =
    sheet
      .getDataRange()
      .getValues()
      .slice(1);

  const exceptionDays = {};

  rows.forEach(function(row) {

    const rowStudentId =
      String(row[1] || '').trim();

    const exceptionDate =
      row[3];

    const active =
      row[15] === true ||
      String(row[15] || '')
        .toLowerCase() === 'true';

    const regularStart =
  row[7];

const regularEnd =
  row[8];

const replacedHours =
  Number(row[11] || 0);

    if (
      rowStudentId !==
        String(studentId).trim() ||
      !(exceptionDate instanceof Date) ||
      !active ||
      approvalStatus !== 'APPROVED'
    ) {
      return;
    }
const regularHours =
  (
    regularStart instanceof Date &&
    regularEnd instanceof Date
  )
    ? (
        regularEnd.getTime() -
        regularStart.getTime()
      ) / 3600000
    : 0;

const isFullReplacement =
  regularHours > 0 &&
  replacedHours >=
    regularHours - 0.001;

if (!isFullReplacement) {
  return;
}
    if (
      exceptionDate < rangeStart ||
      exceptionDate > rangeEnd
    ) {
      return;
    }

    const dateKey =
      Utilities.formatDate(
        exceptionDate,
        tz,
        'yyyy-MM-dd'
      );

    exceptionDays[dateKey] = true;
  });

  return Object.keys(
    exceptionDays
  ).length;
}

function getCompletedActivityCountForPeriod_(
  studentId,
  periodStart,
  periodEnd
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      'Scheduled Shifts'
    );

  if (!sheet || sheet.getLastRow() < 2) {
    return 0;
  }

  const now =
    new Date();

  const rows =
    sheet
      .getDataRange()
      .getValues()
      .slice(1);

  const completedActivities = {};

  rows.forEach(function(row) {

    const activityId =
      String(row[1] || '').trim();

    const rowStudentId =
      String(row[3] || '').trim();

    const scheduledDate =
      row[5];

    if (
      rowStudentId !==
        String(studentId).trim() ||
      !activityId ||
      !(scheduledDate instanceof Date)
    ) {
      return;
    }

    if (
      scheduledDate < periodStart ||
      scheduledDate > periodEnd ||
      scheduledDate > now
    ) {
      return;
    }

    completedActivities[
      activityId
    ] = true;
  });

  return Object.keys(
    completedActivities
  ).length;
}

function getFiestaActivityLevel_(
  activityCount,
  periodStart
) {

  const count =
    Number(activityCount || 0);

  const month =
    periodStart instanceof Date
      ? periodStart.getMonth() + 1
      : 0;

  const isLongPeriod =
    month === 10 ||
    month === 1;

  if (isLongPeriod) {

    if (count >= 10) {
      return {
        level: 'CHAMPION',
        label: '🦁 FIESTA Champion'
      };
    }

    if (count >= 6) {
      return {
        level: 'SUPERFAN',
        label: '🔥 FIESTA Superfan'
      };
    }

    if (count >= 3) {
      return {
        level: 'ACTIVE',
        label: '⭐ FIESTA Active'
      };
    }

    if (count >= 1) {
      return {
        level: 'ROOKIE',
        label: '🎉 Activity Rookie'
      };
    }

  } else {

    if (count >= 6) {
      return {
        level: 'CHAMPION',
        label: '🦁 FIESTA Champion'
      };
    }

    if (count >= 4) {
      return {
        level: 'SUPERFAN',
        label: '🔥 FIESTA Superfan'
      };
    }

    if (count >= 2) {
      return {
        level: 'ACTIVE',
        label: '⭐ FIESTA Active'
      };
    }

    if (count >= 1) {
      return {
        level: 'ROOKIE',
        label: '🎉 Activity Rookie'
      };
    }
  }

  return {
    level: 'NONE',
    label: '🌱 Ready to Participate'
  };
}

function getStudentCurrentActivityProgress_(
  studentId
) {

  const period =
    getFiestaAchievementPeriod_(
      new Date()
    );

  const periodStart =
    new Date(
      Number(period.year),
      Number(period.startMonth) - 1,
      1
    );

  const now =
    new Date();

  const activityCount =
    getCompletedActivityCountForPeriod_(
      studentId,
      periodStart,
      now
    );

  const activityLevel =
    getFiestaActivityLevel_(
      activityCount,
      periodStart
    );

  return {
    periodKey:
      String(period.key || ''),

    periodLabel:
      String(period.label || ''),

    activityCount:
      activityCount,

    level:
      activityLevel.level,

    label:
      activityLevel.label
  };
}
function getFiestaActivityRanking_() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const profileSheet =
    ss.getSheetByName(
      'Student Profiles'
    );

  if (
    !profileSheet ||
    profileSheet.getLastRow() < 2
  ) {
    return [];
  }

  const students =
    profileSheet
      .getDataRange()
      .getValues()
      .slice(1);

  const ranking = [];

  students.forEach(function(row) {

    const studentId =
      String(row[0] || '').trim();

    const studentName =
      String(row[1] || '').trim();

    const active =
      row[7] === true ||
      String(row[7] || '')
        .toLowerCase() === 'true';

    if (
      !studentId ||
      !active
    ) {
      return;
    }

    const progress =
      getStudentCurrentActivityProgress_(
        studentId
      );

if (
  progress.activityCount <= 0
) {
  return;
}

    ranking.push({
      studentId: studentId,
      studentName: studentName,
      activityCount:
        progress.activityCount,
      level:
        progress.level,
      label:
        progress.label,
      periodKey:
        progress.periodKey,
      periodLabel:
        progress.periodLabel
    });
  });

  ranking.sort(function(a, b) {

    if (
      b.activityCount !==
      a.activityCount
    ) {
      return (
        b.activityCount -
        a.activityCount
      );
    }

    return a.studentName.localeCompare(
      b.studentName
    );
  });

  let currentRank = 0;
let previousCount = null;

ranking.forEach(function(student, index) {

  if (
    previousCount === null ||
    student.activityCount !== previousCount
  ) {
    currentRank = index + 1;
  }

  student.rank = currentRank;

  previousCount =
    student.activityCount;
});

  return ranking;
}

function getStudentActivityRanking() {

  const ranking =
    getFiestaActivityRanking_();

  const period =
    getFiestaAchievementPeriod_(
      new Date()
    );

  return {
    periodKey:
      String(period.key || ''),

    periodLabel:
      String(period.label || ''),

    ranking:
      ranking
  };
}


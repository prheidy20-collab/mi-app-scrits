function getStudentMonthlyAttendanceData_(studentId, monthText) {
  if (!studentId) {
    throw new Error('Student ID is required.');
  }

  if (!/^\d{4}-\d{2}$/.test(String(monthText))) {
    throw new Error('Month must use YYYY-MM format.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();

  const student =
    getStudentProfileDetail(studentId);

  if (!student) {
    throw new Error('Student not found.');
  }

  const shifts = [];

const attendanceCorrections =
  getAttendanceCorrections_(
    studentId,
    monthText
  );

  // ===================================
  // 1. REGULAR TABLET CLOCK ENTRIES
  // ===================================
  const timeEntriesSheet =
    ss.getSheetByName('Time Entries');

  if (
    timeEntriesSheet &&
    timeEntriesSheet.getLastRow() >= 2
  ) {
    const rows =
      timeEntriesSheet
        .getDataRange()
        .getValues()
        .slice(1);

    rows.forEach(row => {
      const rowStudentId =
        String(row[1] || '').trim();

      const clockIn = row[3];
      const clockOut = row[4];

      const storedHours =
        Number(row[5]);

const entryId =
  String(row[0] || '').trim();

const correction =
  attendanceCorrections[entryId] || null;

      const status =
        String(row[7] || '')
          .trim()
          .toUpperCase();

const notes =
  String(row[8] || '').trim();

let activityContext = null;

if (notes) {
  try {
    activityContext = JSON.parse(notes);
  } catch (error) {
    activityContext = null;
  }
}

      if (
        rowStudentId !==
          String(studentId).trim() ||
        !(clockIn instanceof Date) ||
        !(clockOut instanceof Date) ||
        status !== 'COMPLETE'
      ) {
        return;
      }

      const rowMonth =
        Utilities.formatDate(
          clockIn,
          tz,
          'yyyy-MM'
        );

      if (rowMonth !== monthText) {
        return;
      }

      if (
  activityContext &&
  activityContext.type === 'ACTIVITY' &&
  (
    activityContext.hoursTreatment === 'VOLUNTEER HOURS' ||
    activityContext.hoursTreatment === 'REPLACES REGULAR SHIFT'
  )
) {
  return;
}

let effectiveIn = clockIn;
let effectiveOut = clockOut;
let hours = storedHours;
let source = 'RECORDED';

if (
  correction &&
  correction.approvedIn instanceof Date &&
  correction.approvedOut instanceof Date
) {
  effectiveIn =
    correction.approvedIn;

  effectiveOut =
    correction.approvedOut;

  hours =
    Number(correction.approvedHours);

  source = 'CORRECTED';
}

if (!Number.isFinite(hours)) {
  hours =
    (
      effectiveOut.getTime() -
      effectiveIn.getTime()
    ) / 3600000;
}

shifts.push({
  entryId: entryId,
  source: source,
  workDate: effectiveIn,
  startTime: effectiveIn,
  endTime: effectiveOut,
  hours:
    Math.round(hours * 100) / 100
});

    });

  }

  // ===================================
  // 2. MANUAL SHIFT ADJUSTMENTS
  // ===================================
  const adjustmentSheet =
    ss.getSheetByName(
      'Payroll Shift Adjustments'
    );

  if (
    adjustmentSheet &&
    adjustmentSheet.getLastRow() >= 2
  ) {
    const rows =
      adjustmentSheet
        .getDataRange()
        .getValues()
        .slice(1);

    rows.forEach(row => {
      const rowStudentId =
        String(row[1] || '').trim();

      const workDate = row[3];
      const startTime = row[4];
      const endTime = row[5];

      const hours =
        Number(row[6]);

      const active =
        row[10] === true ||
        String(row[10] || '')
          .toLowerCase() === 'true';

      if (
        !active ||
        rowStudentId !==
          String(studentId).trim() ||
        !(workDate instanceof Date)
      ) {
        return;
      }

      const rowMonth =
        Utilities.formatDate(
          workDate,
          tz,
          'yyyy-MM'
        );

      if (rowMonth !== monthText) {
        return;
      }

      shifts.push({
        source: 'ADJUSTMENT',
        workDate: workDate,
        startTime: startTime,
        endTime: endTime,
        hours:
          Number.isFinite(hours)
            ? hours
            : 0
      });
    });
  }
// ===================================
// 3. ATTENDANCE MANUAL SHIFTS
// ===================================

const attendanceManualShifts =
  getAttendanceManualShifts_(
    studentId,
    monthText
  );

attendanceManualShifts.forEach(function(shift) {

  if (
    !(shift.workDate instanceof Date) ||
    !(shift.startTime instanceof Date) ||
    !(shift.endTime instanceof Date)
  ) {
    return;
  }

  shifts.push({
    shiftId:
      String(shift.shiftId || ''),

    source:
      'MANUAL',

    workDate:
      shift.workDate,

    startTime:
      shift.startTime,

    endTime:
      shift.endTime,

    hours:
      Number(shift.hours || 0),

    reason:
      String(shift.reason || '')
  });
});
  // ===================================
  // 4. APPROVED ACTIVITY REPLACEMENTS
  // ===================================

  const activityReplacementShifts =
    getActivityReplacementShifts_(
      studentId,
      monthText
    );

  activityReplacementShifts.forEach(function(shift) {
    shifts.push({
      source:
        'ACTIVITY_REPLACEMENT',

      workDate:
        shift.replacementDate,

      startTime:
        shift.regularStart,

      endTime:
        shift.regularEnd,

      hours:
        Number(shift.hours || 0),

      activityId:
        shift.activityId,

      activityName:
        shift.activityName
    });
  });
  // ===================================
  // 5. SORT SHIFTS
  // ===================================
  shifts.sort((a, b) => {
    const aDate =
      a.startTime instanceof Date
        ? a.startTime.getTime()
        : a.workDate.getTime();

    const bDate =
      b.startTime instanceof Date
        ? b.startTime.getTime()
        : b.workDate.getTime();

    return aDate - bDate;
  });

  // ===================================
  // 6. TOTAL HOURS
  // ===================================
  const totalHours =
    Math.round(
      shifts.reduce(
        (sum, shift) =>
          sum + Number(shift.hours || 0),
        0
      ) * 100
    ) / 100;

  return {
    studentId: String(studentId),

    studentName:
      student.fullName ||
      student.studentName ||
      '',

    studentNumber:
      student.studentNumber ||
      '',

    last4:
      student.last4 ||
      student.last4SSN ||
      '',

    month: monthText,

    shifts: shifts,

    totalHours: totalHours
  };
}
function getActivityReplacementShifts_(
  studentId,
  monthText
) {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const tz =
    Session.getScriptTimeZone();

  const sheet =
    ss.getSheetByName('Schedule Exceptions');

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const rows =
    sheet.getDataRange()
      .getValues()
      .slice(1);

  const replacements = [];
const seenReplacementKeys = {};
  rows.forEach(function(row) {

    const rowStudentId =
      String(row[1] || '').trim();

    const exceptionDate = row[3];

    const exceptionType =
      String(row[4] || '')
        .trim()
        .toUpperCase();

    const activityId =
      String(row[5] || '').trim();

    const activityName =
      String(row[6] || '').trim();

    const regularStart = row[7];
    const regularEnd = row[8];

    const activityStart = row[9];
    const activityEnd = row[10];

    const paidHours =
      Number(row[11] || 0);

    const active =
      row[15] === true ||
      String(row[15] || '')
        .toLowerCase() === 'true';

const approvalStatus =
  String(row[16] || '')
    .trim()
    .toUpperCase();

    if (
  !active ||
  (
    approvalStatus &&
    approvalStatus !== 'APPROVED'
  ) ||
      rowStudentId !==
        String(studentId).trim() ||
      !(exceptionDate instanceof Date) ||
      !Number.isFinite(paidHours) ||
      paidHours <= 0
    ) {
      return;
    }

    const rowMonth =
      Utilities.formatDate(
        exceptionDate,
        tz,
        'yyyy-MM'
      );

    if (rowMonth !== monthText) {
      return;
    }

const replacementKey = [
  activityId,
  exceptionDate.getTime(),
  regularStart instanceof Date
    ? regularStart.getTime()
    : String(regularStart || ''),
  regularEnd instanceof Date
    ? regularEnd.getTime()
    : String(regularEnd || ''),
  Math.round(paidHours * 100) / 100
].join('|');

if (seenReplacementKeys[replacementKey]) {
  return;
}

seenReplacementKeys[replacementKey] = true;

    replacements.push({
      activityId: activityId,
      activityName: activityName,

      replacementDate:
        exceptionDate,

      regularStart:
        regularStart,

      regularEnd:
        regularEnd,

      activityStart:
        activityStart,

      activityEnd:
        activityEnd,

      hours:
        Math.round(paidHours * 100) / 100,

      source:
        'ACTIVITY_REPLACEMENT',

      exceptionType:
        exceptionType
    });
  });

  return replacements;
}

function debugJuanitoAttendancePDF() {
  const data =
    getStudentMonthlyAttendanceData_(
      'STU-123456',
      '2026-08'
    );

  Logger.log(
    JSON.stringify(data, null, 2)
  );
}
function buildMonthlyAttendanceReportData_(studentId, monthText) {
  const attendance =
    getStudentMonthlyAttendanceData_(
      studentId,
      monthText
    );

  const approvals =
    getExistingPayrollApprovals_(
      monthText
    );

  const approval =
    approvals[String(studentId)] || null;

  if (!approval) {
    throw new Error(
      'This student does not have a payroll record for the selected month.'
    );
  }

  if (
    String(approval.status || '')
      .toUpperCase() !== 'APPROVED'
  ) {
    throw new Error(
      'Payroll must be APPROVED before an official attendance report can be generated.'
    );
  }

  const approvedHours =
    Number(approval.approvedHours || 0);

  const tz =
    Session.getScriptTimeZone();

  // -------------------------------
  // GROUP SHIFTS BY DATE
  // -------------------------------
  const days = {};

  attendance.shifts.forEach(shift => {
    if (!(shift.workDate instanceof Date)) {
      return;
    }

    const dateKey =
      Utilities.formatDate(
        shift.workDate,
        tz,
        'yyyy-MM-dd'
      );

    if (!days[dateKey]) {
      days[dateKey] = {
        date: new Date(shift.workDate),
        shifts: [],
        totalHours: 0
      };
    }

    days[dateKey].shifts.push({
      source: shift.source,

      start:
        shift.startTime instanceof Date
          ? Utilities.formatDate(
              shift.startTime,
              tz,
              'h:mm a'
            )
          : '',

      end:
        shift.endTime instanceof Date
          ? Utilities.formatDate(
              shift.endTime,
              tz,
              'h:mm a'
            )
          : '',

      hours:
        Number(shift.hours || 0)
    });

    days[dateKey].totalHours +=
      Number(shift.hours || 0);
  });

  // Sort shifts within each day
  Object.keys(days).forEach(key => {
    days[key].shifts.sort((a, b) =>
      String(a.start).localeCompare(
        String(b.start)
      )
    );

    days[key].totalHours =
      Math.round(
        days[key].totalHours * 100
      ) / 100;
  });

  // -------------------------------
  // BUILD WEEK BLOCKS
  // -------------------------------
  const year =
    Number(monthText.substring(0, 4));

  const month =
    Number(monthText.substring(5, 7)) - 1;

  const firstDay =
    new Date(year, month, 1);

  const lastDay =
    new Date(year, month + 1, 0);

  const firstMonday =
    getMondayForAttendance_(
      firstDay
    );

  const weeks = [];

  let cursor =
    new Date(firstMonday);

  while (cursor <= lastDay) {
    const weekStart =
      new Date(cursor);

    const weekEnd =
      new Date(cursor);

    weekEnd.setDate(
      weekEnd.getDate() + 4
    );
// Skip a Monday-Friday block that has
// no weekdays inside the selected month.
if (
  weekEnd < firstDay ||
  weekStart > lastDay
) {

  cursor.setDate(
    cursor.getDate() + 7
  );

  continue;
}
const displayStart =
  weekStart < firstDay
    ? new Date(firstDay)
    : new Date(weekStart);

const displayEnd =
  weekEnd > lastDay
    ? new Date(lastDay)
    : new Date(weekEnd);

    const week = {

      startDate:
  Utilities.formatDate(
    weekStart,
    tz,
    'yyyy-MM-dd'
  ),

endDate:
  Utilities.formatDate(
    weekEnd,
    tz,
    'yyyy-MM-dd'
  ),

      weekStart:
        Utilities.formatDate(
          weekStart,
          tz,
          'MMM d, yyyy'
        ),

     weekLabel:
  Utilities.formatDate(
    displayStart,
    tz,
    'MMM d'
  ) ===
  Utilities.formatDate(
    displayEnd,
    tz,
    'MMM d'
  )

    ? Utilities.formatDate(
        displayStart,
        tz,
        'MMM d'
      )

    : Utilities.formatDate(
        displayStart,
        tz,
        'MMM d'
      ) +
      ' - ' +
      Utilities.formatDate(
        displayEnd,
        tz,
        'MMM d'
      ),

      days: []
    };

    const labels = [
      'L',
      'M',
      'W',
      'J',
      'V'
    ];

let weekTotalHours = 0;

    for (let i = 0; i < 5; i++) {
      const date =
        new Date(weekStart);

      date.setDate(
        weekStart.getDate() + i
      );

      const dateKey =
        Utilities.formatDate(
          date,
          tz,
          'yyyy-MM-dd'
        );

      const dateData =
        days[dateKey] || null;

if (dateData) {
  weekTotalHours +=
    Number(
      dateData.totalHours || 0
    );
}
      const shiftPairs = [];

      if (dateData) {
        dateData.shifts
          .slice(0, 3)
          .forEach(shift => {
            shiftPairs.push({
              entrada:
                shift.start || '',
              salida:
                shift.end || ''
            });
          });
      }

      while (shiftPairs.length < 3) {
        shiftPairs.push({
          entrada: '',
          salida: ''
        });
      }

      week.days.push({
        dayLabel: labels[i],

        date:
          Utilities.formatDate(
            date,
            tz,
            'M/d'
          ),

        inMonth:
          date.getMonth() === month,

        shifts:
          shiftPairs,

        totalHours:
          dateData
            ? Number(
                dateData.totalHours
              ).toFixed(2)
            : ''
      });
    }
week.totalHours =
  Math.round(
    weekTotalHours * 100
  ) / 100;
    weeks.push(week);

    cursor.setDate(
      cursor.getDate() + 7
    );
  }

  // Four week blocks per page
  const pages = [];

  for (
    let i = 0;
    i < weeks.length;
    i += 4
  ) {
    pages.push(
      weeks.slice(i, i + 4)
    );
  }

  const detailedHours =
    Math.round(
      attendance.totalHours * 100
    ) / 100;

  const difference =
    Math.round(
      (
        approvedHours -
        detailedHours
      ) * 100
    ) / 100;

  return {
    studentId:
      attendance.studentId,

    studentName:
      attendance.studentName,

    studentNumber:
      attendance.studentNumber,

    last4:
      attendance.last4,

    month:
      monthText,

    monthName:
      Utilities.formatDate(
        new Date(year, month, 1),
        tz,
        'MMMM yyyy'
      ),

    approvedHours:
      approvedHours,

    detailedHours:
      detailedHours,

    difference:
      difference,

    readyForOfficialPDF:
      Math.abs(difference) < 0.01,

    pages:
      pages
  };
}


function getMondayForAttendance_(date) {
  const copy =
    new Date(
      date.getFullYear(),
      date.getMonth(),
      date.getDate()
    );

  const day =
    copy.getDay();

  const difference =
    day === 0
      ? -6
      : 1 - day;

  copy.setDate(
    copy.getDate() + difference
  );

  return copy;
}


function getMonthlyAttendancePreview(
  studentId,
  monthText
) {

  return buildMonthlyAttendanceReportData_(
    studentId,
    monthText
  );
}

/**
 * Returns the Monday-Friday payroll weeks
 * available for a selected payroll month.
 *
 * A week belongs to the payroll month if
 * at least one weekday in that week falls
 * inside the selected month.
 *
 * Example:
 * September 2026 includes:
 * Aug 31 - Sep 4
 */
function getPayrollWeeksForMonth(monthText) {

  if (!/^\d{4}-\d{2}$/.test(String(monthText || ''))) {
    throw new Error(
      'Month must use YYYY-MM format.'
    );
  }

  const parts =
    monthText.split('-');

  const year =
    Number(parts[0]);

  const monthIndex =
    Number(parts[1]) - 1;

  const firstDay =
    new Date(
      year,
      monthIndex,
      1
    );

  const lastDay =
    new Date(
      year,
      monthIndex + 1,
      0
    );


  // Find Monday of the week
  // containing the first day of month.
  const firstDayNumber =
    firstDay.getDay();

  const daysBackToMonday =
    firstDayNumber === 0
      ? 6
      : firstDayNumber - 1;

  let monday =
    new Date(firstDay);

  monday.setDate(
    monday.getDate() -
    daysBackToMonday
  );


  const weeks = [];

  let weekNumber = 1;


  while (monday <= lastDay) {

    const friday =
      new Date(monday);

    friday.setDate(
      monday.getDate() + 4
    );


    // A payroll week belongs to the month
// where Friday falls.
const touchesMonth =
  friday.getFullYear() === year &&
  friday.getMonth() === monthIndex;


    if (touchesMonth) {

      weeks.push({
        id:
          formatPayrollDateKey_(monday) +
          '_' +
          formatPayrollDateKey_(friday),

        weekNumber:
          weekNumber,

        startDate:
          formatPayrollDateKey_(monday),

        endDate:
          formatPayrollDateKey_(friday),

        label:
          formatPayrollWeekLabel_(
            monday,
            friday
          )
      });

      weekNumber++;
    }


    monday =
      new Date(monday);

    monday.setDate(
      monday.getDate() + 7
    );
  }


  return weeks;
}


/**
 * YYYY-MM-DD
 */
function formatPayrollDateKey_(date) {

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}


/**
 * Example:
 * Aug 31 - Sep 4, 2026
 */
function formatPayrollWeekLabel_(
  startDate,
  endDate
) {

  const tz =
    Session.getScriptTimeZone();

  const startMonth =
    Utilities.formatDate(
      startDate,
      tz,
      'MMM'
    );

  const startDay =
    Utilities.formatDate(
      startDate,
      tz,
      'd'
    );

  const endMonth =
    Utilities.formatDate(
      endDate,
      tz,
      'MMM'
    );

  const endDay =
    Utilities.formatDate(
      endDate,
      tz,
      'd'
    );

  const year =
    Utilities.formatDate(
      endDate,
      tz,
      'yyyy'
    );


  if (startMonth === endMonth) {

    return (
      startMonth +
      ' ' +
      startDay +
      ' - ' +
      endDay +
      ', ' +
      year
    );
  }


  return (
    startMonth +
    ' ' +
    startDay +
    ' - ' +
    endMonth +
    ' ' +
    endDay +
    ', ' +
    year
  );
}
function testSeptemberPayrollWeeks() {

  const weeks =
    getPayrollWeeksForMonth(
      '2026-09'
    );

  Logger.log(
    JSON.stringify(
      weeks,
      null,
      2
    )
  );
}
function getMonthlyAttendanceDataForAllStudents_(
  monthText
) {

  if (
    !/^\d{4}-\d{2}$/.test(
      String(monthText || '')
    )
  ) {
    throw new Error(
      'Month must use YYYY-MM format.'
    );
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const tz =
    Session.getScriptTimeZone();

  const profiles =
    getStudentProfiles();

  const results = {};

  // ---------------------------------
  // CREATE STUDENT BUCKETS
  // ---------------------------------

  profiles.forEach(function(profile) {

    const studentId =
      String(
        profile.studentId || ''
      ).trim();

    if (!studentId) {
      return;
    }

    results[studentId] = {
      studentId: studentId,

      studentName:
        profile.fullName ||
        profile.studentName ||
        '',

      studentNumber:
        profile.studentNumber ||
        '',

      last4:
        profile.last4 ||
        profile.last4SSN ||
        '',

      month:
        monthText,

      shifts: [],

      totalHours: 0
    };
  });


  // =================================
  // 1. LOAD ACTIVE CORRECTIONS ONCE
  // =================================

  const correctionsByEntry = {};

  const correctionsSheet =
    ss.getSheetByName(
      'Attendance Corrections'
    );

  if (
    correctionsSheet &&
    correctionsSheet.getLastRow() >= 2
  ) {

    const rows =
      correctionsSheet
        .getDataRange()
        .getValues()
        .slice(1);

    rows.forEach(function(row) {

      const entryId =
        String(row[1] || '').trim();

      const studentId =
        String(row[2] || '').trim();

      const workDate =
        row[4];

      const active =
        row[13] === true ||
        String(row[13] || '')
          .toLowerCase() === 'true';

      if (
        !active ||
        !entryId ||
        !studentId ||
        !(workDate instanceof Date)
      ) {
        return;
      }

      const rowMonth =
        Utilities.formatDate(
          workDate,
          tz,
          'yyyy-MM'
        );

      if (rowMonth !== monthText) {
        return;
      }

      correctionsByEntry[entryId] = {
        studentId: studentId,

        approvedIn:
          row[7],

        approvedOut:
          row[8],

        approvedHours:
          Number(row[9] || 0)
      };
    });
  }


  // =================================
  // 2. TABLET ENTRIES — READ ONCE
  // =================================

  const timeEntriesSheet =
    ss.getSheetByName(
      'Time Entries'
    );

  if (
    timeEntriesSheet &&
    timeEntriesSheet.getLastRow() >= 2
  ) {

    const rows =
      timeEntriesSheet
        .getDataRange()
        .getValues()
        .slice(1);

    rows.forEach(function(row) {

      const entryId =
        String(row[0] || '').trim();

      const studentId =
        String(row[1] || '').trim();

      const clockIn =
        row[3];

      const clockOut =
        row[4];

      const storedHours =
        Number(row[5]);

      const status =
        String(row[7] || '')
          .trim()
          .toUpperCase();

const notes =
  String(row[8] || '').trim();

let activityContext = null;

if (notes) {
  try {
    activityContext = JSON.parse(notes);
  } catch (error) {
    activityContext = null;
  }
}

      if (
        !results[studentId] ||
        !(clockIn instanceof Date) ||
        !(clockOut instanceof Date) ||
        status !== 'COMPLETE'
      ) {
        return;
      }

      const rowMonth =
        Utilities.formatDate(
          clockIn,
          tz,
          'yyyy-MM'
        );

      if (rowMonth !== monthText) {
        return;
      }

if (
  activityContext &&
  activityContext.type === 'ACTIVITY' &&
  (
    activityContext.hoursTreatment === 'VOLUNTEER HOURS' ||
    activityContext.hoursTreatment === 'REPLACES REGULAR SHIFT'
  )
) {
  return;
}

      const correction =
        correctionsByEntry[entryId] ||
        null;

      let effectiveIn =
        clockIn;

      let effectiveOut =
        clockOut;

      let hours =
        storedHours;

      let source =
        'RECORDED';

      if (
        correction &&
        correction.studentId === studentId &&
        correction.approvedIn instanceof Date &&
        correction.approvedOut instanceof Date
      ) {

        effectiveIn =
          correction.approvedIn;

        effectiveOut =
          correction.approvedOut;

        hours =
          Number(
            correction.approvedHours
          );

        source =
          'CORRECTED';
      }

      if (!Number.isFinite(hours)) {
        hours =
          (
            effectiveOut.getTime() -
            effectiveIn.getTime()
          ) / 3600000;
      }

      results[studentId]
        .shifts
        .push({
          entryId: entryId,

          source: source,

          workDate:
            effectiveIn,

          startTime:
            effectiveIn,

          endTime:
            effectiveOut,

          hours:
            Math.round(
              hours * 100
            ) / 100
        });
    });
  }


  // =================================
  // 3. PAYROLL SHIFT ADJUSTMENTS
  // =================================

  const adjustmentSheet =
    ss.getSheetByName(
      'Payroll Shift Adjustments'
    );

  if (
    adjustmentSheet &&
    adjustmentSheet.getLastRow() >= 2
  ) {

    const rows =
      adjustmentSheet
        .getDataRange()
        .getValues()
        .slice(1);

    rows.forEach(function(row) {

      const studentId =
        String(row[1] || '').trim();

      const workDate =
        row[3];

      const startTime =
        row[4];

      const endTime =
        row[5];

      const hours =
        Number(row[6]);

      const active =
        row[10] === true ||
        String(row[10] || '')
          .toLowerCase() === 'true';

      if (
        !active ||
        !results[studentId] ||
        !(workDate instanceof Date)
      ) {
        return;
      }

      const rowMonth =
        Utilities.formatDate(
          workDate,
          tz,
          'yyyy-MM'
        );

      if (rowMonth !== monthText) {
        return;
      }

      results[studentId]
        .shifts
        .push({
          source:
            'ADJUSTMENT',

          workDate:
            workDate,

          startTime:
            startTime,

          endTime:
            endTime,

          hours:
            Number.isFinite(hours)
              ? hours
              : 0
        });
    });
  }


  // =================================
  // 4. ATTENDANCE MANUAL SHIFTS
  // =================================

  const manualSheet =
    ss.getSheetByName(
      'Attendance Manual Shifts'
    );

  if (
    manualSheet &&
    manualSheet.getLastRow() >= 2
  ) {

    const rows =
      manualSheet
        .getDataRange()
        .getValues()
        .slice(1);

    rows.forEach(function(row) {

      const studentId =
        String(row[1] || '').trim();

      const workDate =
        row[3];

      const startTime =
        row[4];

      const endTime =
        row[5];

      const hours =
        Number(row[6]);

      const active =
        row[10] === true ||
        String(row[10] || '')
          .toLowerCase() === 'true';

      if (
        !active ||
        !results[studentId] ||
        !(workDate instanceof Date) ||
        !(startTime instanceof Date) ||
        !(endTime instanceof Date)
      ) {
        return;
      }

      const rowMonth =
        Utilities.formatDate(
          workDate,
          tz,
          'yyyy-MM'
        );

      if (rowMonth !== monthText) {
        return;
      }

      results[studentId]
        .shifts
        .push({
          shiftId:
            String(row[0] || ''),

          source:
            'MANUAL',

          workDate:
            workDate,

          startTime:
            startTime,

          endTime:
            endTime,

          hours:
            Number.isFinite(hours)
              ? hours
              : 0,

          reason:
            String(row[7] || '')
        });
    });
  }

// =================================
// 5. APPROVED ACTIVITY REPLACEMENTS
// =================================

Object.keys(results).forEach(function(studentId) {

  const activityReplacementShifts =
    getActivityReplacementShifts_(
      studentId,
      monthText
    );

  activityReplacementShifts.forEach(function(shift) {

    results[studentId]
      .shifts
      .push({
        source:
          'ACTIVITY_REPLACEMENT',

        workDate:
          shift.replacementDate,

        startTime:
          shift.regularStart,

        endTime:
          shift.regularEnd,

        hours:
          Number(shift.hours || 0),

        activityId:
          shift.activityId,

        activityName:
          shift.activityName
      });
  });
});


  // =================================
  // 6. SORT + TOTAL EACH STUDENT
  // =================================

  Object.keys(results)
    .forEach(function(studentId) {

      const student =
        results[studentId];

      student.shifts.sort(
        function(a, b) {

          const aTime =
            a.startTime instanceof Date
              ? a.startTime.getTime()
              : 0;

          const bTime =
            b.startTime instanceof Date
              ? b.startTime.getTime()
              : 0;

          return aTime - bTime;
        }
      );

      student.totalHours =
        Math.round(
          student.shifts.reduce(
            function(total, shift) {
              return (
                total +
                Number(
                  shift.hours || 0
                )
              );
            },
            0
          ) * 100
        ) / 100;
    });


  return results;
}
function testAllStudentsAttendanceSeptember() {
  const data =
    getMonthlyAttendanceDataForAllStudents_(
      '2026-09'
    );

  Logger.log(
    JSON.stringify(
      data,
      null,
      2
    )
  );
}

function testAllStudentsAttendanceAugust() {
  const data =
    getMonthlyAttendanceDataForAllStudents_(
      '2026-08'
    );

  Logger.log(
    JSON.stringify(
      data,
      null,
      2
    )
  );
}
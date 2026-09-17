function getMissionControlData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const now = new Date();

  const todayKey = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const weekday = Utilities.formatDate(now, tz, 'EEEE');

  const result = {
    totalStudents: 0,
    workingNow: 0,
    scheduledLater: 0,
    missingClockIn: 0,
    pendingRequests: 0,
    todayActivities: [],
    attentionItems: [],
    studentsToday: []
  };

  // -------------------------
  // ACTIVE STUDENTS
  // -------------------------
  const profilesSheet = ss.getSheetByName('Student Profiles');

  const activeStudents = [];

  if (profilesSheet && profilesSheet.getLastRow() > 1) {
    const profiles = profilesSheet.getDataRange().getValues();

    profiles.slice(1).forEach(row => {
      const active =
        row[7] === true ||
        String(row[7]).toLowerCase() === 'true';

      if (!active) return;

      activeStudents.push({
        studentId: String(row[0] || ''),
        name: String(row[1] || '')
      });
    });
  }

  result.totalStudents = activeStudents.length;

  // -------------------------
  // TODAY'S REGULAR SCHEDULES
  // -------------------------
  const scheduleSheet = ss.getSheetByName('Regular Schedules');
  const todaySchedules = {};

  if (scheduleSheet && scheduleSheet.getLastRow() > 1) {
    const schedules = scheduleSheet.getDataRange().getValues();

    schedules.slice(1).forEach(row => {
      const active =
        row[8] === true ||
        String(row[8]).toLowerCase() === 'true';

      if (!active) return;
      if (String(row[4]) !== weekday) return;
const scheduleStartDate =
  row[9];

if (
  scheduleStartDate instanceof Date
) {
  const scheduleStartKey =
    Utilities.formatDate(
      scheduleStartDate,
      tz,
      'yyyy-MM-dd'
    );

  if (
    todayKey < scheduleStartKey
  ) {
    return;
  }
}

const scheduleEndDate =
  row[10];

if (
  scheduleEndDate instanceof Date
) {
  const scheduleEndKey =
    Utilities.formatDate(
      scheduleEndDate,
      tz,
      'yyyy-MM-dd'
    );

  if (
    todayKey > scheduleEndKey
  ) {
    return;
  }
}

      const studentId = String(row[1] || '');
      const studentName = String(row[2] || '');

      const start = row[5];
      const end = row[6];

      if (!(start instanceof Date) || !(end instanceof Date)) {
        return;
      }

      const startToday = new Date(
        now.getFullYear(),
        now.getMonth(),
        now.getDate(),
        start.getHours(),
        start.getMinutes(),
        0
      );

      const endToday = new Date(
        now.getFullYear(),
        now.getMonth(),
        now.getDate(),
        end.getHours(),
        end.getMinutes(),
        0
      );

      todaySchedules[studentId] = {
        studentId,
        studentName,
        start: startToday,
        end: endToday
      };
    });
  }

  // -------------------------
// SCHEDULE EXCEPTIONS TODAY
// -------------------------

const exceptionsSheet =
  ss.getSheetByName(
    'Schedule Exceptions'
  );

if (
  exceptionsSheet &&
  exceptionsSheet.getLastRow() > 1
) {

  const exceptions =
    exceptionsSheet
      .getDataRange()
      .getValues();

  exceptions
    .slice(1)
    .forEach(function(row) {

      const active =
        row[15] === true ||
        String(row[15] || '')
          .toLowerCase() === 'true';

      if (!active) {
        return;
      }

      const studentId =
        String(row[1] || '').trim();

      if (!studentId) {
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

      const exceptionType =
        String(row[4] || '');

      const activityName =
        String(row[6] || 'Activity');


      // ==================================
      // A. ORIGINAL REGULAR WORK DATE
      // ==================================

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
          originalDateKey === todayKey &&
          regularStart instanceof Date &&
          regularEnd instanceof Date
        ) {

          const regularHours =
            (
              regularEnd.getTime() -
              regularStart.getTime()
            ) / 3600000;

          // Entire regular shift was replaced.
          // Student should NOT be expected
          // to clock in on this original date.
          if (
            replacedHours >=
            regularHours - 0.001
          ) {

            delete todaySchedules[
              studentId
            ];
          }

          // Partial replacement:
          // Keep the remaining regular shift
          // expectation for now.
          else if (
            todaySchedules[studentId]
          ) {

            todaySchedules[
              studentId
            ].partialReplacementHours =
              replacedHours;

            todaySchedules[
              studentId
            ].activityName =
              activityName;
          }
        }
      }


      // ==================================
      // B. ACTIVITY / REPLACEMENT DATE
      // ==================================

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
          replacementDateKey === todayKey
        ) {

          todaySchedules[studentId] = {
            studentId:
              studentId,

            studentName:
              String(row[2] || ''),

            start:
              replacementStart,

            end:
              replacementEnd,

            exceptionType:
              exceptionType,

            activityName:
              activityName,

            isActivityReplacement:
              true
          };
        }
      }
    });
}

  // -------------------------
  // CLOCK ENTRIES TODAY
  // -------------------------
  const timeSheet = ss.getSheetByName('Time Entries');
  const todayEntries = {};

  if (timeSheet && timeSheet.getLastRow() > 1) {
    const entries = timeSheet.getDataRange().getValues();

    entries.slice(1).forEach(row => {
      const clockIn = row[3];

      if (!(clockIn instanceof Date)) return;

      const entryKey =
        Utilities.formatDate(clockIn, tz, 'yyyy-MM-dd');

      if (entryKey !== todayKey) return;

      todayEntries[String(row[1] || '')] = {
        studentName: String(row[2] || ''),
        clockIn,
        clockOut:
          row[4] instanceof Date
            ? row[4]
            : null
      };
    });
  }

  // -------------------------
  // BUILD TODAY'S STATUS LIST
  // -------------------------
  const graceMinutes = 5;

  activeStudents.forEach(student => {
    const schedule = todaySchedules[student.studentId];
    const entry = todayEntries[student.studentId];

    let status = 'OFF';
    let detail = 'Not scheduled today';

    if (entry && !entry.clockOut) {
      status = 'WORKING';
      detail =
        'Clocked in at ' +
        Utilities.formatDate(
          entry.clockIn,
          tz,
          'h:mm a'
        );

      result.workingNow++;
    }

    else if (schedule) {
      const graceTime =
        new Date(
          schedule.start.getTime() +
          graceMinutes * 60000
        );

      if (now < schedule.start) {
        status = 'LATER';

        detail =
          'Scheduled ' +
          Utilities.formatDate(
            schedule.start,
            tz,
            'h:mm a'
          );

        result.scheduledLater++;
      }

      else if (!entry && now > graceTime) {
        status = 'MISSING';

        createAttendanceAlert_(
  student.studentId,
  student.name,
  '',
  now,
  schedule.start,
  'MISSING'
);

        detail =
          'Expected at ' +
          Utilities.formatDate(
            schedule.start,
            tz,
            'h:mm a'
          );

        result.missingClockIn++;

        result.attentionItems.push({
          type: 'MISSING',
          title:
            student.name +
            ' has not clocked in',
          detail
        });
      }

      else if (entry && entry.clockOut) {
        status = 'DONE';

        detail =
          'Worked ' +
          Utilities.formatDate(
            entry.clockIn,
            tz,
            'h:mm a'
          ) +
          '–' +
          Utilities.formatDate(
            entry.clockOut,
            tz,
            'h:mm a'
          );
      }

      else {
        status = 'SCHEDULED';

        detail =
          'Scheduled ' +
          Utilities.formatDate(
            schedule.start,
            tz,
            'h:mm a'
          ) +
          '–' +
          Utilities.formatDate(
            schedule.end,
            tz,
            'h:mm a'
          );
      }
    }

    result.studentsToday.push({
      studentId: student.studentId,
      name: student.name,
      status,
      detail
    });
  });

  // -------------------------
  // TODAY'S ACTIVITIES
  // -------------------------
  const activitiesSheet = ss.getSheetByName('Activities');

  if (activitiesSheet && activitiesSheet.getLastRow() > 1) {
    const activities = activitiesSheet.getDataRange().getValues();

    activities.slice(1).forEach(row => {
      const date = row[2];

      if (!(date instanceof Date)) return;

      const activityKey =
        Utilities.formatDate(date, tz, 'yyyy-MM-dd');

      if (activityKey !== todayKey) return;

      result.todayActivities.push({
        activityId: String(row[0] || ''),
        name: String(row[1] || ''),
        time:
          (row[3] instanceof Date
            ? Utilities.formatDate(row[3], tz, 'h:mm a')
            : '') +
          '–' +
          (row[4] instanceof Date
            ? Utilities.formatDate(row[4], tz, 'h:mm a')
            : ''),
        location: String(row[5] || '')
      });
    });
  }

  // -------------------------
  // PENDING SCHEDULE REQUESTS
  // -------------------------
  const requestSheet =
    ss.getSheetByName('Schedule Requests');

  if (requestSheet && requestSheet.getLastRow() > 1) {
    const requests =
      requestSheet.getDataRange().getValues();

    requests.slice(1).forEach(row => {
      if (
        String(row[12] || '').toUpperCase() ===
        'PENDING'
      ) {
        result.pendingRequests++;

        result.attentionItems.push({
          type: 'REQUEST',
          title:
            String(row[2] || '') +
            ' requested a schedule change',
          detail:
            row[5] instanceof Date
              ? Utilities.formatDate(
                  row[5],
                  tz,
                  'MMM d'
                )
              : ''
        });
      }
    });
  }

  result.studentsToday.sort((a, b) =>
    a.name.localeCompare(b.name)
  );

  return result;
}
function getActivitiesForAssignment() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName('Activities');

  if (
    !sheet ||
    sheet.getLastRow() < 2
  ) {
    return [];
  }

  const tz =
    Session.getScriptTimeZone();

  const rows =
    sheet
      .getDataRange()
      .getValues()
      .slice(1);

  const activities = [];

  rows.forEach(function(row) {

    const activityId =
      String(row[0] || '').trim();

    const activityName =
      String(row[1] || '').trim();

    const date =
      row[2];

    const startTime =
      row[3];

    const endTime =
      row[4];

    if (
      !activityId ||
      !activityName ||
      !(date instanceof Date)
    ) {
      return;
    }

    activities.push({
      activityId:
        activityId,

      activityName:
        activityName,

      date:
        Utilities.formatDate(
          date,
          tz,
          'MMM d, yyyy'
        ),

      startTime:
        startTime instanceof Date
          ? Utilities.formatDate(
              startTime,
              tz,
              'h:mm a'
            )
          : '',

      endTime:
        endTime instanceof Date
          ? Utilities.formatDate(
              endTime,
              tz,
              'h:mm a'
            )
          : ''
    });
  });

  activities.sort(function(a, b) {
    return a.date.localeCompare(b.date);
  });

  return activities;
}
function getDashboardPayrollApprovals() {

  const tz =
    Session.getScriptTimeZone();

  const now =
    new Date();

  const monthText =
    Utilities.formatDate(
      now,
      tz,
      'yyyy-MM'
    );

  const payroll =
    getPayrollReview(
      monthText
    );

  const items = [];

  payroll.forEach(function(row) {

    const status =
      String(
        row.status || 'PENDING'
      ).toUpperCase();

    if (
      status === 'NEEDS REVIEW'
    ) {

      items.push({
        type:
          'PAYROLL_REVIEW',

        title:
          row.studentName +
          ' payroll needs review',

        detail:
          Number(
            row.calculatedHours || 0
          ).toFixed(2) +
          ' calculated hours'
      });

      return;
    }

    if (
      status === 'PENDING' &&
      Number(
        row.calculatedHours || 0
      ) > 0
    ) {

      items.push({
        type:
          'PAYROLL_PENDING',

        title:
          row.studentName +
          ' payroll needs approval',

        detail:
          Number(
            row.calculatedHours || 0
          ).toFixed(2) +
          ' hours waiting for approval'
      });
    }

  });

  return {
    month:
      monthText,

    items:
      items
  };
}
function testMissionControl() {
  const start = new Date();

  const data = getMissionControlData();

  Logger.log(
    'Finished in ' +
    ((new Date() - start) / 1000) +
    ' seconds'
  );

  Logger.log(
    JSON.stringify(data, null, 2)
  );
}
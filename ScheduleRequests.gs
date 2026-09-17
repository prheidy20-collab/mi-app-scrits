function setupScheduleRequestsModule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName('Schedule Requests');

  if (!sheet) {
    sheet = ss.insertSheet('Schedule Requests');
  }

  const headers = [
    'Request ID',
    'Student ID',
    'Student Name',
    'Student Email',
    'Request Date',
    'Schedule Date',
    'Current Start',
    'Current End',
    'Requested Start',
    'Requested End',
    'Reason',
    'Comments',
    'Status',
    'Submitted On',
    'Reviewed By',
    'Reviewed On',
    'Admin Notes'
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  sheet.setFrozenRows(1);

  sheet.getRange('E:F').setNumberFormat('m/d/yyyy');
  sheet.getRange('G:J').setNumberFormat('h:mm AM/PM');
  sheet.getRange('N:P').setNumberFormat('m/d/yyyy h:mm AM/PM');

  sheet.autoResizeColumns(1, headers.length);

  return 'Schedule Requests module created.';
}


function submitScheduleRequest(formData) {
  if (!formData) {
    throw new Error('No request information was received.');
  }

  const studentId = String(formData.studentId || '').trim();
  const scheduleDateText = String(formData.scheduleDate || '').trim();
  const requestedStartText =
    String(formData.requestedStart || '').trim();
  const requestedEndText =
    String(formData.requestedEnd || '').trim();
  const reason = String(formData.reason || '').trim();
  const comments = String(formData.comments || '').trim();

  if (
    !studentId ||
    !scheduleDateText ||
    !requestedStartText ||
    !requestedEndText ||
    !reason
  ) {
    throw new Error('Complete all required request fields.');
  }

  const student = getScheduleRequestStudent_(studentId);

  const scheduleDate = parseScheduleRequestDate_(scheduleDateText);

  validateScheduleRequestDate_(scheduleDate);

  const requestedStart =
    combineScheduleRequestDateTime_(
      scheduleDate,
      requestedStartText
    );

  const requestedEnd =
    combineScheduleRequestDateTime_(
      scheduleDate,
      requestedEndText
    );

  if (requestedEnd <= requestedStart) {
    throw new Error(
      'Requested ending time must be later than starting time.'
    );
  }

  const currentSchedule =
    getCurrentScheduleForRequest_(
      studentId,
      scheduleDate
    );

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Schedule Requests');

  if (!sheet) {
    throw new Error(
      'Schedule Requests sheet not found. Run setupScheduleRequestsModule first.'
    );
  }

  const requestId = Utilities.getUuid();
  const now = new Date();

  sheet.appendRow([
    requestId,
    student.studentId,
    student.studentName,
    student.studentEmail,
    now,
    scheduleDate,
    currentSchedule ? currentSchedule.start : '',
    currentSchedule ? currentSchedule.end : '',
    requestedStart,
    requestedEnd,
    reason,
    comments,
    'PENDING',
    now,
    '',
    '',
    ''
  ]);

  return {
    ok: true,
    message: 'Schedule change request submitted for approval.'
  };
}


function getPendingScheduleRequests() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Schedule Requests');

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const tz = Session.getScriptTimeZone();

  return sheet.getDataRange()
    .getValues()
    .slice(1)
    .filter(row =>
      String(row[12]).toUpperCase() === 'PENDING'
    )
    .map(row => ({
      requestId: String(row[0]),
      studentId: String(row[1]),
      studentName: String(row[2]),
      studentEmail: String(row[3] || ''),

      scheduleDate:
        row[5] instanceof Date
          ? Utilities.formatDate(
              row[5],
              tz,
              'MMM d, yyyy'
            )
          : '',

      currentStart:
        row[6] instanceof Date
          ? Utilities.formatDate(
              row[6],
              tz,
              'h:mm a'
            )
          : 'No regular shift',

      currentEnd:
        row[7] instanceof Date
          ? Utilities.formatDate(
              row[7],
              tz,
              'h:mm a'
            )
          : '',

      requestedStart:
        row[8] instanceof Date
          ? Utilities.formatDate(
              row[8],
              tz,
              'h:mm a'
            )
          : '',

      requestedEnd:
        row[9] instanceof Date
          ? Utilities.formatDate(
              row[9],
              tz,
              'h:mm a'
            )
          : '',

      reason: String(row[10] || ''),
      comments: String(row[11] || '')
    }));
}


function approveScheduleRequest(requestId) {
  return reviewScheduleRequest_(
    requestId,
    'APPROVED'
  );
}


function rejectScheduleRequest(requestId) {
  return reviewScheduleRequest_(
    requestId,
    'REJECTED'
  );
}


function reviewScheduleRequest_(requestId, decision) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const requestSheet =
    ss.getSheetByName('Schedule Requests');

  if (!requestSheet) {
    throw new Error('Schedule Requests sheet not found.');
  }

  const rows =
    requestSheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(requestId)) {
      continue;
    }

    if (
      String(rows[i][12]).toUpperCase() !==
      'PENDING'
    ) {
      throw new Error(
        'This request has already been reviewed.'
      );
    }

    const now = new Date();

    if (decision === 'APPROVED') {
      applyApprovedScheduleRequest_(rows[i]);
    }

    requestSheet
      .getRange(i + 1, 13)
      .setValue(decision);

    requestSheet
      .getRange(i + 1, 15)
      .setValue(
        getSettings_().COORDINATOR_EMAIL ||
        'Administrator'
      );

    requestSheet
      .getRange(i + 1, 16)
      .setValue(now);

    return {
      ok: true,
      message:
        'Schedule request ' +
        decision.toLowerCase() +
        '.'
    };
  }

  throw new Error('Schedule request not found.');
}


function applyApprovedScheduleRequest_(requestRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const exceptionsSheet =
    ss.getSheetByName('Schedule Exceptions');

  if (!exceptionsSheet) {
    throw new Error(
      'Schedule Exceptions sheet not found.'
    );
  }

  const studentId = String(requestRow[1]);
  const studentName = String(requestRow[2]);

  const scheduleDate = requestRow[5];

  const currentStart = requestRow[6];
  const currentEnd = requestRow[7];

  const requestedStart = requestRow[8];
  const requestedEnd = requestRow[9];

  const reason = String(requestRow[10] || '');
  const comments = String(requestRow[11] || '');

  const approvedHours =
    Math.round(
      (
        (
          requestedEnd.getTime() -
          requestedStart.getTime()
        ) /
        3600000
      ) * 100
    ) / 100;

  exceptionsSheet.appendRow([
    Utilities.getUuid(),
    studentId,
    studentName,
    scheduleDate,
    'APPROVED SCHEDULE CHANGE',
    '',
    '',
    currentStart,
    currentEnd,
    requestedStart,
    requestedEnd,
    approvedHours,
    getSettings_().COORDINATOR_EMAIL || '',
    new Date(),
    reason +
      (comments ? ' — ' + comments : ''),
    true
  ]);
}


function getScheduleRequestStudent_(studentId) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Student Profiles');

  if (!sheet) {
    throw new Error('Student Profiles sheet not found.');
  }

  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (
      String(rows[i][0]) === String(studentId)
    ) {
      return {
        studentId: String(rows[i][0]),
        studentName: String(rows[i][1]),
        studentEmail: String(rows[i][4] || '')
      };
    }
  }

  throw new Error('Student not found.');
}

function getScheduleShiftsForRequest(
  studentId,
  dateText
) {

  if (!studentId || !dateText) {
    return [];
  }

  const schedules =
    getExpectedSchedulesForDate_(
      String(studentId).trim(),
      String(dateText).trim()
    );

  const tz =
    Session.getScriptTimeZone();

  return schedules.map(function(schedule) {

    return {
      scheduleId:
        schedule.scheduleId || '',

      start:
        Utilities.formatDate(
          schedule.start,
          tz,
          'HH:mm'
        ),

      end:
        Utilities.formatDate(
          schedule.end,
          tz,
          'HH:mm'
        ),

      label:
        Utilities.formatDate(
          schedule.start,
          tz,
          'h:mm a'
        ) +
        ' - ' +
        Utilities.formatDate(
          schedule.end,
          tz,
          'h:mm a'
        )
    };

  });
}

function getCurrentScheduleForRequest_(
  studentId,
  scheduleDate
) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Regular Schedules');

  if (!sheet) return null;

  const weekday =
    Utilities.formatDate(
      scheduleDate,
      Session.getScriptTimeZone(),
      'EEEE'
    );

  const rows =
    sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const active =
      rows[i][8] === true ||
      String(rows[i][8]).toLowerCase() ===
        'true';

    if (
      String(rows[i][1]) ===
        String(studentId) &&
      String(rows[i][4]) === weekday &&
      active
    ) {
      return {
        start: rows[i][5],
        end: rows[i][6]
      };
    }
  }

  return null;
}


function parseScheduleRequestDate_(dateText) {
  const parts =
    String(dateText).split('-').map(Number);

  if (
    parts.length !== 3 ||
    parts.some(Number.isNaN)
  ) {
    throw new Error('Invalid schedule date.');
  }

  return new Date(
    parts[0],
    parts[1] - 1,
    parts[2]
  );
}


function combineScheduleRequestDateTime_(
  date,
  timeText
) {
  const parts =
    String(timeText).split(':').map(Number);

  if (
    parts.length !== 2 ||
    parts.some(Number.isNaN)
  ) {
    throw new Error('Invalid requested time.');
  }

  return new Date(
    date.getFullYear(),
    date.getMonth(),
    date.getDate(),
    parts[0],
    parts[1],
    0
  );
}


function validateScheduleRequestDate_(scheduleDate) {
  const today = new Date();

  const todayOnly =
    new Date(
      today.getFullYear(),
      today.getMonth(),
      today.getDate()
    );

  if (scheduleDate < todayOnly) {
    throw new Error(
      'Past schedules cannot be changed.'
    );
  }

  const maxDate =
    new Date(todayOnly);

  maxDate.setDate(
    maxDate.getDate() + 14
  );

  if (scheduleDate > maxDate) {
    throw new Error(
      'Schedule requests can only be submitted for the current or next two weeks.'
    );
  }
}function approveScheduleRequestWithScope(
  requestId,
  scope,
  throughDateText
) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requestSheet = ss.getSheetByName('Schedule Requests');

  if (!requestSheet) {
    throw new Error('Schedule Requests sheet not found.');
  }

  const rows = requestSheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];

    if (String(row[0]) !== String(requestId)) {
      continue;
    }

    if (String(row[12]).toUpperCase() !== 'PENDING') {
      throw new Error('This request has already been reviewed.');
    }

    const studentId = String(row[1]);
    const studentName = String(row[2]);
    const scheduleDate = row[5];
    const currentStart = row[6];
    const currentEnd = row[7];
    const requestedStart = row[8];
    const requestedEnd = row[9];
    const reason = String(row[10] || '');
    const comments = String(row[11] || '');

    if (!(scheduleDate instanceof Date)) {
      throw new Error('Invalid schedule date.');
    }

    if (!(requestedStart instanceof Date) ||
        !(requestedEnd instanceof Date)) {
      throw new Error('Invalid requested schedule.');
    }

    if (!['DATE_ONLY', 'TEMPORARY', 'PERMANENT'].includes(scope)) {
      throw new Error('Invalid approval scope.');
    }

    if (scope === 'DATE_ONLY') {
      createApprovedScheduleException_({
        studentId,
        studentName,
        scheduleDate,
        currentStart,
        currentEnd,
        requestedStart,
        requestedEnd,
        reason,
        comments,
        scope: 'DATE ONLY'
      });
    }

    if (scope === 'TEMPORARY') {
      if (!throughDateText) {
        throw new Error('Temporary end date is required.');
      }

      const throughDate =
        parseScheduleRequestDate_(throughDateText);

      if (throughDate < scheduleDate) {
        throw new Error(
          'Temporary end date cannot be before the requested schedule date.'
        );
      }

      createTemporaryScheduleExceptions_({
        studentId,
        studentName,
        fromDate: scheduleDate,
        throughDate,
        requestedStart,
        requestedEnd,
        reason,
        comments
      });
    }

    if (scope === 'PERMANENT') {
      updatePermanentRegularSchedule_({
        studentId,
        studentName,
        scheduleDate,
        requestedStart,
        requestedEnd
      });
    }

    requestSheet.getRange(i + 1, 13).setValue('APPROVED');
    requestSheet.getRange(i + 1, 15).setValue(
      getSettings_().COORDINATOR_EMAIL || 'Administrator'
    );
    requestSheet.getRange(i + 1, 16).setValue(new Date());

    requestSheet.getRange(i + 1, 17).setValue(
      'Approved as ' +
      scope +
      (throughDateText ? ' through ' + throughDateText : '')
    );

    return {
      ok: true,
      message: 'Schedule request approved successfully.'
    };
  }

  throw new Error('Schedule request not found.');
}


function createApprovedScheduleException_(data) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Schedule Exceptions');

  if (!sheet) {
    throw new Error('Schedule Exceptions sheet not found.');
  }

  const hours =
    Math.round(
      ((data.requestedEnd - data.requestedStart) / 3600000) * 100
    ) / 100;

  sheet.appendRow([
    Utilities.getUuid(),
    data.studentId,
    data.studentName,
    data.scheduleDate,
    'APPROVED SCHEDULE CHANGE',
    '',
    '',
    data.currentStart || '',
    data.currentEnd || '',
    data.requestedStart,
    data.requestedEnd,
    hours,
    getSettings_().COORDINATOR_EMAIL || '',
    new Date(),
    data.reason +
      (data.comments ? ' — ' + data.comments : '') +
      ' — ' + data.scope,
    true
  ]);
}


function createTemporaryScheduleExceptions_(data) {
  const currentDate = new Date(
    data.fromDate.getFullYear(),
    data.fromDate.getMonth(),
    data.fromDate.getDate()
  );

  const endDate = new Date(
    data.throughDate.getFullYear(),
    data.throughDate.getMonth(),
    data.throughDate.getDate()
  );

  const originalWeekday =
    Utilities.formatDate(
      data.fromDate,
      Session.getScriptTimeZone(),
      'EEEE'
    );

  while (currentDate <= endDate) {
    const weekday =
      Utilities.formatDate(
        currentDate,
        Session.getScriptTimeZone(),
        'EEEE'
      );

    if (weekday === originalWeekday) {
      const requestedStart = new Date(
        currentDate.getFullYear(),
        currentDate.getMonth(),
        currentDate.getDate(),
        data.requestedStart.getHours(),
        data.requestedStart.getMinutes(),
        0
      );

      const requestedEnd = new Date(
        currentDate.getFullYear(),
        currentDate.getMonth(),
        currentDate.getDate(),
        data.requestedEnd.getHours(),
        data.requestedEnd.getMinutes(),
        0
      );

      createApprovedScheduleException_({
        studentId: data.studentId,
        studentName: data.studentName,
        scheduleDate: new Date(currentDate),
        currentStart: '',
        currentEnd: '',
        requestedStart,
        requestedEnd,
        reason: data.reason,
        comments: data.comments,
        scope: 'TEMPORARY SCHEDULE CHANGE'
      });
    }

    currentDate.setDate(currentDate.getDate() + 1);
  }
}


function updatePermanentRegularSchedule_(data) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Regular Schedules');

  if (!sheet) {
    throw new Error('Regular Schedules sheet not found.');
  }

  const weekday =
    Utilities.formatDate(
      data.scheduleDate,
      Session.getScriptTimeZone(),
      'EEEE'
    );

  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const active =
      rows[i][8] === true ||
      String(rows[i][8]).toLowerCase() === 'true';

    if (
      String(rows[i][1]) === String(data.studentId) &&
      String(rows[i][4]) === weekday &&
      active
    ) {
      sheet.getRange(i + 1, 6).setValue(
        timeOnlyFromDate_(data.requestedStart)
      );

      sheet.getRange(i + 1, 7).setValue(
        timeOnlyFromDate_(data.requestedEnd)
      );

      return;
    }
  }

  throw new Error(
    'No active regular schedule was found for that weekday.'
  );
}


function timeOnlyFromDate_(date) {
  return new Date(
    1899,
    11,
    30,
    date.getHours(),
    date.getMinutes(),
    0
  );
}


function checkScheduleRequestConflict(
  studentId,
  scheduleDateText,
  requestedStartText,
  requestedEndText
) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const activitiesSheet =
    ss.getSheetByName('Scheduled Shifts');

  if (!activitiesSheet || activitiesSheet.getLastRow() < 2) {
    return {
      hasConflict: false,
      message: ''
    };
  }

  const date =
    parseDisplayDateForConflict_(scheduleDateText);

  const requestedStart =
    parseDisplayTimeForConflict_(date, requestedStartText);

  const requestedEnd =
    parseDisplayTimeForConflict_(date, requestedEndText);

  const rows =
    activitiesSheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];

    if (String(row[3]) !== String(studentId)) {
      continue;
    }

    const activityDate = row[5];
    const activityStart = row[6];
    const activityEnd = row[7];

    if (!(activityDate instanceof Date) ||
        !(activityStart instanceof Date) ||
        !(activityEnd instanceof Date)) {
      continue;
    }

    const sameDate =
      activityDate.getFullYear() === date.getFullYear() &&
      activityDate.getMonth() === date.getMonth() &&
      activityDate.getDate() === date.getDate();

    if (!sameDate) {
      continue;
    }

    const overlap =
      requestedStart < activityEnd &&
      requestedEnd > activityStart;

    if (overlap) {
      return {
        hasConflict: true,
        message:
          'Student is already assigned to ' +
          String(row[2] || 'an activity') +
          ' from ' +
          Utilities.formatDate(
            activityStart,
            Session.getScriptTimeZone(),
            'h:mm a'
          ) +
          ' to ' +
          Utilities.formatDate(
            activityEnd,
            Session.getScriptTimeZone(),
            'h:mm a'
          ) +
          '.'
      };
    }
  }

  return {
    hasConflict: false,
    message: ''
  };
}


function parseDisplayDateForConflict_(dateText) {
  return Utilities.parseDate(
    String(dateText),
    Session.getScriptTimeZone(),
    'MMM d, yyyy'
  );
}


function parseDisplayTimeForConflict_(date, timeText) {
  const time =
    Utilities.parseDate(
      String(timeText),
      Session.getScriptTimeZone(),
      'h:mm a'
    );

  return new Date(
    date.getFullYear(),
    date.getMonth(),
    date.getDate(),
    time.getHours(),
    time.getMinutes(),
    0
  );
}
function getPendingScheduleChanges() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const tz =
    Session.getScriptTimeZone();

  const changes = [];


  // ==================================
  // 1. STUDENT SCHEDULE REQUESTS
  // ==================================

  const studentRequests =
    getPendingScheduleRequests();

  studentRequests.forEach(function(request) {

    changes.push({
      changeType:
        'STUDENT_REQUEST',

      requestId:
        request.requestId,

      studentId:
        request.studentId,

      studentName:
        request.studentName,

      scheduleDate:
        request.scheduleDate,

      currentStart:
        request.currentStart,

      currentEnd:
        request.currentEnd,

      requestedStart:
        request.requestedStart,

      requestedEnd:
        request.requestedEnd,

      reason:
        request.reason,

      comments:
        request.comments
    });
  });


  // ==================================
  // 2. PENDING ACTIVITY CHANGES
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

    const activityGroups = {};


    rows.forEach(function(row) {

      const exceptionType =
        String(row[4] || '')
          .trim()
          .toUpperCase();

      const approvalStatus =
        String(row[16] || '')
          .trim()
          .toUpperCase();

      if (
        exceptionType !==
          'ACTIVITY PARTIAL REPLACEMENT' ||
        approvalStatus !== 'PENDING'
      ) {
        return;
      }

      const studentId =
        String(row[1] || '').trim();

      const activityId =
        String(row[5] || '').trim();

      const activityName =
        String(row[6] || '').trim();

      const activityStart =
        row[9];

      const activityEnd =
        row[10];

      if (
        !studentId ||
        !activityId
      ) {
        return;
      }

      const groupKey =
        studentId +
        '|' +
        activityId +
        '|' +
        (
          activityStart instanceof Date
            ? activityStart.getTime()
            : ''
        );

      if (!activityGroups[groupKey]) {

        activityGroups[groupKey] = {
          changeType:
            'ACTIVITY_CHANGE',

          groupKey:
            groupKey,

          studentId:
            studentId,

          studentName:
            String(row[2] || ''),

          activityId:
            activityId,

          activityName:
            activityName,

          activityDate:
            activityStart instanceof Date
              ? Utilities.formatDate(
                  activityStart,
                  tz,
                  'MMM d, yyyy'
                )
              : '',

          activityStart:
            activityStart instanceof Date
              ? Utilities.formatDate(
                  activityStart,
                  tz,
                  'h:mm a'
                )
              : '',

          activityEnd:
            activityEnd instanceof Date
              ? Utilities.formatDate(
                  activityEnd,
                  tz,
                  'h:mm a'
                )
              : '',

          totalReplacementHours:
            0,

          allocations: []
        };
      }

      const replacementHours =
        Number(row[11] || 0);

      activityGroups[
        groupKey
      ].totalReplacementHours +=
        replacementHours;

      activityGroups[
        groupKey
      ].allocations.push({
        exceptionId:
          String(row[0] || ''),

        regularDate:
          row[3] instanceof Date
            ? Utilities.formatDate(
                row[3],
                tz,
                'MMM d, yyyy'
              )
            : '',

        regularStart:
          row[7] instanceof Date
            ? Utilities.formatDate(
                row[7],
                tz,
                'h:mm a'
              )
            : '',

        regularEnd:
          row[8] instanceof Date
            ? Utilities.formatDate(
                row[8],
                tz,
                'h:mm a'
              )
            : '',

        replaceHours:
          replacementHours
      });
    });


    Object.keys(
      activityGroups
    ).forEach(function(key) {

      const group =
        activityGroups[key];

      group.totalReplacementHours =
        Math.round(
          group.totalReplacementHours *
          100
        ) / 100;

      changes.push(group);
    });
  }

  // ==================================
  // 3. ACTIVITY PARTICIPATION REQUESTS
  // ==================================

  const participationSheet =
    ss.getSheetByName(
      'Activity Participation Requests'
    );

  if (
    participationSheet &&
    participationSheet.getLastRow() >= 2
  ) {

    const participationRows =
      participationSheet
        .getDataRange()
        .getValues()
        .slice(1);

const activitiesSheet =
  ss.getSheetByName('Activities');

const profilesSheet =
  ss.getSheetByName('Student Profiles');

const activityRows =
  activitiesSheet
    ? activitiesSheet.getDataRange().getValues()
    : [];

const profileRows =
  profilesSheet
    ? profilesSheet.getDataRange().getValues()
    : [];

    participationRows.forEach(function(row) {

            const requestId =
        String(row[0] || '').trim();

      const activityId =
        String(row[1] || '').trim();

      const studentId =
        String(row[2] || '').trim();

      let activityName = '';
      let activityDate = '';
      let activityStart = '';
      let activityEnd = '';
      let studentName = '';

      for (let i = 1; i < activityRows.length; i++) {
        if (
          String(activityRows[i][0] || '').trim() ===
          activityId
        ) {
          activityName =
            String(activityRows[i][1] || '');

          activityDate =
            activityRows[i][2] instanceof Date
              ? Utilities.formatDate(
                  activityRows[i][2],
                  tz,
                  'MMM d, yyyy'
                )
              : '';

          activityStart =
            activityRows[i][3] instanceof Date
              ? Utilities.formatDate(
                  activityRows[i][3],
                  tz,
                  'h:mm a'
                )
              : '';

          activityEnd =
            activityRows[i][4] instanceof Date
              ? Utilities.formatDate(
                  activityRows[i][4],
                  tz,
                  'h:mm a'
                )
              : '';

          break;
        }
      }

      for (let i = 1; i < profileRows.length; i++) {
        if (
          String(profileRows[i][0] || '').trim() ===
          studentId
        ) {
          studentName =
            String(profileRows[i][1] || '');

          break;
        }
      }

      const status =
        String(row[4] || '')
          .trim()
          .toUpperCase();

      if (status !== 'PENDING') {
        return;
      }

     changes.push({
  changeType:
    'ACTIVITY_PARTICIPATION',

  requestId:
    requestId,

  activityId:
    activityId,

  activityName:
    activityName,

  activityDate:
    activityDate,

  activityStart:
    activityStart,

  activityEnd:
    activityEnd,

  studentId:
    studentId,

  studentName:
    studentName,

    hoursTreatment:
  String(row[5] || '').trim(),

replacementDetails:
  String(row[6] || '').trim(),

  requestedOn:
    row[3] instanceof Date
      ? Utilities.formatDate(
          row[3],
          tz,
          'MMM d, yyyy h:mm a'
        )
      : ''
});
    });
  }

  // ==================================
  // 4. MISSING ATTENDANCE REVIEWS
  // ==================================

  const attendanceAlertsSheet =
    ss.getSheetByName('Attendance Alerts');

  if (
    attendanceAlertsSheet &&
    attendanceAlertsSheet.getLastRow() >= 2
  ) {

    const alertRows =
      attendanceAlertsSheet
        .getDataRange()
        .getValues()
        .slice(1);

    alertRows.forEach(function(row) {

      const alertType =
        String(row[6] || '')
          .trim()
          .toUpperCase();

      const status =
        String(row[11] || '')
          .trim()
          .toUpperCase();

      if (
        alertType !== 'MISSING' ||
        status !== 'OPEN'
      ) {
        return;
      }

      const scheduleDateValue =
        row[4] instanceof Date
          ? Utilities.formatDate(
              row[4],
              tz,
              'yyyy-MM-dd'
            )
          : '';

      const expectedSchedule =
        scheduleDateValue
          ? getExpectedScheduleForDate_(
              String(row[1] || ''),
              scheduleDateValue
            )
          : null;

      const scheduledEnd =
        expectedSchedule &&
        expectedSchedule.end instanceof Date
          ? Utilities.formatDate(
              expectedSchedule.end,
              tz,
              'h:mm a'
            )
          : '';

      const scheduledHours =
        expectedSchedule &&
        expectedSchedule.start instanceof Date &&
        expectedSchedule.end instanceof Date
          ? Math.round(
              (
                (
                  expectedSchedule.end.getTime() -
                  expectedSchedule.start.getTime()
                ) / 3600000
              ) * 100
            ) / 100
          : 0;

      changes.push({
        changeType:
          'MISSING_ATTENDANCE',

        alertId:
          String(row[0] || ''),

        studentId:
          String(row[1] || ''),

        studentName:
          String(row[2] || ''),

        scheduleDate:
          row[4] instanceof Date
            ? Utilities.formatDate(
                row[4],
                tz,
                'MMM d, yyyy'
              )
            : '',

     scheduledStart:
  row[5] instanceof Date
    ? Utilities.formatDate(
        row[5],
        tz,
        'h:mm a'
      )
    : '',

scheduledEnd:
  scheduledEnd,

scheduledHours:
  scheduledHours
      });
    });
  }

  return changes;
}
function approvePendingActivityChange(groupKey) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Schedule Exceptions');

  if (!sheet) {
    throw new Error(
      'Schedule Exceptions sheet was not found.'
    );
  }

  const rows =
    sheet
      .getDataRange()
      .getValues();

  let updated = 0;

  for (let i = 1; i < rows.length; i++) {

    const studentId =
      String(rows[i][1] || '').trim();

    const activityId =
      String(rows[i][5] || '').trim();

    const activityStart =
      rows[i][9];

    const approvalStatus =
      String(rows[i][16] || '')
        .trim()
        .toUpperCase();

    if (!(activityStart instanceof Date)) {
      continue;
    }

    const rowGroupKey =
      studentId +
      '|' +
      activityId +
      '|' +
      activityStart.getTime();

    if (
      rowGroupKey !== String(groupKey) ||
      approvalStatus !== 'PENDING'
    ) {
      continue;
    }

    // Active = TRUE
    sheet
      .getRange(i + 1, 16)
      .setValue(true);

    // Approval Status = APPROVED
    sheet
      .getRange(i + 1, 17)
      .setValue('APPROVED');

    updated++;
  }

  if (!updated) {
    throw new Error(
      'Pending activity change was not found.'
    );
  }

  SpreadsheetApp.flush();

  return {
    ok: true,
    message:
      'Activity schedule change approved.'
  };
}


function rejectPendingActivityChange(groupKey) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Schedule Exceptions');

  if (!sheet) {
    throw new Error(
      'Schedule Exceptions sheet was not found.'
    );
  }

  const rows =
    sheet
      .getDataRange()
      .getValues();

  let updated = 0;

  for (let i = 1; i < rows.length; i++) {

    const studentId =
      String(rows[i][1] || '').trim();

    const activityId =
      String(rows[i][5] || '').trim();

    const activityStart =
      rows[i][9];

    const approvalStatus =
      String(rows[i][16] || '')
        .trim()
        .toUpperCase();

    if (!(activityStart instanceof Date)) {
      continue;
    }

    const rowGroupKey =
      studentId +
      '|' +
      activityId +
      '|' +
      activityStart.getTime();

    if (
      rowGroupKey !== String(groupKey) ||
      approvalStatus !== 'PENDING'
    ) {
      continue;
    }

    // Keep Active = FALSE
    sheet
      .getRange(i + 1, 16)
      .setValue(false);

    // Approval Status = REJECTED
    sheet
      .getRange(i + 1, 17)
      .setValue('REJECTED');

    updated++;
  }

  if (!updated) {
    throw new Error(
      'Pending activity change was not found.'
    );
  }

  SpreadsheetApp.flush();

  return {
    ok: true,
    message:
      'Activity schedule change rejected.'
  };
}

function getStudentScheduleViewerData(studentId) {

  if (!studentId) {
    throw new Error(
      'Select a student first.'
    );
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const tz =
    Session.getScriptTimeZone();

  const student =
    getScheduleRequestStudent_(
      studentId
    );

  const result = {
    studentId:
      student.studentId,

    studentName:
      student.studentName,

    regularSchedules: [],

    approvedChanges: [],

    activities: []
  };


  // ==================================
  // 1. REGULAR WEEKLY SCHEDULE
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

      const rowStudentId =
        String(row[1] || '').trim();

      const active =
        row[8] === true ||
        String(row[8] || '')
          .toLowerCase() === 'true';

      if (
        !active ||
        rowStudentId !==
          String(studentId).trim()
      ) {
        return;
      }

      result.regularSchedules.push({
        day:
          String(row[4] || ''),

        start:
          row[5] instanceof Date
            ? Utilities.formatDate(
                row[5],
                tz,
                'h:mm a'
              )
            : '',

        end:
          row[6] instanceof Date
            ? Utilities.formatDate(
                row[6],
                tz,
                'h:mm a'
              )
            : '',

        location:
          String(row[7] || '')
      });
    });
  }


  // ==================================
  // 2. APPROVED SCHEDULE CHANGES
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

    const seenApprovedChanges = {};

    rows.forEach(function(row) {

      const rowStudentId =
        String(row[1] || '').trim();

      const active =
        row[15] === true ||
        String(row[15] || '')
          .toLowerCase() === 'true';

      if (
        !active ||
        rowStudentId !==
          String(studentId).trim()
      ) {
        return;
      }

      const exceptionDate =
        row[3];

      const exceptionType =
        String(row[4] || '');

      const activityName =
        String(row[6] || '');

      const originalStart =
        row[7];

      const originalEnd =
        row[8];

      const replacementStart =
        row[9];

      const replacementEnd =
        row[10];

      const changedHours =
        Number(row[11] || 0);

      const approvalStatus =
        String(row[16] || '')
          .trim()
          .toUpperCase();

      /*
       * Older approved exceptions may not
       * have an Approval Status value.
       * Active=TRUE is still treated as
       * approved for those existing rows.
       */
const approvedKey =
  String(rowStudentId) + '|' +
  String(
    replacementStart instanceof Date
      ? replacementStart.getTime()
      : exceptionDate instanceof Date
        ? exceptionDate.getTime()
        : ''
  ) + '|' +
  String(
    replacementEnd instanceof Date
      ? replacementEnd.getTime()
      : originalEnd instanceof Date
        ? originalEnd.getTime()
        : ''
  ) + '|' +
  String(exceptionType || '') + '|' +
  String(activityName || '');

if (seenApprovedChanges[approvedKey]) {
  return;
}

seenApprovedChanges[approvedKey] = true;

      result.approvedChanges.push({
        exceptionType:
          exceptionType,

        approvalStatus:
          approvalStatus ||
          'APPROVED',

        originalDate:
          exceptionDate instanceof Date
            ? Utilities.formatDate(
                exceptionDate,
                tz,
                'MMM d, yyyy'
              )
            : '',

        originalDateValue:
          exceptionDate instanceof Date
            ? Utilities.formatDate(
                exceptionDate,
                tz,
                'yyyy-MM-dd'
              )
            : '',

        originalStart:
          originalStart instanceof Date
            ? Utilities.formatDate(
                originalStart,
                tz,
                'h:mm a'
              )
            : '',

        originalEnd:
          originalEnd instanceof Date
            ? Utilities.formatDate(
                originalEnd,
                tz,
                'h:mm a'
              )
            : '',

        replacementDate:
          replacementStart instanceof Date
            ? Utilities.formatDate(
                replacementStart,
                tz,
                'MMM d, yyyy'
              )
            : '',

        replacementDateValue:
          replacementStart instanceof Date
            ? Utilities.formatDate(
                replacementStart,
                tz,
                'yyyy-MM-dd'
              )
            : '',

        replacementStart:
          replacementStart instanceof Date
            ? Utilities.formatDate(
                replacementStart,
                tz,
                'h:mm a'
              )
            : '',

        replacementEnd:
          replacementEnd instanceof Date
            ? Utilities.formatDate(
                replacementEnd,
                tz,
                'h:mm a'
              )
            : '',

        changedHours:
          Math.round(
            changedHours * 100
          ) / 100,

        activityName:
          activityName,

        notes:
          String(row[14] || '')
      });
    });
  }


  // ==================================
  // 3. ASSIGNED ACTIVITIES
  // ==================================

  const activitySheet =
    ss.getSheetByName(
      'Scheduled Shifts'
    );

  if (
    activitySheet &&
    activitySheet.getLastRow() >= 2
  ) {

    const rows =
      activitySheet
        .getDataRange()
        .getValues()
        .slice(1);

const seenActivities = {};

    rows.forEach(function(row) {

      const rowStudentId =
        String(row[3] || '').trim();

      if (
        rowStudentId !==
          String(studentId).trim()
      ) {
        return;
      }

      const activityDate =
        row[5];

      const activityStart =
        row[6];

      const activityEnd =
        row[7];

const activityKey =
  String(row[1] || '') +
  '|' +
  (
    activityDate instanceof Date
      ? Utilities.formatDate(
          activityDate,
          tz,
          'yyyy-MM-dd'
        )
      : ''
  ) +
  '|' +
  (
    activityStart instanceof Date
      ? Utilities.formatDate(
          activityStart,
          tz,
          'HH:mm'
        )
      : ''
  ) +
  '|' +
  (
    activityEnd instanceof Date
      ? Utilities.formatDate(
          activityEnd,
          tz,
          'HH:mm'
        )
      : ''
  );

if (seenActivities[activityKey]) {
  return;
}

seenActivities[activityKey] = true;

      result.activities.push({
        shiftId:
          String(row[0] || ''),

        activityId:
          String(row[1] || ''),

        activityName:
          String(row[2] || ''),

        date:
          activityDate instanceof Date
            ? Utilities.formatDate(
                activityDate,
                tz,
                'MMM d, yyyy'
              )
            : '',

        dateValue:
          activityDate instanceof Date
            ? Utilities.formatDate(
                activityDate,
                tz,
                'yyyy-MM-dd'
              )
            : '',

        start:
          activityStart instanceof Date
            ? Utilities.formatDate(
                activityStart,
                tz,
                'h:mm a'
              )
            : '',

        end:
          activityEnd instanceof Date
            ? Utilities.formatDate(
                activityEnd,
                tz,
                'h:mm a'
              )
            : '',

        location:
          String(row[8] || ''),

        status:
          String(row[9] || ''),

        treatment:
          String(row[10] || '')
      });
    });
  }


  // ==================================
  // SORT RESULTS
  // ==================================

  const weekdayOrder = {
    Monday: 1,
    Tuesday: 2,
    Wednesday: 3,
    Thursday: 4,
    Friday: 5,
    Saturday: 6,
    Sunday: 7
  };

  result.regularSchedules.sort(
    function(a, b) {
      return (
        (weekdayOrder[a.day] || 99) -
        (weekdayOrder[b.day] || 99)
      );
    }
  );

  result.approvedChanges.sort(
    function(a, b) {
      return String(
        a.replacementDateValue ||
        a.originalDateValue ||
        ''
      ).localeCompare(
        String(
          b.replacementDateValue ||
          b.originalDateValue ||
          ''
        )
      );
    }
  );

  result.activities.sort(
    function(a, b) {
      return String(
        a.dateValue || ''
      ).localeCompare(
        String(
          b.dateValue || ''
        )
      );
    }
  );


  return result;
}

function approveActivityParticipationRequest(requestId) {

  requestId =
    String(requestId || '').trim();

  if (!requestId) {
    throw new Error(
      'Participation Request ID is required.'
    );
  }


  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  // ==================================
  // PARTICIPATION REQUEST
  // ==================================

  const requestSheet =
    ss.getSheetByName(
      'Activity Participation Requests'
    );

  if (!requestSheet) {
    throw new Error(
      'Activity Participation Requests sheet was not found.'
    );
  }


  const requestRows =
    requestSheet
      .getDataRange()
      .getValues();


  let requestRowNumber = 0;
  let activityId = '';
  let studentId = '';
  let hoursTreatment = '';
  let replacementDetails = '';


  for (
    let i = 1;
    i < requestRows.length;
    i++
  ) {

    const rowRequestId =
      String(
        requestRows[i][0] || ''
      ).trim();

    if (
      rowRequestId !== requestId
    ) {
      continue;
    }


    const status =
      String(
        requestRows[i][4] || ''
      )
        .trim()
        .toUpperCase();


    if (status !== 'PENDING') {
      throw new Error(
        'This participation request is no longer pending.'
      );
    }


    requestRowNumber =
      i + 1;

    activityId =
      String(
        requestRows[i][1] || ''
      ).trim();

    studentId =
      String(
        requestRows[i][2] || ''
      ).trim();

    hoursTreatment =
      String(
        requestRows[i][5] || ''
      )
        .trim()
        .toUpperCase();

    replacementDetails =
      String(
        requestRows[i][6] || ''
      ).trim();

    break;
  }


  if (!requestRowNumber) {
    throw new Error(
      'Pending participation request was not found.'
    );
  }


  // ==================================
  // VALIDATE HOURS TREATMENT
  // ==================================

  const allowedTreatments = [
    'ADDITIONAL HOURS',
    'VOLUNTEER HOURS',
    'REPLACES REGULAR SHIFT'
  ];


  if (
    allowedTreatments.indexOf(
      hoursTreatment
    ) === -1
  ) {
    throw new Error(
      'This request does not contain a valid hours treatment.'
    );
  }


  // ==================================
  // ACTIVITY INFORMATION
  // ==================================

  const activitiesSheet =
    ss.getSheetByName(
      'Activities'
    );

  if (!activitiesSheet) {
    throw new Error(
      'Activities sheet was not found.'
    );
  }


  const activityRows =
    activitiesSheet
      .getDataRange()
      .getValues();


  let activityStart = null;
  let activityEnd = null;


  for (
    let i = 1;
    i < activityRows.length;
    i++
  ) {

    const rowActivityId =
      String(
        activityRows[i][0] || ''
      ).trim();

    if (
      rowActivityId !== activityId
    ) {
      continue;
    }


    activityStart =
      activityRows[i][3];

    activityEnd =
      activityRows[i][4];

    break;
  }


  if (
    !(activityStart instanceof Date) ||
    !(activityEnd instanceof Date)
  ) {
    throw new Error(
      'Activity start/end time could not be found.'
    );
  }


  const activityHours =
    Math.round(
      (
        (
          activityEnd.getTime() -
          activityStart.getTime()
        ) / 3600000
      ) * 100
    ) / 100;


  // ==================================
  // PREVENT DUPLICATE ASSIGNMENT
  // ==================================

  const shiftsSheet =
    ss.getSheetByName(
      'Scheduled Shifts'
    );

  if (!shiftsSheet) {
    throw new Error(
      'Scheduled Shifts sheet was not found.'
    );
  }


  if (
    shiftsSheet.getLastRow() >= 2
  ) {

    const shiftRows =
      shiftsSheet
        .getDataRange()
        .getValues();


    for (
      let i = 1;
      i < shiftRows.length;
      i++
    ) {

      const existingActivityId =
        String(
          shiftRows[i][1] || ''
        ).trim();

      const existingStudentId =
        String(
          shiftRows[i][3] || ''
        ).trim();


      if (
        existingActivityId ===
          activityId &&
        existingStudentId ===
          studentId
      ) {
        throw new Error(
          'This student is already assigned to this activity.'
        );
      }
    }
  }


  // ==================================
  // REPLACEMENT ALLOCATIONS
  // ==================================

  let replacementAllocations = [];


  if (
    hoursTreatment ===
    'REPLACES REGULAR SHIFT'
  ) {

    try {
      replacementAllocations =
        JSON.parse(
          replacementDetails || '[]'
        );
    } catch (error) {
      throw new Error(
        'Replacement schedule information is invalid.'
      );
    }


    if (
      !Array.isArray(
        replacementAllocations
      ) ||
      !replacementAllocations.length
    ) {
      throw new Error(
        'Replacement schedule information is missing.'
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
        'Replacement hours must equal the activity duration.'
      );
    }
  }


  // ==================================
  // CREATE ACTIVITY ASSIGNMENT
  // ==================================

  const assignmentResult =
    saveActivityAssignment({
      studentId:
        studentId,

      activityId:
        activityId,

      scheduleTreatment:
        hoursTreatment,

      replacementAllocations:
        replacementAllocations
    });


  if (
    !assignmentResult ||
    !assignmentResult.ok
  ) {
    throw new Error(
      'The activity assignment could not be created.'
    );
  }


  // ==================================
  // AUTO-APPROVE REPLACEMENT HOURS
  // Coordinator already approved them
  // here, so no second approval.
  // ==================================

  if (
    hoursTreatment ===
    'REPLACES REGULAR SHIFT'
  ) {

    const groupKey =
      studentId +
      '|' +
      activityId +
      '|' +
      activityStart.getTime();


    approvePendingActivityChange(
      groupKey
    );
  }


  // ==================================
  // APPROVE PARTICIPATION REQUEST
  // ==================================

  requestSheet
    .getRange(
      requestRowNumber,
      5
    )
    .setValue(
      'APPROVED'
    );


  SpreadsheetApp.flush();


  return {
    ok: true,

    message:
      'Activity participation approved.'
  };
}

function testJuanitoScheduleViewer() {

  const result =
    getStudentScheduleViewerData(
      'STU-123456'
    );

  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );
}
function rejectActivityParticipationRequest(requestId) {

  requestId =
    String(requestId || '').trim();

  if (!requestId) {
    throw new Error(
      'Participation Request ID is required.'
    );
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      'Activity Participation Requests'
    );

  if (!sheet) {
    throw new Error(
      'Activity Participation Requests sheet was not found.'
    );
  }

  const rows =
    sheet
      .getDataRange()
      .getValues();

  let found = false;

  for (
    let i = 1;
    i < rows.length;
    i++
  ) {

    const rowRequestId =
      String(
        rows[i][0] || ''
      ).trim();

    const status =
      String(
        rows[i][4] || ''
      )
        .trim()
        .toUpperCase();

    if (
      rowRequestId !== requestId ||
      status !== 'PENDING'
    ) {
      continue;
    }

    sheet
      .getRange(
        i + 1,
        5
      )
      .setValue(
        'REJECTED'
      );

    found = true;

    break;
  }

  if (!found) {
    throw new Error(
      'Pending participation request was not found.'
    );
  }

  SpreadsheetApp.flush();

  return {
    ok: true,
    message:
      'Activity participation request rejected.'
  };
}
function getCalendarMonthData(monthText) {

  if (!/^\d{4}-\d{2}$/.test(String(monthText))) {
    throw new Error(
      'Month must use YYYY-MM format.'
    );
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const tz =
    Session.getScriptTimeZone();

  const parts =
    monthText.split('-').map(Number);

  const year =
    parts[0];

  const monthIndex =
    parts[1] - 1;

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

  const result = {
    month:
      monthText,

    monthLabel:
      Utilities.formatDate(
        firstDay,
        tz,
        'MMMM yyyy'
      ),

    days: {}
  };


  // ==================================
  // 1. CREATE ALL DAYS IN MONTH
  // ==================================

  for (
    let day = 1;
    day <= lastDay.getDate();
    day++
  ) {

    const date =
      new Date(
        year,
        monthIndex,
        day
      );

    const dateKey =
      Utilities.formatDate(
        date,
        tz,
        'yyyy-MM-dd'
      );

    result.days[dateKey] = {
      date:
        dateKey,

      displayDate:
        Utilities.formatDate(
          date,
          tz,
          'EEE, MMM d'
        ),

      dayNumber:
        day,

      weekday:
        Utilities.formatDate(
          date,
          tz,
          'EEEE'
        ),

      students: [],

      activities: []
    };
  }


  // ==================================
  // 2. ACTIVE STUDENTS
  // ==================================

  const profilesSheet =
    ss.getSheetByName(
      'Student Profiles'
    );

  const students = {};

  if (
    profilesSheet &&
    profilesSheet.getLastRow() >= 2
  ) {

    const rows =
      profilesSheet
        .getDataRange()
        .getValues()
        .slice(1);

    rows.forEach(function(row) {

      const active =
        row[7] === true ||
        String(row[7] || '')
          .toLowerCase() === 'true';

      if (!active) {
        return;
      }

      const studentId =
        String(row[0] || '').trim();

      if (!studentId) {
        return;
      }

      students[studentId] = {
        studentId:
          studentId,

        studentName:
          String(row[1] || '')
      };
    });
  }


  // ==================================
  // 3. REGULAR SCHEDULES
  // ==================================

  const regularSheet =
    ss.getSheetByName(
      'Regular Schedules'
    );

  const regularSchedules = {};

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

      const studentId =
        String(row[1] || '').trim();

      const active =
        row[8] === true ||
        String(row[8] || '')
          .toLowerCase() === 'true';

      if (
        !active ||
        !students[studentId]
      ) {
        return;
      }

      const weekday =
        String(row[4] || '').trim();

      const start =
        row[5];

      const end =
        row[6];

      if (
        !weekday ||
        !(start instanceof Date) ||
        !(end instanceof Date)
      ) {
        return;
      }

      if (!regularSchedules[weekday]) {
        regularSchedules[weekday] = [];
      }

      regularSchedules[
        weekday
      ].push({
        studentId:
          studentId,

        studentName:
          students[
            studentId
          ].studentName,

        start:
          start,

        end:
          end,

        location:
          String(row[7] || '')
      });
    });
  }


  // ==================================
  // 4. BUILD REGULAR EXPECTED SHIFTS
  // ==================================

  Object.keys(
    result.days
  ).forEach(function(dateKey) {

    const day =
      result.days[dateKey];

    const schedules =
      regularSchedules[
        day.weekday
      ] || [];

    schedules.forEach(
      function(schedule) {

        day.students.push({
          studentId:
            schedule.studentId,

          studentName:
            schedule.studentName,

          start:
            Utilities.formatDate(
              schedule.start,
              tz,
              'h:mm a'
            ),

          end:
            Utilities.formatDate(
              schedule.end,
              tz,
              'h:mm a'
            ),

          location:
            schedule.location,

          source:
            'REGULAR'
        });
      }
    );
  });


  // ==================================
  // 5. APPROVED SCHEDULE EXCEPTIONS
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

const replacementChangeMap = {};

    rows.forEach(function(row) {

      const active =
        row[15] === true ||
        String(row[15] || '')
          .toLowerCase() === 'true';

      if (!active) {
        return;
      }

      const studentId =
        String(row[1] || '').trim();

      if (!students[studentId]) {
        return;
      }

      const originalDate =
        row[3];

      const replacementStart =
        row[9];

      const replacementEnd =
        row[10];

      const exceptionType =
        String(row[4] || '');

      const activityName =
        String(row[6] || '');

      const changedHours =
        Number(row[11] || 0);


      // --------------------------
      // ORIGINAL DATE
      // --------------------------

      if (
        originalDate instanceof Date
      ) {

        const originalKey =
          Utilities.formatDate(
            originalDate,
            tz,
            'yyyy-MM-dd'
          );

        if (
          result.days[originalKey]
        ) {

          result.days[
            originalKey
          ].students.forEach(
            function(item) {

              if (
                item.studentId ===
                studentId
              ) {

                item.hasChange =
                  true;

                item.changedHours =
                  changedHours;

                item.changeType =
                  exceptionType;

                item.activityName =
                  activityName;
              }
            }
          );
        }
      }


      // --------------------------
// REPLACEMENT DATE/TIME
// --------------------------

if (
  replacementStart instanceof Date &&
  replacementEnd instanceof Date
) {

  const replacementKey =
    Utilities.formatDate(
      replacementStart,
      tz,
      'yyyy-MM-dd'
    );

  if (
    result.days[
      replacementKey
    ]
  ) {

    const uniqueChangeKey =
      studentId +
      '|' +
      String(row[5] || '') +
      '|' +
      replacementStart.getTime() +
      '|' +
      replacementEnd.getTime();

    if (
      replacementChangeMap[
        uniqueChangeKey
      ]
    ) {

      const maxReplacementHours =
  Math.round(
    (
      (
        replacementEnd.getTime() -
        replacementStart.getTime()
      ) / 3600000
    ) * 100
  ) / 100;

replacementChangeMap[
  uniqueChangeKey
].changedHours +=
  changedHours;

replacementChangeMap[
  uniqueChangeKey
].changedHours =
  Math.min(
    maxReplacementHours,
    Math.round(
      replacementChangeMap[
        uniqueChangeKey
      ].changedHours * 100
    ) / 100
  );

    } else {

      const changeItem = {
        studentId:
          studentId,

        studentName:
          students[
            studentId
          ].studentName,

        start:
          Utilities.formatDate(
            replacementStart,
            tz,
            'h:mm a'
          ),

        end:
          Utilities.formatDate(
            replacementEnd,
            tz,
            'h:mm a'
          ),

        location:
          '',

        source:
          'CHANGE',

        changeType:
          exceptionType,

        activityName:
          activityName,

        changedHours:
  Math.min(
    changedHours,
    Math.round(
      (
        (
          replacementEnd.getTime() -
          replacementStart.getTime()
        ) / 3600000
      ) * 100
    ) / 100
  )
      };

      result.days[
        replacementKey
      ].students.push(
        changeItem
      );

      replacementChangeMap[
        uniqueChangeKey
      ] = changeItem;
    }
   }
}

    }); // closes rows.forEach(function(row) {

  } // closes if (exceptionsSheet && ...)

  // ==================================
  // 6. ASSIGNED ACTIVITIES
  // ==================================

  const activitySheet =
    ss.getSheetByName(
      'Scheduled Shifts'
    );

  const seenActivityAssignments = {};

  if (
    activitySheet &&
    activitySheet.getLastRow() >= 2
  ) {

    const rows =
      activitySheet
        .getDataRange()
        .getValues()
        .slice(1);

    rows.forEach(function(row) {

      const activityId =
        String(row[1] || '').trim();

      const activityName =
        String(row[2] || '').trim();

      const studentId =
        String(row[3] || '').trim();

      const activityDate =
        row[5];

      const activityStart =
        row[6];

      const activityEnd =
        row[7];

      if (
        !activityId ||
        !studentId ||
        !(activityDate instanceof Date)
      ) {
        return;
      }

      const dateKey =
        Utilities.formatDate(
          activityDate,
          tz,
          'yyyy-MM-dd'
        );

      if (
        !result.days[dateKey]
      ) {
        return;
      }

      const activityKey =
        activityId +
        '|' +
        studentId +
        '|' +
        dateKey;

      if (
        seenActivityAssignments[
          activityKey
        ]
      ) {
        return;
      }

      seenActivityAssignments[
        activityKey
      ] = true;

      result.days[
        dateKey
      ].activities.push({
        activityId:
          activityId,

        activityName:
          activityName,

        studentId:
          studentId,

        studentName:
          students[studentId]
            ? students[
                studentId
              ].studentName
            : String(row[4] || ''),

        start:
          activityStart instanceof Date
            ? Utilities.formatDate(
                activityStart,
                tz,
                'h:mm a'
              )
            : '',

        end:
          activityEnd instanceof Date
            ? Utilities.formatDate(
                activityEnd,
                tz,
                'h:mm a'
              )
            : '',

        location:
          String(row[8] || '')
      });
    });
  }
  return result;
}

function testCalendarAugust() {

  const result =
    getCalendarMonthData(
      '2026-08'
    );

  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );
}


function exportFiestaHubProject() {
  const scriptId = ScriptApp.getScriptId();

  const response = UrlFetchApp.fetch(
    'https://script.googleapis.com/v1/projects/' +
      scriptId +
      '/content',
    {
      method: 'get',
      headers: {
        Authorization:
          'Bearer ' + ScriptApp.getOAuthToken()
      },
      muteHttpExceptions: true
    }
  );

  Logger.log(response.getResponseCode());
  Logger.log(response.getContentText());
}
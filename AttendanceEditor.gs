function getAttendanceAlertsForStudentMonth_(
  studentId,
  monthText
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      'Attendance Alerts'
    );

  if (
    !sheet ||
    sheet.getLastRow() < 2
  ) {
    return [];
  }

  const tz =
    Session.getScriptTimeZone();

  const cleanStudentId =
    String(studentId || '').trim();

  const rows =
    sheet
      .getDataRange()
      .getValues()
      .slice(1);

  const alerts = [];

  rows.forEach(function(row) {

    const rowStudentId =
      String(row[1] || '').trim();

    const scheduledDate =
      row[4];

    if (
      rowStudentId !== cleanStudentId ||
      !(scheduledDate instanceof Date)
    ) {
      return;
    }

    const rowMonth =
      Utilities.formatDate(
        scheduledDate,
        tz,
        'yyyy-MM'
      );

    if (rowMonth !== monthText) {
      return;
    }

    alerts.push({
      alertId:
        String(row[0] || ''),

      studentId:
        rowStudentId,

      studentName:
        String(row[2] || ''),

      scheduledDate:
        scheduledDate,

      scheduledStart:
        row[5],

      alertType:
        String(row[6] || ''),

      alertTime:
        row[7],

      status:
        String(row[11] || ''),

      excusedReason:
        String(row[12] || ''),

      excusedBy:
        String(row[13] || ''),

      excusedOn:
        row[14]
    });
  });

  return alerts;
}


function getAttendanceEditorData(studentId, monthText) {

  if (!studentId) {
    throw new Error('Student ID is required.');
  }

  if (!/^\d{4}-\d{2}$/.test(String(monthText || ''))) {
    throw new Error('Month must use YYYY-MM format.');
  }

  const attendance =
    getStudentMonthlyAttendanceData_(
      studentId,
      monthText
    );

  const corrections =
    getAttendanceCorrections_(
      studentId,
      monthText
    );

  const attendanceAlerts =
    getAttendanceAlertsForStudentMonth_(
      studentId,
      monthText
    );

  const approvals =
    getExistingPayrollApprovals_(
      monthText
    );

  const payrollStudent =
    approvals[
      String(studentId).trim()
    ] || null;

  const tz =
    Session.getScriptTimeZone();


  // Apply corrections without changing Time Entries
  const correctedShifts =
    (attendance.shifts || []).map(function(shift) {

      const entryId =
        String(shift.entryId || '');

      const correction =
        entryId
          ? corrections[entryId]
          : null;


      // Tablet entry WITH correction
      if (
        shift.source === 'RECORDED' &&
        correction
      ) {

        return {
          entryId: entryId,

          shiftId:
            String(shift.shiftId || ''),

          source: 'RECORDED',

          corrected: true,

          correctionReason:
            correction.reason || '',

          workDate:
            shift.workDate,

          originalStartTime:
            shift.startTime,

          originalEndTime:
            shift.endTime,

          startTime:
            correction.approvedIn,

          endTime:
            correction.approvedOut,

          hours:
            Number(
              correction.approvedHours || 0
            )
        };
      }


      // Normal entry / manual detailed shift
      return {
        entryId: entryId,

        shiftId:
          String(shift.shiftId || ''),

        source:
          String(shift.source || ''),

        corrected: false,

        correctionReason: '',

        workDate:
          shift.workDate,

        originalStartTime: null,

        originalEndTime: null,

        startTime:
          shift.startTime,

        endTime:
          shift.endTime,

        hours:
          Number(shift.hours || 0)
      };
    });


  // Recalculate detailed hours using corrected times
  const detailedHours =
    Math.round(
      correctedShifts.reduce(
        function(total, shift) {
          return total +
            Number(shift.hours || 0);
        },
        0
      ) * 100
    ) / 100;


  const approvedHours =
    payrollStudent
      ? Number(
          payrollStudent.approvedHours || 0
        )
      : 0;


  const difference =
    Math.round(
      (
        approvedHours -
        detailedHours
      ) * 100
    ) / 100;


  const unassignedHours =
    Math.max(
      0,
      difference
    );


  // Convert Dates before sending to HTML
  const safeShifts =
    correctedShifts.map(function(shift) {
const matchingAttendanceAlerts =
  attendanceAlerts.filter(
    function(alert) {

      const alertType =
        String(alert.alertType || '')
          .toUpperCase();

      const allowedTypes = [
        'LATE',
        'MISSING',
        'EARLY CLOCK OUT'
      ];

      if (!allowedTypes.includes(alertType)) {
        return false;
      }

     if (
  ['EXCUSED', 'WORKED'].includes(
    String(alert.status || '')
      .toUpperCase()
  )
) {
  return false;
}

      if (
        !(alert.scheduledDate instanceof Date) ||
        !(shift.workDate instanceof Date)
      ) {
        return false;
      }

      return (
        Utilities.formatDate(
          alert.scheduledDate,
          tz,
          'yyyy-MM-dd'
        ) ===
        Utilities.formatDate(
          shift.workDate,
          tz,
          'yyyy-MM-dd'
        )
      );
    }
  );
      return {
        entryId:
          String(shift.entryId || ''),

        shiftId:
          String(shift.shiftId || ''),

        source:
          String(shift.source || ''),

        corrected:
          Boolean(shift.corrected),

        correctionReason:
          String(
            shift.correctionReason || ''
          ),
attendanceAlerts:
  matchingAttendanceAlerts.map(
    function(alert) {
      return {
        alertId:
          String(alert.alertId || ''),

        alertType:
          String(alert.alertType || ''),

        status:
          String(alert.status || '')
      };
    }
  ),

hasAttendanceAlert:
  matchingAttendanceAlerts.length > 0,
        workDate:
          shift.workDate instanceof Date
            ? Utilities.formatDate(
                shift.workDate,
                tz,
                'MM/dd/yyyy'
              )
            : '',

        startTime:
          shift.startTime instanceof Date
            ? Utilities.formatDate(
                shift.startTime,
                tz,
                'h:mm a'
              )
            : '',

        endTime:
          shift.endTime instanceof Date
            ? Utilities.formatDate(
                shift.endTime,
                tz,
                'h:mm a'
              )
            : '',

        originalStartTime:
          shift.originalStartTime instanceof Date
            ? Utilities.formatDate(
                shift.originalStartTime,
                tz,
                'h:mm a'
              )
            : '',

        originalEndTime:
          shift.originalEndTime instanceof Date
            ? Utilities.formatDate(
                shift.originalEndTime,
                tz,
                'h:mm a'
              )
            : '',

        hours:
          Number(shift.hours || 0)
      };
    });

const unmatchedAttendanceAlerts =
  attendanceAlerts
    .filter(function(alert) {

      const alertType =
        String(alert.alertType || '')
          .toUpperCase();

      if (alertType !== 'MISSING') {
        return false;
      }

      if (
  ['EXCUSED', 'WORKED'].includes(
    String(alert.status || '')
      .toUpperCase()
  )
) {
  return false;
}

      const hasMatchingShift =
        correctedShifts.some(
          function(shift) {

            if (
              !(alert.scheduledDate instanceof Date) ||
              !(shift.workDate instanceof Date)
            ) {
              return false;
            }

            return (
              Utilities.formatDate(
                alert.scheduledDate,
                tz,
                'yyyy-MM-dd'
              ) ===
              Utilities.formatDate(
                shift.workDate,
                tz,
                'yyyy-MM-dd'
              )
            );
          }
        );

      return !hasMatchingShift;
    })
    .map(function(alert) {

      return {
        alertId:
          String(alert.alertId || ''),

        alertType:
          String(alert.alertType || ''),

        scheduledDate:
          alert.scheduledDate instanceof Date
            ? Utilities.formatDate(
                alert.scheduledDate,
                tz,
                'MM/dd/yyyy'
              )
            : '',

        scheduledStart:
          alert.scheduledStart instanceof Date
            ? Utilities.formatDate(
                alert.scheduledStart,
                tz,
                'h:mm a'
              )
            : ''
      };
    });
  return {
    studentId:
      String(attendance.studentId || ''),

    studentName:
      String(attendance.studentName || ''),

    studentNumber:
      String(attendance.studentNumber || ''),

    last4:
      String(attendance.last4 || ''),

    month:
      String(attendance.month || ''),

    shifts:
      safeShifts,

unmatchedAttendanceAlerts:
  unmatchedAttendanceAlerts,

    detailedHours:
      detailedHours,

    approvedHours:
      Math.round(
        approvedHours * 100
      ) / 100,

    unassignedHours:
      unassignedHours,

    readyForAttendanceSheet:
      Math.abs(difference) < 0.01
  };
}

function excuseAttendanceAlert(
  alertId,
  reason
) {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      'Attendance Alerts'
    );

  if (!sheet) {
    throw new Error(
      'Attendance Alerts sheet was not found.'
    );
  }

  const cleanAlertId =
    String(alertId || '').trim();

  const cleanReason =
    String(reason || '').trim();

  if (!cleanAlertId) {
    throw new Error(
      'Alert ID is required.'
    );
  }

  if (!cleanReason) {
    throw new Error(
      'A reason is required.'
    );
  }

  const rows =
    sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {

    const rowAlertId =
      String(rows[i][0] || '').trim();

    if (rowAlertId !== cleanAlertId) {
      continue;
    }

    const alertType =
      String(rows[i][6] || '')
        .trim()
        .toUpperCase();

   const allowedTypes = [
  'LATE',
  'MISSING',
  'EARLY CLOCK OUT'
];

if (!allowedTypes.includes(alertType)) {
  throw new Error(
    'This attendance alert cannot be excused.'
  );
}

    sheet
      .getRange(i + 1, 12)
      .setValue('EXCUSED');

    sheet
  .getRange(i + 1, 13)
  .setValue(cleanReason);

sheet
  .getRange(i + 1, 14)
  .setValue(
    Session.getActiveUser().getEmail() ||
    'Coordinator'
  );

sheet
  .getRange(i + 1, 15)
  .setValue(new Date());

    SpreadsheetApp.flush();

    return {
      ok: true,
      message:
  'Attendance alert was excused successfully.'
    };
  }

  throw new Error(
    'Attendance alert was not found.'
  );
}
function markMissingAttendanceNotExcused(
  alertId
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      'Attendance Alerts'
    );

  if (!sheet) {
    throw new Error(
      'Attendance Alerts sheet was not found.'
    );
  }

  const cleanAlertId =
    String(alertId || '').trim();

  if (!cleanAlertId) {
    throw new Error(
      'Alert ID is required.'
    );
  }

  const rows =
    sheet.getDataRange().getValues();

  for (
    let i = 1;
    i < rows.length;
    i++
  ) {

    const rowAlertId =
      String(rows[i][0] || '').trim();

    if (
      rowAlertId !== cleanAlertId
    ) {
      continue;
    }

    const alertType =
      String(rows[i][6] || '')
        .trim()
        .toUpperCase();

    const status =
      String(rows[i][11] || '')
        .trim()
        .toUpperCase();

    if (
      alertType !== 'MISSING'
    ) {
      throw new Error(
        'Only MISSING attendance alerts can be marked as not excused.'
      );
    }

    if (
      status !== 'OPEN'
    ) {
      throw new Error(
        'This attendance alert has already been reviewed.'
      );
    }

    const coordinator =
      Session
        .getActiveUser()
        .getEmail() ||
      'Coordinator';

    sheet
      .getRange(
        i + 1,
        12
      )
      .setValue(
        'UNEXCUSED'
      );

    sheet
      .getRange(
        i + 1,
        13
      )
      .setValue(
        'Reviewed - absence not excused.'
      );

    sheet
      .getRange(
        i + 1,
        14
      )
      .setValue(
        coordinator
      );

    sheet
      .getRange(
        i + 1,
        15
      )
      .setValue(
        new Date()
      );

    SpreadsheetApp.flush();

    return {
      ok: true,
      message:
        'Attendance absence was marked as not excused.'
    };
  }

  throw new Error(
    'Attendance alert was not found.'
  );
}
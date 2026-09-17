function setupPayrollModule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  createOrResetHeaders_(ss, 'Payroll Adjustments', [
    'Adjustment ID',
    'Student ID',
    'Student Name',
    'Work Date',
    'Adjustment Type',
    'Hours',
    'Reason',
    'Created By',
    'Created On',
    'Active'
  ]);

  createOrResetHeaders_(ss, 'Payroll Approvals', [
    'Approval ID',
    'Student ID',
    'Student Name',
    'Payroll Month',
    'Recorded Hours',
    'Adjustment Hours',
    'Approved Hours',
    'Hourly Rate',
    'Approved Amount',
    'Status',
    'Approved By',
    'Approved On',
    'Notes'
  ]);

  const adjustments =
    ss.getSheetByName('Payroll Adjustments');

  const approvals =
    ss.getSheetByName('Payroll Approvals');

  adjustments.setFrozenRows(1);
  approvals.setFrozenRows(1);

  adjustments
    .getRange('D:D')
    .setNumberFormat('m/d/yyyy');

  adjustments
    .getRange('F:F')
    .setNumberFormat('0.00');

  adjustments
    .getRange('I:I')
    .setNumberFormat('m/d/yyyy h:mm AM/PM');

  approvals
    .getRange('D:D')
    .setNumberFormat('yyyy-mm');

  approvals
    .getRange('E:G')
    .setNumberFormat('0.00');

  approvals
    .getRange('H:I')
    .setNumberFormat('$0.00');

  approvals
    .getRange('L:L')
    .setNumberFormat('m/d/yyyy h:mm AM/PM');

  adjustments.autoResizeColumns(
    1,
    adjustments.getLastColumn()
  );

  approvals.autoResizeColumns(
    1,
    approvals.getLastColumn()
  );

  return 'Payroll module created successfully.';
}

function getPayrollReview(monthText) {
  if (!/^\d{4}-\d{2}$/.test(String(monthText))) {
    throw new Error('Month must use YYYY-MM format.');
  }

  const profiles =
    getStudentProfiles();

  const approvals =
    getExistingPayrollApprovals_(monthText);
const attendanceByStudent =
  getMonthlyAttendanceDataForAllStudents_(
    monthText
  );
  return profiles.map(profile => {
    const id = profile.studentId;

    const attendance =
  attendanceByStudent[id] || {
    studentId: id,
    shifts: [],
    totalHours: 0
  };

    let recordedHours = 0;
    let adjustmentHours = 0;

    attendance.shifts.forEach(function(shift) {
      const hours =
        Number(shift.hours || 0);

      if (
        shift.source === 'RECORDED' ||
        shift.source === 'CORRECTED'
      ) {
        recordedHours += hours;
      } else {
        adjustmentHours += hours;
      }
    });

    recordedHours =
      Math.round(
        recordedHours * 100
      ) / 100;

    adjustmentHours =
      Math.round(
        adjustmentHours * 100
      ) / 100;

    const calculatedHours =
      Math.round(
        attendance.totalHours * 100
      ) / 100;

    const existing =
      approvals[id] || null;

    let payrollStatus =
      existing
        ? existing.status
        : 'PENDING';

    if (existing) {
      const recordedChanged =
        Math.abs(
          Number(existing.recordedHours || 0) -
          recordedHours
        ) > 0.001;

      const adjustmentsChanged =
        Math.abs(
          Number(existing.adjustmentHours || 0) -
          adjustmentHours
        ) > 0.001;

      if (
        payrollStatus === 'APPROVED' &&
        (
          recordedChanged ||
          adjustmentsChanged
        )
      ) {
        payrollStatus =
          'NEEDS REVIEW';

        const approvalSheet =
          SpreadsheetApp
            .getActiveSpreadsheet()
            .getSheetByName(
              'Payroll Approvals'
            );

        approvalSheet
          .getRange(
            existing.rowNumber,
            10
          )
          .setValue(
            'NEEDS REVIEW'
          );

        approvalSheet
          .getRange(
            existing.rowNumber,
            13
          )
          .setValue(
            'Recorded or adjustment hours changed after approval.'
          );
      }
    }

    return {
      studentId: id,
      studentName: profile.fullName,
      studentNumber: profile.studentNumber,
      last4Masked: profile.last4Masked,

      recordedHours:
        recordedHours,

      adjustmentHours:
        adjustmentHours,

      calculatedHours:
        calculatedHours,

     approvedHours:
  payrollStatus === 'APPROVED' &&
  existing
    ? Number(
        existing.approvedHours
      )
    : calculatedHours,
      hourlyRate:
        10.50,

      amount:
  payrollStatus === 'APPROVED' &&
  existing
    ? Number(existing.amount)
    : Math.round(
        calculatedHours *
        10.50 *
        100
      ) / 100,

      status:
        payrollStatus
    };
  });
}

function addPayrollAdjustment(formData) {
  if (!formData) {
    throw new Error('No adjustment information was received.');
  }

  const studentId =
    String(formData.studentId || '').trim();

  const workDateText =
    String(formData.workDate || '').trim();

  const adjustmentType =
    String(formData.adjustmentType || '').trim();

  const hours =
    Number(formData.hours);

  const reason =
    String(formData.reason || '').trim();

  if (
    !studentId ||
    !workDateText ||
    !adjustmentType ||
    !Number.isFinite(hours) ||
    !reason
  ) {
    throw new Error(
      'Complete all required adjustment fields.'
    );
  }

  if (hours === 0) {
    throw new Error(
      'Adjustment hours cannot be zero.'
    );
  }

  const student =
    getScheduleRequestStudent_(studentId);

  const workDate =
    parseScheduleRequestDate_(workDateText);

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Payroll Adjustments');

  if (!sheet) {
    throw new Error(
      'Payroll Adjustments sheet not found. Run setupPayrollModule first.'
    );
  }

  sheet.appendRow([
    Utilities.getUuid(),
    student.studentId,
    student.studentName,
    workDate,
    adjustmentType,
    hours,
    reason,
    getSettings_().COORDINATOR_EMAIL ||
      'Administrator',
    new Date(),
    true
  ]);

  

  return {
    ok: true,
    message:
      'Payroll adjustment added for ' +
      student.studentName +
      '.'
  };
}


function approveStudentPayroll(
  monthText,
  studentId,
  approvedHours,
  notes
) {
  if (
    !/^\d{4}-\d{2}$/.test(
      String(monthText)
    )
  ) {
    throw new Error(
      'Invalid payroll month.'
    );
  }

  const hours =
    Number(approvedHours);

  if (
    !Number.isFinite(hours) ||
    hours < 0
  ) {
    throw new Error(
      'Approved hours must be zero or greater.'
    );
  }

  const student =
    getStudentProfileDetail(studentId);

 const attendance =
  getStudentMonthlyAttendanceData_(
    studentId,
    monthText
  );

let recordedHours = 0;
let adjustmentHours = 0;

attendance.shifts.forEach(function(shift) {

  const shiftHours =
    Number(shift.hours || 0);

  if (
    shift.source === 'RECORDED' ||
    shift.source === 'CORRECTED'
  ) {
    recordedHours += shiftHours;
  } else {
    adjustmentHours += shiftHours;
  }

});

recordedHours =
  Math.round(
    recordedHours * 100
  ) / 100;

adjustmentHours =
  Math.round(
    adjustmentHours * 100
  ) / 100;

  const amount =
    Math.round(
      hours * 10.50 * 100
    ) / 100;

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        'Payroll Approvals'
      );

  if (!sheet) {
    throw new Error(
      'Payroll Approvals sheet not found.'
    );
  }

  const rows =
    sheet.getDataRange().getValues();

  const monthDate =
    new Date(
      Number(
        monthText.substring(0, 4)
      ),
      Number(
        monthText.substring(5, 7)
      ) - 1,
      1
    );

  const tz =
    Session.getScriptTimeZone();

  for (
    let i = 1;
    i < rows.length;
    i++
  ) {
    const rowMonthValue =
      rows[i][3];

    let rowMonth = '';

    if (
      rowMonthValue instanceof Date
    ) {
      rowMonth =
        Utilities.formatDate(
          rowMonthValue,
          tz,
          'yyyy-MM'
        );
    } else {
      rowMonth =
        String(
          rowMonthValue || ''
        ).substring(0, 7);
    }

    if (
      String(
        rows[i][1] || ''
      ).trim() ===
        String(studentId).trim() &&
      rowMonth === monthText
    ) {

      sheet
        .getRange(
          i + 1,
          5,
          1,
          9
        )
        .setValues([[
          recordedHours,
          adjustmentHours,
          hours,
          10.50,
          amount,
          'APPROVED',
          getSettings_()
            .COORDINATOR_EMAIL ||
            'Administrator',
          new Date(),
          String(notes || '')
        ]]);

      SpreadsheetApp.flush();

      return {
        ok: true,
        message:
          student.fullName +
          ' payroll approved.',
        approvedHours: hours,
        amount: amount
      };
    }
  }

  sheet.appendRow([
    Utilities.getUuid(),
    studentId,
    student.fullName,
    monthDate,
    recordedHours,
    adjustmentHours,
    hours,
    10.50,
    amount,
    'APPROVED',
    getSettings_()
      .COORDINATOR_EMAIL ||
      'Administrator',
    new Date(),
    String(notes || '')
  ]);

  SpreadsheetApp.flush();

  return {
    ok: true,
    message:
      student.fullName +
      ' payroll approved.',
    approvedHours: hours,
    amount: amount
  };
}



function getRecordedHoursForPayroll_(monthText) {
  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Time Entries');

  const totals = {};

  if (!sheet || sheet.getLastRow() < 2) {
    return totals;
  }

  const rows =
    sheet.getDataRange().getValues();

  rows.slice(1).forEach(row => {
    const studentId =
      String(row[1] || '');

    const clockIn =
      row[3];

    const totalHours =
      Number(row[5]);

    if (
      !studentId ||
      !(clockIn instanceof Date) ||
      !Number.isFinite(totalHours)
    ) {
      return;
    }

    const entryMonth =
      Utilities.formatDate(
        clockIn,
        Session.getScriptTimeZone(),
        'yyyy-MM'
      );

    if (entryMonth !== monthText) {
      return;
    }

    totals[studentId] =
      (totals[studentId] || 0) +
      totalHours;
  });

  return totals;
}


function getPayrollAdjustmentsByStudent_(monthText) {
  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Payroll Adjustments');

  const totals = {};

  if (!sheet || sheet.getLastRow() < 2) {
    return totals;
  }

  const rows =
    sheet.getDataRange().getValues();

  rows.slice(1).forEach(row => {
    const studentId =
      String(row[1] || '');

    const workDate =
      row[3];

    const hours =
      Number(row[5]);

    const active =
      row[9] === true ||
      String(row[9]).toLowerCase() ===
        'true';

    if (
      !active ||
      !studentId ||
      !(workDate instanceof Date) ||
      !Number.isFinite(hours)
    ) {
      return;
    }

    const adjustmentMonth =
      Utilities.formatDate(
        workDate,
        Session.getScriptTimeZone(),
        'yyyy-MM'
      );

    if (adjustmentMonth !== monthText) {
      return;
    }

    totals[studentId] =
      (totals[studentId] || 0) +
      hours;
  });

  return totals;
}


function getExistingPayrollApprovals_(monthText) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Payroll Approvals');

  const approvals = {};

  if (!sheet || sheet.getLastRow() < 2) {
    return approvals;
  }

  const rows = sheet
    .getDataRange()
    .getValues();

  const tz = Session.getScriptTimeZone();

  rows.slice(1).forEach((row, rowIndex) => {
    const payrollMonth = row[3];

    let rowMonth = '';

    if (payrollMonth instanceof Date) {
      rowMonth = Utilities.formatDate(
        payrollMonth,
        tz,
        'yyyy-MM'
      );
    } else {
      rowMonth = String(
        payrollMonth || ''
      ).substring(0, 7);
    }

    if (rowMonth !== monthText) {
      return;
    }

    const studentId =
      String(row[1] || '').trim();

    if (!studentId) {
      return;
    }

    approvals[studentId] = {
      recordedHours:
        Number(row[4] || 0),

      adjustmentHours:
        Number(row[5] || 0),

      approvedHours:
        Number(row[6] || 0),

      hourlyRate:
        Number(row[7] || 10.50),

      amount:
        Number(row[8] || 0),

      status:
        String(row[9] || 'PENDING'),

      rowNumber:
        rowIndex + 2
    };
  });

  return approvals;
}function debugJuanitoPayroll() {
  const monthText = '2026-12';

  const approvals =
    getExistingPayrollApprovals_(monthText);

  const recorded =
    getRecordedHoursForPayroll_(monthText);

  const oldAdjustments =
    getPayrollAdjustmentsByStudent_(monthText);

  const shiftAdjustments =
    getPayrollShiftAdjustmentTotals_(monthText);

  Logger.log('APPROVALS: ' + JSON.stringify(approvals));
  Logger.log('RECORDED: ' + JSON.stringify(recorded));
  Logger.log('OLD ADJUSTMENTS: ' + JSON.stringify(oldAdjustments));
  Logger.log('SHIFT ADJUSTMENTS: ' + JSON.stringify(shiftAdjustments));
}
function debugPayrollReview() {
  const result = getPayrollReview('2026-08');

  Logger.log(
    JSON.stringify(result)
  );
}
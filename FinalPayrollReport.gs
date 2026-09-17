  function getFinalPayrollReport(monthText) {
  if (!/^\d{4}-\d{2}$/.test(String(monthText))) {
    throw new Error('Month must use YYYY-MM format.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const approvalsSheet =
    ss.getSheetByName('Payroll Approvals');

  const profilesSheet =
    ss.getSheetByName('Student Profiles');

  if (!approvalsSheet || !profilesSheet) {
    throw new Error(
      'Payroll Approvals or Student Profiles sheet is missing.'
    );
  }

  const tz = Session.getScriptTimeZone();

  const approvals =
    approvalsSheet.getDataRange().getValues();

  const profiles =
    profilesSheet.getDataRange().getValues();

  const profileMap = {};

  profiles.slice(1).forEach(row => {
    const studentId = String(row[0] || '');

    if (!studentId) return;

    profileMap[studentId] = {
      studentName: String(row[1] || ''),
      studentNumber: String(row[2] || ''),
      last4: String(row[3] || '')
    };
  });

  const rows = [];

  approvals.slice(1).forEach(row => {
    const payrollMonth = row[3];

    if (!(payrollMonth instanceof Date)) {
      return;
    }

    const rowMonth =
      Utilities.formatDate(
        payrollMonth,
        tz,
        'yyyy-MM'
      );

    if (rowMonth !== monthText) {
      return;
    }

    const status =
      String(row[9] || '').toUpperCase();

    if (status !== 'APPROVED') {
      return;
    }

    const studentId =
      String(row[1] || '');

    const profile =
      profileMap[studentId] || {};

    const approvedHours =
      Number(row[6] || 0);

    const hourlyRate =
      Number(row[7] || 10.50);

    const amount =
      Number(row[8] || 0);

    rows.push({
      studentId: studentId,

      studentName:
        profile.studentName ||
        String(row[2] || ''),

      studentNumber:
        profile.studentNumber || '',

      last4:
        profile.last4 || '',

      approvedHours:
        Math.round(approvedHours * 100) / 100,

      hourlyRate:
        Math.round(hourlyRate * 100) / 100,

      amount:
        Math.round(amount * 100) / 100
    });
  });

  rows.sort((a, b) =>
    a.studentName.localeCompare(b.studentName)
  );

  const grandHours =
    Math.round(
      rows.reduce(
        (sum, row) =>
          sum + Number(row.approvedHours || 0),
        0
      ) * 100
    ) / 100;

  const grandTotal =
    Math.round(
      rows.reduce(
        (sum, row) =>
          sum + Number(row.amount || 0),
        0
      ) * 100
    ) / 100;

  return {
    month: monthText,
    rows: rows,
    studentCount: rows.length,
    grandHours: grandHours,
    grandTotal: grandTotal
  };
}function exportFinalPayrollExcel(monthText) {
  const report = getFinalPayrollReport(monthText);

  if (!report.rows.length) {
    throw new Error(
      'There are no approved payroll records for this month.'
    );
  }

  const tempSpreadsheet = SpreadsheetApp.create(
    'FIESTA Hub Payroll ' + monthText
  );

  const sheet = tempSpreadsheet.getSheets()[0];
  sheet.setName('Stipend Report');

  buildPayrollExportSheet_(sheet, report);

  SpreadsheetApp.flush();

  const spreadsheetId = tempSpreadsheet.getId();

  const exportUrl =
    'https://docs.google.com/spreadsheets/d/' +
    spreadsheetId +
    '/export?format=xlsx';

  const response = UrlFetchApp.fetch(
    exportUrl,
    {
      headers: {
        Authorization:
          'Bearer ' +
          ScriptApp.getOAuthToken()
      }
    }
  );

  const fileName =
    'FIESTA_Hub_Stipend_Report_' +
    monthText +
    '.xlsx';

  const blob =
    response.getBlob().setName(fileName);

  const folder = getPayrollExportFolder_();

  const file =
    folder.createFile(blob);

  // Trash the temporary Google Sheet.
  DriveApp
    .getFileById(spreadsheetId)
    .setTrashed(true);

  return {
    ok: true,
    fileName: fileName,
    url: file.getUrl(),
    message:
      'Excel payroll report created successfully.'
  };
}


function exportFinalPayrollPDF(monthText) {
  const report = getFinalPayrollReport(monthText);

  if (!report.rows.length) {
    throw new Error(
      'There are no approved payroll records for this month.'
    );
  }

  const tempSpreadsheet = SpreadsheetApp.create(
    'FIESTA Hub Payroll PDF ' + monthText
  );

  const sheet = tempSpreadsheet.getSheets()[0];
  sheet.setName('Stipend Report');

  buildPayrollExportSheet_(sheet, report);

  SpreadsheetApp.flush();

  const spreadsheetId = tempSpreadsheet.getId();
  const sheetId = sheet.getSheetId();

  const exportUrl =
    'https://docs.google.com/spreadsheets/d/' +
    spreadsheetId +
    '/export?' +
    'format=pdf' +
    '&gid=' + sheetId +
    '&size=letter' +
    '&portrait=false' +
    '&fitw=true' +
    '&sheetnames=false' +
    '&printtitle=false' +
    '&pagenumbers=true' +
    '&gridlines=false' +
    '&fzr=true';

  const response = UrlFetchApp.fetch(
    exportUrl,
    {
      headers: {
        Authorization:
          'Bearer ' +
          ScriptApp.getOAuthToken()
      }
    }
  );

  const fileName =
    'FIESTA_Hub_Stipend_Report_' +
    monthText +
    '.pdf';

  const blob =
    response.getBlob().setName(fileName);

  const folder = getPayrollExportFolder_();

  const file =
    folder.createFile(blob);

  DriveApp
    .getFileById(spreadsheetId)
    .setTrashed(true);

  return {
    ok: true,
    fileName: fileName,
    url: file.getUrl(),
    message:
      'PDF payroll report created successfully.'
  };
}


function buildPayrollExportSheet_(sheet, report) {
  const headers = [
    'Name',
    'Student Number',
    'Last 4 SSN',
    'Approved Hours',
    'Hourly Rate',
    'Amount'
  ];

  sheet.clear();

  // Title
  sheet.getRange('A1:F1')
    .merge()
    .setValue('FIESTA IX - Student Stipend Report')
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Month
  sheet.getRange('A2:F2')
    .merge()
    .setValue(
      'Payroll Month: ' +
      formatPayrollMonth_(report.month)
    )
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Headers
  sheet.getRange(
    4,
    1,
    1,
    headers.length
  )
    .setValues([headers])
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  const data = report.rows.map(row => [
    row.studentName,
    row.studentNumber,
    row.last4,
    row.approvedHours,
    row.hourlyRate,
    row.amount
  ]);

  if (data.length) {
    sheet.getRange(
      5,
      1,
      data.length,
      headers.length
    ).setValues(data);
  }

  const totalRow =
    5 + data.length;

  sheet.getRange(
    totalRow,
    1,
    1,
    3
  )
    .merge()
    .setValue('TOTAL')
    .setFontWeight('bold');

  sheet.getRange(
    totalRow,
    4
  )
    .setValue(report.grandHours)
    .setFontWeight('bold');

  sheet.getRange(
    totalRow,
    6
  )
    .setValue(report.grandTotal)
    .setFontWeight('bold');

  sheet.getRange(
    5,
    4,
    Math.max(data.length + 1, 1),
    1
  ).setNumberFormat('0.00');

  sheet.getRange(
    5,
    5,
    Math.max(data.length, 1),
    2
  ).setNumberFormat('$0.00');

  sheet.getRange(
    totalRow,
    6
  ).setNumberFormat('$0.00');

  sheet.setFrozenRows(4);

  sheet.autoResizeColumns(
    1,
    6
  );

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 120);
}


function getPayrollExportFolder_() {
  const folderName =
    'FIESTA Hub Payroll Reports';

  const folders =
    DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  }

  return DriveApp.createFolder(folderName);
}


function formatPayrollMonth_(monthText) {
  const parts =
    String(monthText)
      .split('-')
      .map(Number);

  const date =
    new Date(
      parts[0],
      parts[1] - 1,
      1
    );

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'MMMM yyyy'
  );
}function authorizePayrollDrive() {
  const folders = DriveApp.getFoldersByName(
    'FIESTA Hub Payroll Reports'
  );

  if (folders.hasNext()) {
    return folders.next().getName();
  }

  return DriveApp
    .createFolder('FIESTA Hub Payroll Reports')
    .getName();
}
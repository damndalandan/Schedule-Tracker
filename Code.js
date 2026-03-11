// ============================================================
// OPS TRIP MONITORING SYSTEM — Code.gs
// Google Apps Script Backend
// ============================================================

const SHEET_VEHICLES   = 'Vehicles';
const SHEET_TRIPS      = 'Trips';
const SHEET_SETTINGS   = 'Settings';
const SHEET_RENEWALS   = 'Renewal Alerts';
const SHEET_EMPLOYEES  = 'Employees'; // local fallback; replace with IMPORTRANGE if needed

// Replace with your employee sheet ID when ready
const EMPLOYEE_SHEET_ID = 'YOUR_EMPLOYEE_SHEET_ID_HERE';

// ============================================================
// WEB APP ENTRY POINTS
// ============================================================

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('OPS Trip Monitoring System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// SHEET INITIALIZATION
// ============================================================

function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _ensureVehiclesSheet(ss);
  _ensureTripsSheet(ss);
  _ensureSettingsSheet(ss);
  _ensureRenewalAlertsSheet(ss);
  _ensureEmployeesSheet(ss);
  return { success: true, message: 'All sheets initialized.' };
}

function _ensureVehiclesSheet(ss) {
  let sh = ss.getSheetByName(SHEET_VEHICLES);
  if (!sh) {
    sh = ss.insertSheet(SHEET_VEHICLES);
    sh.appendRow([
      'Vehicle ID','Plate Number','Vehicle Type','Brand/Model',
      'Beginning Mileage','Status','Insurance Expiry Date','Insurance PDF Link',
      'LTO Expiry Date','LTO PDF Link','Notes','Created At','Updated At'
    ]);
    sh.getRange(1,1,1,13).setFontWeight('bold').setBackground('#1a3c5e').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
}

function _ensureTripsSheet(ss) {
  let sh = ss.getSheetByName(SHEET_TRIPS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_TRIPS);
    sh.appendRow([
      'Trip ID','Request Date','Requestor Employee ID','Requestor Name',
      'Trip Type','Purpose','Related JO','From Location','To Location',
      'Planned Start DateTime','Planned End DateTime','Vehicle ID','Plate Number',
      'Driver Employee ID','Driver Name','Status','Approved By',
      'Approval/Rejection Date','Rejection Reason','Cancel Reason',
      'Actual Start DateTime','Actual End DateTime','Start Mileage','End Mileage',
      'Distance Travelled','GPS / Proof Link','Remarks','Updated At','Updated By'
    ]);
    sh.getRange(1,1,1,29).setFontWeight('bold').setBackground('#1a3c5e').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
}

function _ensureSettingsSheet(ss) {
  let sh = ss.getSheetByName(SHEET_SETTINGS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_SETTINGS);
    sh.appendRow(['Category','Value']);
    sh.getRange(1,1,1,2).setFontWeight('bold').setBackground('#1a3c5e').setFontColor('#ffffff');
    const defaults = [
      ['Trip Type','Owner Errand'],
      ['Trip Type','Supplier Delivery - Ormoc'],
      ['Trip Type','Supplier Delivery - Outside Ormoc'],
      ['Trip Type','Signage Installation'],
      ['Trip Type','Other'],
      ['Vehicle Status','Active'],
      ['Vehicle Status','Under Repair'],
      ['Vehicle Status','Inactive'],
      ['Trip Status','Draft'],
      ['Trip Status','Submitted'],
      ['Trip Status','Approved'],
      ['Trip Status','Rejected'],
      ['Trip Status','Cancelled'],
      ['Trip Status','Completed'],
      ['Vehicle Type','Van'],
      ['Vehicle Type','Truck'],
      ['Vehicle Type','Motorcycle'],
      ['Vehicle Type','Car'],
    ];
    sh.getRange(2, 1, defaults.length, 2).setValues(defaults);
  }
}

function _ensureRenewalAlertsSheet(ss) {
  let sh = ss.getSheetByName(SHEET_RENEWALS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_RENEWALS);
    sh.appendRow(['Vehicle ID','Plate Number','Document Type','Expiry Date','Days Left','Alert Status']);
    sh.getRange(1,1,1,6).setFontWeight('bold').setBackground('#1a3c5e').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
}

function _ensureEmployeesSheet(ss) {
  let sh = ss.getSheetByName(SHEET_EMPLOYEES);
  if (!sh) {
    sh = ss.insertSheet(SHEET_EMPLOYEES);
    sh.appendRow(['Employee ID','Full Name','Email','Department','Active']);
    sh.getRange(1,1,1,5).setFontWeight('bold').setBackground('#1a3c5e').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    // Sample employees
    sh.appendRow(['EMP-001','Juan dela Cruz','juan@example.com','Operations','Yes']);
    sh.appendRow(['EMP-002','Maria Santos','maria@example.com','Admin','Yes']);
    sh.appendRow(['EMP-003','Pedro Reyes','pedro@example.com','Driver','Yes']);
  }
}

// ============================================================
// ID GENERATORS
// ============================================================

function _generateVehicleId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_VEHICLES);
  const year = new Date().getFullYear();
  const lastRow = sh.getLastRow();
  let seq = 1;
  if (lastRow > 1) {
    const ids = sh.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(String);
    const nums = ids.map(id => parseInt(id.split('-')[2] || '0'));
    seq = Math.max(...nums) + 1;
  }
  return `V-${year}-${String(seq).padStart(4,'0')}`;
}

function _generateTripId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_TRIPS);
  const year = new Date().getFullYear();
  const lastRow = sh.getLastRow();
  let seq = 1;
  if (lastRow > 1) {
    const ids = sh.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(String);
    const nums = ids.map(id => parseInt(id.split('-')[2] || '0'));
    seq = Math.max(...nums) + 1;
  }
  return `T-${year}-${String(seq).padStart(4,'0')}`;
}

// ============================================================
// VEHICLES CRUD
// ============================================================

function getVehicles() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_VEHICLES);
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function getVehicleById(vehicleId) {
  const vehicles = getVehicles();
  return vehicles.find(v => v['Vehicle ID'] === vehicleId) || null;
}

function addVehicle(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_VEHICLES);
    // Check plate uniqueness
    const existing = getVehicles();
    if (existing.some(v => v['Plate Number'] === data.plateNumber)) {
      return { success: false, message: 'Plate number already exists.' };
    }
    const id = _generateVehicleId();
    const now = new Date();
    sh.appendRow([
      id, data.plateNumber, data.vehicleType, data.brandModel || '',
      data.beginningMileage || 0, data.status || 'Active',
      data.insuranceExpiry || '', data.insurancePdfLink || '',
      data.ltoExpiry || '', data.ltoPdfLink || '',
      data.notes || '', now, now
    ]);
    return { success: true, id, message: 'Vehicle added successfully.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateVehicle(vehicleId, data) {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_VEHICLES);
    const allData = sh.getDataRange().getValues();
    const rowIndex = allData.findIndex((r, i) => i > 0 && r[0] === vehicleId);
    if (rowIndex === -1) return { success: false, message: 'Vehicle not found.' };
    const now = new Date();
    const r = rowIndex + 1;
    sh.getRange(r, 2).setValue(data.plateNumber);
    sh.getRange(r, 3).setValue(data.vehicleType);
    sh.getRange(r, 4).setValue(data.brandModel || '');
    sh.getRange(r, 5).setValue(data.beginningMileage || 0);
    sh.getRange(r, 6).setValue(data.status);
    sh.getRange(r, 7).setValue(data.insuranceExpiry || '');
    sh.getRange(r, 8).setValue(data.insurancePdfLink || '');
    sh.getRange(r, 9).setValue(data.ltoExpiry || '');
    sh.getRange(r, 10).setValue(data.ltoPdfLink || '');
    sh.getRange(r, 11).setValue(data.notes || '');
    sh.getRange(r, 13).setValue(now);
    return { success: true, message: 'Vehicle updated.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================
// TRIPS CRUD
// ============================================================

function getTrips() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TRIPS);
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function getTripById(tripId) {
  const trips = getTrips();
  return trips.find(t => t['Trip ID'] === tripId) || null;
}

function createTrip(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_TRIPS);
    if (!data.purpose) return { success: false, message: 'Purpose is required.' };
    if (!data.fromLocation) return { success: false, message: 'From Location is required.' };
    if (!data.toLocation) return { success: false, message: 'To Location is required.' };
    if (!data.plannedStart) return { success: false, message: 'Planned Start is required.' };
    if (!data.plannedEnd) return { success: false, message: 'Planned End is required.' };
    const id = _generateTripId();
    const now = new Date();
    const user = Session.getActiveUser().getEmail();
    sh.appendRow([
      id, now,
      data.requestorEmpId || '', data.requestorName || '',
      data.tripType || '', data.purpose,
      data.relatedJO || '', data.fromLocation, data.toLocation,
      data.plannedStart, data.plannedEnd,
      data.vehicleId || '', data.plateNumber || '',
      data.driverEmpId || '', data.driverName || '',
      'Draft', '', '', '', '',
      '', '', '', '', '', '', '',
      now, user
    ]);
    return { success: true, id, message: 'Trip created as Draft.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateTripStatus(tripId, action, extras) {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TRIPS);
    const allData = sh.getDataRange().getValues();
    const rowIndex = allData.findIndex((r, i) => i > 0 && r[0] === tripId);
    if (rowIndex === -1) return { success: false, message: 'Trip not found.' };
    const now = new Date();
    const user = Session.getActiveUser().getEmail();
    const r = rowIndex + 1;
    const currentStatus = allData[rowIndex][15];

    if (action === 'submit') {
      if (currentStatus !== 'Draft') return { success: false, message: 'Only Draft trips can be submitted.' };
      if (!allData[rowIndex][11]) return { success: false, message: 'Vehicle is required before submitting.' };
      if (!allData[rowIndex][13]) return { success: false, message: 'Driver is required before submitting.' };
      sh.getRange(r, 16).setValue('Submitted');
    } else if (action === 'approve') {
      if (currentStatus !== 'Submitted') return { success: false, message: 'Only Submitted trips can be approved.' };
      sh.getRange(r, 16).setValue('Approved');
      sh.getRange(r, 17).setValue(user);
      sh.getRange(r, 18).setValue(now);
    } else if (action === 'reject') {
      if (!extras.reason) return { success: false, message: 'Rejection reason is required.' };
      sh.getRange(r, 16).setValue('Rejected');
      sh.getRange(r, 17).setValue(user);
      sh.getRange(r, 18).setValue(now);
      sh.getRange(r, 19).setValue(extras.reason);
    } else if (action === 'cancel') {
      if (!['Draft','Submitted','Approved'].includes(currentStatus))
        return { success: false, message: 'Cannot cancel a trip in current status.' };
      if (!extras.reason) return { success: false, message: 'Cancel reason is required.' };
      sh.getRange(r, 16).setValue('Cancelled');
      sh.getRange(r, 20).setValue(extras.reason);
    } else if (action === 'complete') {
      if (currentStatus !== 'Approved') return { success: false, message: 'Only Approved trips can be completed.' };
      if (!extras.actualStart) return { success: false, message: 'Actual Start is required.' };
      if (!extras.actualEnd) return { success: false, message: 'Actual End is required.' };
      if (!extras.startMileage && extras.startMileage !== 0) return { success: false, message: 'Start mileage is required.' };
      if (!extras.endMileage && extras.endMileage !== 0) return { success: false, message: 'End mileage is required.' };
      if (Number(extras.endMileage) < Number(extras.startMileage))
        return { success: false, message: 'End mileage cannot be less than start mileage.' };
      const distance = Number(extras.endMileage) - Number(extras.startMileage);
      sh.getRange(r, 16).setValue('Completed');
      sh.getRange(r, 21).setValue(extras.actualStart);
      sh.getRange(r, 22).setValue(extras.actualEnd);
      sh.getRange(r, 23).setValue(extras.startMileage);
      sh.getRange(r, 24).setValue(extras.endMileage);
      sh.getRange(r, 25).setValue(distance);
      sh.getRange(r, 26).setValue(extras.proofLink || '');
      sh.getRange(r, 27).setValue(extras.remarks || '');
    }

    sh.getRange(r, 28).setValue(now);
    sh.getRange(r, 29).setValue(user);

    // Send email notifications
    _sendTripNotification(action, tripId, allData[rowIndex]);

    return { success: true, message: `Trip ${action}d successfully.` };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateTripDraft(tripId, data) {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TRIPS);
    const allData = sh.getDataRange().getValues();
    const rowIndex = allData.findIndex((r, i) => i > 0 && r[0] === tripId);
    if (rowIndex === -1) return { success: false, message: 'Trip not found.' };
    if (allData[rowIndex][15] !== 'Draft') return { success: false, message: 'Only Draft trips can be edited.' };
    const now = new Date();
    const user = Session.getActiveUser().getEmail();
    const r = rowIndex + 1;
    sh.getRange(r, 3).setValue(data.requestorEmpId || '');
    sh.getRange(r, 4).setValue(data.requestorName || '');
    sh.getRange(r, 5).setValue(data.tripType || '');
    sh.getRange(r, 6).setValue(data.purpose || '');
    sh.getRange(r, 7).setValue(data.relatedJO || '');
    sh.getRange(r, 8).setValue(data.fromLocation || '');
    sh.getRange(r, 9).setValue(data.toLocation || '');
    sh.getRange(r, 10).setValue(data.plannedStart || '');
    sh.getRange(r, 11).setValue(data.plannedEnd || '');
    sh.getRange(r, 12).setValue(data.vehicleId || '');
    sh.getRange(r, 13).setValue(data.plateNumber || '');
    sh.getRange(r, 14).setValue(data.driverEmpId || '');
    sh.getRange(r, 15).setValue(data.driverName || '');
    sh.getRange(r, 28).setValue(now);
    sh.getRange(r, 29).setValue(user);
    return { success: true, message: 'Draft updated.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================
// EMPLOYEES LOOKUP
// ============================================================

function getEmployees() {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EMPLOYEES);
    if (!sh) return [];
    const data = sh.getDataRange().getValues();
    if (data.length <= 1) return [];
    const headers = data[0];
    return data.slice(1)
      .filter(row => row[4] === 'Yes')
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = row[i]);
        return obj;
      });
  } catch (e) {
    return [];
  }
}

function getEmployeeById(empId) {
  const employees = getEmployees();
  return employees.find(e => e['Employee ID'] === empId) || null;
}

// ============================================================
// SETTINGS
// ============================================================

function getSettings() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  const data = sh.getDataRange().getValues();
  const result = {};
  data.slice(1).forEach(row => {
    const cat = row[0], val = row[1];
    if (!result[cat]) result[cat] = [];
    result[cat].push(val);
  });
  return result;
}

// ============================================================
// RENEWAL ALERTS
// ============================================================

function refreshRenewalAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_RENEWALS);
  // Clear existing data (keep header)
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 6).clearContent();
  const vehicles = getVehicles();
  const today = new Date();
  const rows = [];
  vehicles.forEach(v => {
    ['Insurance', 'LTO'].forEach(docType => {
      const expiryField = docType === 'Insurance' ? 'Insurance Expiry Date' : 'LTO Expiry Date';
      const expiry = v[expiryField];
      if (!expiry) return;
      const expiryDate = new Date(expiry);
      const daysLeft = Math.ceil((expiryDate - today) / (1000 * 60 * 60 * 24));
      let alertStatus = 'OK';
      if (daysLeft < 0) alertStatus = 'Expired';
      else if (daysLeft <= 30) alertStatus = 'Due in 30 Days';
      rows.push([v['Vehicle ID'], v['Plate Number'], docType, expiryDate, daysLeft, alertStatus]);
    });
  });
  if (rows.length > 0) sh.getRange(2, 1, rows.length, 6).setValues(rows);
  return { success: true, alerts: rows.length };
}

function getRenewalAlerts() {
  refreshRenewalAlerts();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RENEWALS);
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

// ============================================================
// EMAIL NOTIFICATIONS
// ============================================================

function _sendTripNotification(action, tripId, rowData) {
  try {
    const approverEmail = 'YOUR_APPROVER_EMAIL@example.com'; // TODO: replace
    const adminEmail = Session.getActiveUser().getEmail();
    let subject = '', body = '';

    if (action === 'submit') {
      subject = `[OPS] Trip ${tripId} Submitted for Approval`;
      body = `A new trip request (${tripId}) has been submitted and is awaiting your approval.\n\nPurpose: ${rowData[5]}\nFrom: ${rowData[7]} → To: ${rowData[8]}\nPlanned Start: ${rowData[9]}`;
      MailApp.sendEmail(approverEmail, subject, body);
    } else if (action === 'approve') {
      subject = `[OPS] Trip ${tripId} Approved`;
      body = `Your trip request (${tripId}) has been approved.\n\nPurpose: ${rowData[5]}\nFrom: ${rowData[7]} → To: ${rowData[8]}`;
      MailApp.sendEmail(adminEmail, subject, body);
    } else if (action === 'reject') {
      subject = `[OPS] Trip ${tripId} Rejected`;
      body = `Your trip request (${tripId}) has been rejected.\n\nPurpose: ${rowData[5]}\nRejection Reason: ${rowData[18]}`;
      MailApp.sendEmail(adminEmail, subject, body);
    }
  } catch (e) {
    console.log('Email error: ' + e.message);
  }
}

function sendRenewalWarningEmails() {
  const alerts = getRenewalAlerts();
  const adminEmail = 'YOUR_ADMIN_EMAIL@example.com'; // TODO: replace
  const expiring = alerts.filter(a => a['Alert Status'] !== 'OK');
  if (expiring.length === 0) return;
  let body = 'The following vehicle documents need attention:\n\n';
  expiring.forEach(a => {
    body += `• Vehicle ${a['Plate Number']} — ${a['Document Type']} — ${a['Alert Status']} (${a['Days Left']} days left)\n`;
  });
  MailApp.sendEmail(adminEmail, '[OPS] Vehicle Document Renewal Alert', body);
}

// ============================================================
// DAILY TRIGGER SETUP
// ============================================================

function setupDailyTrigger() {
  // Run once to install daily trigger
  ScriptApp.newTrigger('sendRenewalWarningEmails')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
  return { success: true, message: 'Daily renewal trigger set for 8AM.' };
}

// ============================================================
// REPORTS
// ============================================================

function getReportTripsByVehicle() {
  const trips = getTrips().filter(t => t['Status'] === 'Completed');
  const map = {};
  trips.forEach(t => {
    const key = t['Vehicle ID'] || 'Unknown';
    if (!map[key]) map[key] = { vehicleId: key, plateNumber: t['Plate Number'], tripCount: 0, totalKm: 0 };
    map[key].tripCount++;
    map[key].totalKm += Number(t['Distance Travelled']) || 0;
  });
  return Object.values(map);
}

function getReportTripsByDriver() {
  const trips = getTrips().filter(t => t['Status'] === 'Completed');
  const map = {};
  trips.forEach(t => {
    const key = t['Driver Employee ID'] || 'Unknown';
    if (!map[key]) map[key] = { driverEmpId: key, driverName: t['Driver Name'], tripCount: 0, totalKm: 0 };
    map[key].tripCount++;
    map[key].totalKm += Number(t['Distance Travelled']) || 0;
  });
  return Object.values(map);
}

function getReportTripsByType() {
  const trips = getTrips();
  const map = {};
  trips.forEach(t => {
    const key = t['Trip Type'] || 'Unknown';
    if (!map[key]) map[key] = { tripType: key, count: 0 };
    map[key].count++;
  });
  return Object.values(map);
}

function getReportMileageSummary() {
  const vehicles = getVehicles();
  const trips = getTrips().filter(t => t['Status'] === 'Completed');
  return vehicles.map(v => {
    const vTrips = trips.filter(t => t['Vehicle ID'] === v['Vehicle ID']);
    const latestEnd = vTrips.length > 0 ? Math.max(...vTrips.map(t => Number(t['End Mileage']) || 0)) : 0;
    const totalTravel = vTrips.reduce((s, t) => s + (Number(t['Distance Travelled']) || 0), 0);
    return {
      vehicleId: v['Vehicle ID'], 
      plateNumber: v['Plate Number'],
      beginningMileage: v['Beginning Mileage'],
      latestEndMileage: latestEnd,
      totalTravel
    };
  });
}
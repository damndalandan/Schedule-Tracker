// ============================================================
// OPS TRIP MONITORING SYSTEM
// Code.gs — Main Server-Side Script
// ============================================================

// ── CONFIGURATION ───────────────────────────────────────────
const CONFIG = {
  SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID_HERE', // ← Replace with your Sheet ID
  SHEETS: {
    EMPLOYEES:      'Employees',
    VEHICLES:       'Vehicles',
    TRIPS:          'Trips',
    SETTINGS:       'Settings',
    RENEWAL_ALERTS: 'Renewal Alerts'
  }
};

// ── ENTRY POINT ─────────────────────────────────────────────
function doGet(e) {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('OPS Trip Monitoring System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Include helper for HTML templates
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── SPREADSHEET HELPER ───────────────────────────────────────
function getSheet(name) {
  return SpreadsheetApp
    .openById(CONFIG.SPREADSHEET_ID)
    .getSheetByName(name);
}

function getSheetData(name) {
  const sheet = getSheet(name);
  const data  = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

// ── SETUP: Initialize Sheet Structure ───────────────────────
function setupSheets() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  const sheetsConfig = {
    'Vehicles': [
      'Vehicle ID','Plate Number','Vehicle Type','Brand/Model',
      'Beginning Mileage','Status','Insurance Expiry Date','Insurance PDF Link',
      'LTO Expiry Date','LTO PDF Link','Notes','Created At','Updated At'
    ],
    'Trips': [
      'Trip ID','Request Date','Requestor Employee ID','Requestor Name',
      'Trip Type','Purpose','Related JO','From Location','To Location',
      'Planned Start DateTime','Planned End DateTime','Vehicle ID','Plate Number',
      'Driver Employee ID','Driver Name','Status','Approved By',
      'Approval/Rejection Date','Rejection Reason','Cancel Reason',
      'Actual Start DateTime','Actual End DateTime','Start Mileage','End Mileage',
      'Distance Travelled','GPS/Proof Link','Remarks','Updated At','Updated By'
    ],
    'Settings': [
      'Category','Value','Sort Order'
    ],
    'Renewal Alerts': [
      'Vehicle ID','Plate Number','Document Type','Expiry Date','Days Left','Alert Status'
    ]
  };

  Object.entries(sheetsConfig).forEach(([name, headers]) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#1a1a2e')
        .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
  });

  // Seed Settings if empty
  seedSettings();
}

function seedSettings() {
  const sheet = getSheet('Settings');
  const data  = sheet.getDataRange().getValues();
  if (data.length > 1) return;

  const settings = [
    ['Trip Type', 'Owner Errand', 1],
    ['Trip Type', 'Supplier Delivery - Ormoc', 2],
    ['Trip Type', 'Supplier Delivery - Outside Ormoc', 3],
    ['Trip Type', 'Signage Installation', 4],
    ['Trip Type', 'Other', 5],
    ['Vehicle Type', 'Van', 1],
    ['Vehicle Type', 'Truck', 2],
    ['Vehicle Type', 'Motorcycle', 3],
    ['Vehicle Type', 'Car', 4],
    ['Vehicle Type', 'Tricycle', 5],
    ['Vehicle Status', 'Active', 1],
    ['Vehicle Status', 'Under Repair', 2],
    ['Vehicle Status', 'Inactive', 3],
    ['Trip Status', 'Draft', 1],
    ['Trip Status', 'Submitted', 2],
    ['Trip Status', 'Approved', 3],
    ['Trip Status', 'Rejected', 4],
    ['Trip Status', 'Cancelled', 5],
    ['Trip Status', 'Completed', 6],
  ];
  settings.forEach(row => sheet.appendRow(row));
}

// ── ID GENERATION ────────────────────────────────────────────
function generateId(prefix, sheetName, idColumn) {
  const sheet = getSheet(sheetName);
  const year  = new Date().getFullYear();
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return `${prefix}-${year}-0001`;

  const ids = data.slice(1)
    .map(r => r[0])
    .filter(id => id && String(id).startsWith(`${prefix}-${year}-`));

  if (ids.length === 0) return `${prefix}-${year}-0001`;

  const nums = ids.map(id => parseInt(id.split('-')[2], 10));
  const next = Math.max(...nums) + 1;
  return `${prefix}-${year}-${String(next).padStart(4, '0')}`;
}

// ── EMPLOYEES ────────────────────────────────────────────────
function getEmployees() {
  return getSheetData(CONFIG.SHEETS.EMPLOYEES);
}

function getEmployeeById(employeeId) {
  const employees = getEmployees();
  return employees.find(e => String(e['Employee ID']) === String(employeeId)) || null;
}

function getEmployeesByRole(role) {
  return getEmployees().filter(e => e['Role'] === role);
}

// ── SETTINGS ─────────────────────────────────────────────────
function getSettings() {
  const data = getSheetData(CONFIG.SHEETS.SETTINGS);
  const grouped = {};
  data.forEach(row => {
    const cat = row['Category'];
    if (!grouped[cat]) grouped[cat] = [];
    grouped[cat].push(row['Value']);
  });
  return grouped;
}

// ── VEHICLES ─────────────────────────────────────────────────
function getVehicles() {
  return getSheetData(CONFIG.SHEETS.VEHICLES);
}

function getActiveVehicles() {
  return getVehicles().filter(v => v['Status'] === 'Active');
}

function getVehicleById(vehicleId) {
  return getVehicles().find(v => String(v['Vehicle ID']) === String(vehicleId)) || null;
}

function saveVehicle(data) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.VEHICLES);
    const now   = new Date();

    // Validate required fields
    if (!data['Plate Number']) return { success: false, message: 'Plate Number is required.' };
    if (!data['Vehicle Type']) return { success: false, message: 'Vehicle Type is required.' };
    if (!data['Status'])       return { success: false, message: 'Status is required.' };

    // Check plate uniqueness
    const existing = getVehicles();
    const duplicate = existing.find(v =>
      v['Plate Number'].toLowerCase() === data['Plate Number'].toLowerCase() &&
      v['Vehicle ID'] !== data['Vehicle ID']
    );
    if (duplicate) return { success: false, message: 'Plate Number already exists.' };

    if (data['Vehicle ID']) {
      // UPDATE
      const allData = sheet.getDataRange().getValues();
      const headers = allData[0];
      const rowIdx  = allData.findIndex((r, i) => i > 0 && String(r[0]) === String(data['Vehicle ID']));
      if (rowIdx === -1) return { success: false, message: 'Vehicle not found.' };

      headers.forEach((h, i) => {
        if (data[h] !== undefined && h !== 'Vehicle ID' && h !== 'Created At') {
          sheet.getRange(rowIdx + 1, i + 1).setValue(data[h]);
        }
      });
      sheet.getRange(rowIdx + 1, headers.indexOf('Updated At') + 1).setValue(now);
      return { success: true, message: 'Vehicle updated successfully.' };

    } else {
      // CREATE
      const vehicleId = generateId('V', CONFIG.SHEETS.VEHICLES, 'Vehicle ID');
      const headers   = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const row       = headers.map(h => {
        if (h === 'Vehicle ID')  return vehicleId;
        if (h === 'Created At')  return now;
        if (h === 'Updated At')  return now;
        return data[h] !== undefined ? data[h] : '';
      });
      sheet.appendRow(row);
      return { success: true, message: 'Vehicle created successfully.', vehicleId };
    }
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteVehicle(vehicleId) {
  // Soft delete — mark as Inactive
  return saveVehicle({ 'Vehicle ID': vehicleId, 'Status': 'Inactive' });
}

// ── TRIPS ────────────────────────────────────────────────────
function getTrips() {
  return getSheetData(CONFIG.SHEETS.TRIPS);
}

function getTripById(tripId) {
  return getTrips().find(t => String(t['Trip ID']) === String(tripId)) || null;
}

function getTripsByStatus(status) {
  return getTrips().filter(t => t['Status'] === status);
}

function getTripsByEmployee(employeeId) {
  return getTrips().filter(t => String(t['Requestor Employee ID']) === String(employeeId));
}

function saveTrip(data) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.TRIPS);
    const now   = new Date();

    // Validate based on status
    const status = data['Status'] || 'Draft';

    if (status === 'Submitted') {
      const required = ['Requestor Employee ID','Trip Type','Purpose','From Location','To Location','Planned Start DateTime','Planned End DateTime','Vehicle ID','Driver Employee ID'];
      for (const field of required) {
        if (!data[field]) return { success: false, message: `${field} is required before submitting.` };
      }
    }

    if (status === 'Rejected' && !data['Rejection Reason']) {
      return { success: false, message: 'Rejection Reason is required.' };
    }

    if (status === 'Cancelled' && !data['Cancel Reason']) {
      return { success: false, message: 'Cancel Reason is required.' };
    }

    if (status === 'Completed') {
      if (!data['Actual Start DateTime']) return { success: false, message: 'Actual Start DateTime is required.' };
      if (!data['Actual End DateTime'])   return { success: false, message: 'Actual End DateTime is required.' };
      if (!data['Start Mileage'])         return { success: false, message: 'Start Mileage is required.' };
      if (!data['End Mileage'])           return { success: false, message: 'End Mileage is required.' };
      if (Number(data['End Mileage']) < Number(data['Start Mileage'])) {
        return { success: false, message: 'End Mileage cannot be less than Start Mileage.' };
      }
      data['Distance Travelled'] = Number(data['End Mileage']) - Number(data['Start Mileage']);
    }

    // Auto-fill from Vehicle
    if (data['Vehicle ID']) {
      const vehicle = getVehicleById(data['Vehicle ID']);
      if (vehicle) data['Plate Number'] = vehicle['Plate Number'];
    }

    // Auto-fill from Employee (Requestor)
    if (data['Requestor Employee ID']) {
      const emp = getEmployeeById(data['Requestor Employee ID']);
      if (emp) data['Requestor Name'] = emp['Employee Name'];
    }

    // Auto-fill Driver Name
    if (data['Driver Employee ID']) {
      const driver = getEmployeeById(data['Driver Employee ID']);
      if (driver) data['Driver Name'] = driver['Employee Name'];
    }

    if (data['Trip ID']) {
      // UPDATE
      const allData = sheet.getDataRange().getValues();
      const headers = allData[0];
      const rowIdx  = allData.findIndex((r, i) => i > 0 && String(r[0]) === String(data['Trip ID']));
      if (rowIdx === -1) return { success: false, message: 'Trip not found.' };

      // Status transition guard
      const currentStatus = allData[rowIdx][headers.indexOf('Status')];
      if (!isValidTransition(currentStatus, status)) {
        return { success: false, message: `Cannot change status from ${currentStatus} to ${status}.` };
      }

      headers.forEach((h, i) => {
        if (data[h] !== undefined && h !== 'Trip ID' && h !== 'Request Date') {
          sheet.getRange(rowIdx + 1, i + 1).setValue(data[h]);
        }
      });
      sheet.getRange(rowIdx + 1, headers.indexOf('Updated At') + 1).setValue(now);
      sheet.getRange(rowIdx + 1, headers.indexOf('Updated By') + 1).setValue(data['Updated By'] || '');

      // Set approval timestamp
      if (status === 'Approved' || status === 'Rejected') {
        sheet.getRange(rowIdx + 1, headers.indexOf('Approval/Rejection Date') + 1).setValue(now);
      }

      return { success: true, message: `Trip ${status} successfully.` };

    } else {
      // CREATE
      const tripId  = generateId('T', CONFIG.SHEETS.TRIPS, 'Trip ID');
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const row     = headers.map(h => {
        if (h === 'Trip ID')      return tripId;
        if (h === 'Request Date') return now;
        if (h === 'Status')       return 'Draft';
        if (h === 'Updated At')   return now;
        if (h === 'Updated By')   return data['Updated By'] || '';
        return data[h] !== undefined ? data[h] : '';
      });
      sheet.appendRow(row);
      return { success: true, message: 'Trip created successfully.', tripId };
    }
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Status transition rules
function isValidTransition(from, to) {
  const allowed = {
    'Draft':     ['Draft', 'Submitted', 'Cancelled'],
    'Submitted': ['Submitted', 'Approved', 'Rejected', 'Cancelled'],
    'Approved':  ['Approved', 'Completed', 'Cancelled'],
    'Rejected':  ['Rejected'],
    'Cancelled': ['Cancelled'],
    'Completed': ['Completed']
  };
  return (allowed[from] || []).includes(to);
}

// ── RENEWAL ALERTS ───────────────────────────────────────────
function refreshRenewalAlerts() {
  try {
    const sheet    = getSheet(CONFIG.SHEETS.RENEWAL_ALERTS);
    const vehicles = getVehicles();
    const now      = new Date();
    now.setHours(0, 0, 0, 0);

    // Clear existing data (keep header)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 6).clearContent();

    const rows = [];
    vehicles.forEach(v => {
      ['Insurance', 'LTO'].forEach(docType => {
        const expiryVal = v[`${docType} Expiry Date`];
        if (!expiryVal) return;

        const expiry   = new Date(expiryVal);
        expiry.setHours(0, 0, 0, 0);
        const daysLeft = Math.round((expiry - now) / (1000 * 60 * 60 * 24));
        let alertStatus = 'OK';
        if (daysLeft < 0)  alertStatus = 'Expired';
        else if (daysLeft <= 30) alertStatus = 'Due in 30 Days';

        rows.push([
          v['Vehicle ID'],
          v['Plate Number'],
          `${docType} Expiry`,
          expiryVal,
          daysLeft,
          alertStatus
        ]);
      });
    });

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 6).setValues(rows);
    }

    return { success: true, data: rows };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Daily trigger — call this via Apps Script time-based trigger
function dailyRenewalCheck() {
  refreshRenewalAlerts();
}

// ── REPORTS ──────────────────────────────────────────────────
function getReportTripsByVehicle() {
  const trips = getTrips().filter(t => t['Status'] === 'Completed');
  const map   = {};
  trips.forEach(t => {
    const key = t['Vehicle ID'] || 'Unknown';
    if (!map[key]) map[key] = { vehicleId: key, plateNumber: t['Plate Number'] || '', tripCount: 0, totalMileage: 0 };
    map[key].tripCount++;
    map[key].totalMileage += Number(t['Distance Travelled']) || 0;
  });
  return Object.values(map);
}

function getReportTripsByDriver() {
  const trips = getTrips().filter(t => t['Status'] === 'Completed');
  const map   = {};
  trips.forEach(t => {
    const key = t['Driver Employee ID'] || 'Unknown';
    if (!map[key]) map[key] = { driverId: key, driverName: t['Driver Name'] || '', tripCount: 0, totalMileage: 0 };
    map[key].tripCount++;
    map[key].totalMileage += Number(t['Distance Travelled']) || 0;
  });
  return Object.values(map);
}

function getReportTripsByType() {
  const trips = getTrips();
  const map   = {};
  trips.forEach(t => {
    const key = t['Trip Type'] || 'Unknown';
    if (!map[key]) map[key] = { tripType: key, tripCount: 0 };
    map[key].tripCount++;
  });
  return Object.values(map);
}

function getReportMileageSummary() {
  const vehicles = getVehicles();
  const trips    = getTrips().filter(t => t['Status'] === 'Completed');

  return vehicles.map(v => {
    const vTrips   = trips.filter(t => String(t['Vehicle ID']) === String(v['Vehicle ID']));
    const totalKm  = vTrips.reduce((sum, t) => sum + (Number(t['Distance Travelled']) || 0), 0);
    const lastTrip = vTrips.sort((a, b) => new Date(b['Actual End DateTime']) - new Date(a['Actual End DateTime']))[0];
    return {
      vehicleId:        v['Vehicle ID'],
      plateNumber:      v['Plate Number'],
      beginningMileage: v['Beginning Mileage'] || 0,
      latestEndMileage: lastTrip ? lastTrip['End Mileage'] : v['Beginning Mileage'] || 0,
      totalRecorded:    totalKm
    };
  });
}

function getReportExpiringDocuments() {
  refreshRenewalAlerts();
  return getSheetData(CONFIG.SHEETS.RENEWAL_ALERTS)
    .filter(r => r['Alert Status'] === 'Expired' || r['Alert Status'] === 'Due in 30 Days');
}

// ── BULK DATA LOADER ─────────────────────────────────────────
// Single call to load all dropdown data needed by the UI
function getAppData() {
  try {
    return {
      success:   true,
      employees: getEmployees(),
      vehicles:  getActiveVehicles(),
      settings:  getSettings()
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
// ============================================================
// OPS TRIP MONITORING SYSTEM
// Code.gs — Main Server-Side Script
// ============================================================

// ── CONFIGURATION ───────────────────────────────────────────
// Leave SPREADSHEET_ID blank to auto-create a new spreadsheet
// on first run. After first run, the ID will be saved to
// Script Properties automatically.
var CONFIG = {
  SPREADSHEET_ID: '',   // ← Leave blank for auto-setup
  SHEETS: {
    EMPLOYEES:      'Employees',
    VEHICLES:       'Vehicles',
    TRIPS:          'Trips',
    SETTINGS:       'Settings',
    RENEWAL_ALERTS: 'Renewal Alerts'
  }
};

// ── GET SPREADSHEET ID (auto or manual) ─────────────────────
function _getSpreadsheetId() {
  // 1. Use hardcoded ID if provided
  if (CONFIG.SPREADSHEET_ID && CONFIG.SPREADSHEET_ID !== '') {
    return CONFIG.SPREADSHEET_ID;
  }
  // 2. Check Script Properties for saved ID
  var props = PropertiesService.getScriptProperties();
  var saved = props.getProperty('SPREADSHEET_ID');
  if (saved) return saved;
  // 3. Auto-create new spreadsheet
  return _createNewSpreadsheet();
}

function _createNewSpreadsheet() {
  var ss  = SpreadsheetApp.create('OPS Trip Monitoring System — Database');
  var id  = ss.getId();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', id);
  Logger.log('New spreadsheet created: ' + ss.getUrl());
  _initializeSpreadsheet(ss);
  return id;
}

// ── ENTRY POINT ─────────────────────────────────────────────
function doGet(e) {
  _getSpreadsheetId(); // Auto-setup on first visit
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('OPS Trip Monitoring System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── SPREADSHEET HELPERS ──────────────────────────────────────
function _getSS() {
  return SpreadsheetApp.openById(_getSpreadsheetId());
}

function getSheet(name) {
  return _getSS().getSheetByName(name);
}

function getSheetData(name) {
  var sheet = getSheet(name);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  var result  = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    result.push(obj);
  }
  return result;
}

// ── INITIALIZE ALL SHEETS ────────────────────────────────────
function _initializeSpreadsheet(ss) {
  ss = ss || _getSS();
  _setupEmployeesSheet(ss);
  _setupVehiclesSheet(ss);
  _setupTripsSheet(ss);
  _setupSettingsSheet(ss);
  _setupRenewalAlertsSheet(ss);
  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet) ss.deleteSheet(defaultSheet);
}

// ── Run manually from Apps Script editor if needed ───────────
function setupDatabase() {
  var props = PropertiesService.getScriptProperties();
  if (CONFIG.SPREADSHEET_ID && CONFIG.SPREADSHEET_ID !== '') {
    props.setProperty('SPREADSHEET_ID', CONFIG.SPREADSHEET_ID);
  }
  _initializeSpreadsheet();
  Logger.log('Database setup complete. URL: ' + _getSS().getUrl());
}

// ── EMPLOYEES SHEET ──────────────────────────────────────────
function _setupEmployeesSheet(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.EMPLOYEES);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEETS.EMPLOYEES);

  var headers = ['Employee ID', 'Employee Name', 'Role', 'Email', 'Status'];
  if (sheet.getLastRow() < 1) {
    sheet.appendRow(headers);
    _styleHeader(sheet, headers.length);
  }

  if (sheet.getLastRow() <= 1) {
    var employees = [
      ['EMP-001', 'Admin User',       'Admin',     'admin@ormocprintshoppe.com',  'Active'],
      ['EMP-002', 'Maria Santos',     'Approver',  'maria@ormocprintshoppe.com',  'Active'],
      ['EMP-003', 'Juan dela Cruz',   'Requestor', 'juan@ormocprintshoppe.com',   'Active'],
      ['EMP-004', 'Ana Reyes',        'Requestor', 'ana@ormocprintshoppe.com',    'Active'],
      ['EMP-005', 'Pedro Villanueva', 'Requestor', 'pedro@ormocprintshoppe.com',  'Active'],
      ['EMP-006', 'Rosa Mendoza',     'Auditor',   'rosa@ormocprintshoppe.com',   'Active'],
      ['EMP-007', 'Carlo Bautista',   'Requestor', 'carlo@ormocprintshoppe.com',  'Active'],
      ['EMP-008', 'Liza Fernandez',   'Admin',     'liza@ormocprintshoppe.com',   'Active']
    ];
    for (var i = 0; i < employees.length; i++) sheet.appendRow(employees[i]);
    sheet.setFrozenRows(1);
    _autoResizeColumns(sheet, headers.length);
  }
}

// ── VEHICLES SHEET ───────────────────────────────────────────
function _setupVehiclesSheet(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.VEHICLES);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEETS.VEHICLES);

  var headers = [
    'Vehicle ID','Plate Number','Vehicle Type','Brand/Model',
    'Beginning Mileage','Status','Insurance Expiry Date','Insurance PDF Link',
    'LTO Expiry Date','LTO PDF Link','Notes','Created At','Updated At'
  ];
  if (sheet.getLastRow() < 1) {
    sheet.appendRow(headers);
    _styleHeader(sheet, headers.length);
  }

  if (sheet.getLastRow() <= 1) {
    var now  = new Date();
    var yr   = now.getFullYear();
    var mo   = now.getMonth();

    var vehicles = [
      ['V-'+yr+'-0001','ORM-1234','Van',       'Toyota Hi-Ace',    15000,'Active',      new Date(yr+1,2,15),  '',new Date(yr+1,5,30), '','Main delivery van',         now,now],
      ['V-'+yr+'-0002','ORM-5678','Motorcycle', 'Honda XRM 125',    8500, 'Active',      new Date(yr,11,31),   '',new Date(yr,11,31),  '','Office messenger bike',      now,now],
      ['V-'+yr+'-0003','ORM-9012','Truck',      'Isuzu Elf',        42000,'Active',      new Date(yr+1,8,20),  '',new Date(yr+1,8,20), '','Heavy delivery truck',       now,now],
      ['V-'+yr+'-0004','ORM-3456','Car',        'Toyota Vios',      22000,'Active',      new Date(yr,mo+1,10), '',new Date(yr+1,3,15), '','Owner errands vehicle',      now,now],
      ['V-'+yr+'-0005','ORM-7890','Tricycle',   'Kawasaki Barako',  5200, 'Under Repair',new Date(yr+1,1,28),  '',new Date(yr+1,1,28), '','Small local deliveries',     now,now]
    ];

    for (var i = 0; i < vehicles.length; i++) sheet.appendRow(vehicles[i]);
    sheet.getRange(2,7,vehicles.length,1).setNumberFormat('mmm dd, yyyy');
    sheet.getRange(2,9,vehicles.length,1).setNumberFormat('mmm dd, yyyy');
    sheet.getRange(2,12,vehicles.length,2).setNumberFormat('mmm dd, yyyy hh:mm');
    sheet.setFrozenRows(1);
    _autoResizeColumns(sheet, headers.length);
  }
}

// ── TRIPS SHEET ──────────────────────────────────────────────
function _setupTripsSheet(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.TRIPS);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEETS.TRIPS);

  var headers = [
    'Trip ID','Request Date','Requestor Employee ID','Requestor Name',
    'Trip Type','Purpose','Related JO','From Location','To Location',
    'Planned Start DateTime','Planned End DateTime','Vehicle ID','Plate Number',
    'Driver Employee ID','Driver Name','Status','Approved By',
    'Approval/Rejection Date','Rejection Reason','Cancel Reason',
    'Actual Start DateTime','Actual End DateTime','Start Mileage','End Mileage',
    'Distance Travelled','GPS/Proof Link','Remarks','Updated At','Updated By'
  ];
  if (sheet.getLastRow() < 1) {
    sheet.appendRow(headers);
    _styleHeader(sheet, headers.length);
    sheet.setFrozenRows(1);
    _autoResizeColumns(sheet, headers.length);
  }
}

// ── SETTINGS SHEET ───────────────────────────────────────────
function _setupSettingsSheet(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEETS.SETTINGS);

  var headers = ['Category','Value','Sort Order'];
  if (sheet.getLastRow() < 1) {
    sheet.appendRow(headers);
    _styleHeader(sheet, headers.length);
  }

  if (sheet.getLastRow() <= 1) {
    var settings = [
      ['Trip Type','Owner Errand',1],
      ['Trip Type','Supplier Delivery - Ormoc',2],
      ['Trip Type','Supplier Delivery - Outside Ormoc',3],
      ['Trip Type','Signage Installation',4],
      ['Trip Type','Other',5],
      ['Vehicle Type','Van',1],
      ['Vehicle Type','Truck',2],
      ['Vehicle Type','Motorcycle',3],
      ['Vehicle Type','Car',4],
      ['Vehicle Type','Tricycle',5],
      ['Vehicle Status','Active',1],
      ['Vehicle Status','Under Repair',2],
      ['Vehicle Status','Inactive',3],
      ['Trip Status','Draft',1],
      ['Trip Status','Submitted',2],
      ['Trip Status','Approved',3],
      ['Trip Status','Rejected',4],
      ['Trip Status','Cancelled',5],
      ['Trip Status','Completed',6]
    ];
    for (var i = 0; i < settings.length; i++) sheet.appendRow(settings[i]);
    sheet.setFrozenRows(1);
    _autoResizeColumns(sheet, headers.length);
  }
}

// ── RENEWAL ALERTS SHEET ─────────────────────────────────────
function _setupRenewalAlertsSheet(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.RENEWAL_ALERTS);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEETS.RENEWAL_ALERTS);

  var headers = ['Vehicle ID','Plate Number','Document Type','Expiry Date','Days Left','Alert Status'];
  if (sheet.getLastRow() < 1) {
    sheet.appendRow(headers);
    _styleHeader(sheet, headers.length);
    sheet.setFrozenRows(1);
    _autoResizeColumns(sheet, headers.length);
  }
}

// ── STYLE HELPERS ────────────────────────────────────────────
function _styleHeader(sheet, colCount) {
  sheet.getRange(1,1,1,colCount)
    .setFontWeight('bold')
    .setBackground('#1a1a2e')
    .setFontColor('#ffffff')
    .setFontSize(11);
}

function _autoResizeColumns(sheet, colCount) {
  for (var i = 1; i <= colCount; i++) sheet.autoResizeColumn(i);
}

// ── ID GENERATION ────────────────────────────────────────────
function generateId(prefix, sheetName) {
  var sheet = getSheet(sheetName);
  var year  = new Date().getFullYear();
  if (!sheet || sheet.getLastRow() <= 1) return prefix + '-' + year + '-0001';
  var data  = sheet.getDataRange().getValues();
  var nums  = [];
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][0]);
    if (id.indexOf(prefix + '-' + year + '-') === 0) {
      nums.push(parseInt(id.split('-')[2], 10));
    }
  }
  if (nums.length === 0) return prefix + '-' + year + '-0001';
  var next = Math.max.apply(null, nums) + 1;
  var s    = String(next);
  while (s.length < 4) s = '0' + s;
  return prefix + '-' + year + '-' + s;
}

// ── EMPLOYEES ────────────────────────────────────────────────
function getEmployees() {
  var sheet = getSheet(CONFIG.SHEETS.EMPLOYEES);
  if (!sheet) { _initializeSpreadsheet(); }
  return getSheetData(CONFIG.SHEETS.EMPLOYEES);
}

function getEmployeeById(employeeId) {
  var list = getEmployees();
  for (var i = 0; i < list.length; i++) {
    if (String(list[i]['Employee ID']) === String(employeeId)) return list[i];
  }
  return null;
}

// ── SETTINGS ─────────────────────────────────────────────────
function getSettings() {
  var data    = getSheetData(CONFIG.SHEETS.SETTINGS);
  var grouped = {};
  for (var i = 0; i < data.length; i++) {
    var cat = data[i]['Category'];
    if (!grouped[cat]) grouped[cat] = [];
    grouped[cat].push(data[i]['Value']);
  }
  return grouped;
}

// ── VEHICLES ─────────────────────────────────────────────────
function getVehicles() {
  return getSheetData(CONFIG.SHEETS.VEHICLES);
}

function getActiveVehicles() {
  var all    = getVehicles();
  var result = [];
  for (var i = 0; i < all.length; i++) {
    if (all[i]['Status'] === 'Active') result.push(all[i]);
  }
  return result;
}

function getVehicleById(vehicleId) {
  var all = getVehicles();
  for (var i = 0; i < all.length; i++) {
    if (String(all[i]['Vehicle ID']) === String(vehicleId)) return all[i];
  }
  return null;
}

function saveVehicle(data) {
  try {
    var sheet = getSheet(CONFIG.SHEETS.VEHICLES);
    var now   = new Date();

    if (!data['Plate Number']) return { success: false, message: 'Plate Number is required.' };
    if (!data['Vehicle Type'] && !data['Vehicle ID']) return { success: false, message: 'Vehicle Type is required.' };
    if (!data['Status']) data['Status'] = 'Active';

    var existing  = getVehicles();
    for (var i = 0; i < existing.length; i++) {
      if (String(existing[i]['Plate Number']).toLowerCase() === String(data['Plate Number']).toLowerCase() &&
          String(existing[i]['Vehicle ID']) !== String(data['Vehicle ID'])) {
        return { success: false, message: 'Plate Number already exists.' };
      }
    }

    var allData = sheet.getDataRange().getValues();
    var headers = allData[0];

    if (data['Vehicle ID']) {
      var rowIdx = -1;
      for (var r = 1; r < allData.length; r++) {
        if (String(allData[r][0]) === String(data['Vehicle ID'])) { rowIdx = r; break; }
      }
      if (rowIdx === -1) return { success: false, message: 'Vehicle not found.' };
      for (var h = 0; h < headers.length; h++) {
        if (data[headers[h]] !== undefined && headers[h] !== 'Vehicle ID' && headers[h] !== 'Created At') {
          sheet.getRange(rowIdx + 1, h + 1).setValue(data[headers[h]]);
        }
      }
      var ui = headers.indexOf('Updated At');
      if (ui > -1) sheet.getRange(rowIdx + 1, ui + 1).setValue(now);
      return { success: true, message: 'Vehicle updated successfully.' };
    } else {
      var vehicleId = generateId('V', CONFIG.SHEETS.VEHICLES);
      var row = [];
      for (var k = 0; k < headers.length; k++) {
        if (headers[k] === 'Vehicle ID')  row.push(vehicleId);
        else if (headers[k] === 'Created At') row.push(now);
        else if (headers[k] === 'Updated At') row.push(now);
        else row.push(data[headers[k]] !== undefined ? data[headers[k]] : '');
      }
      sheet.appendRow(row);
      return { success: true, message: 'Vehicle created successfully.', vehicleId: vehicleId };
    }
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ── TRIPS ────────────────────────────────────────────────────
function getTrips() {
  return getSheetData(CONFIG.SHEETS.TRIPS);
}

function saveTrip(data) {
  try {
    var sheet  = getSheet(CONFIG.SHEETS.TRIPS);
    var now    = new Date();
    var status = data['Status'] || 'Draft';

    if (status === 'Submitted') {
      var reqFields = ['Requestor Employee ID','Trip Type','Purpose','From Location',
                       'To Location','Planned Start DateTime','Planned End DateTime',
                       'Vehicle ID','Driver Employee ID'];
      for (var i = 0; i < reqFields.length; i++) {
        if (!data[reqFields[i]]) return { success: false, message: reqFields[i] + ' is required before submitting.' };
      }
    }
    if (status === 'Rejected' && !data['Rejection Reason'])
      return { success: false, message: 'Rejection Reason is required.' };
    if (status === 'Cancelled' && !data['Cancel Reason'])
      return { success: false, message: 'Cancel Reason is required.' };
    if (status === 'Completed') {
      if (!data['Actual Start DateTime']) return { success: false, message: 'Actual Start DateTime is required.' };
      if (!data['Actual End DateTime'])   return { success: false, message: 'Actual End DateTime is required.' };
      if (!data['Start Mileage'])         return { success: false, message: 'Start Mileage is required.' };
      if (!data['End Mileage'])           return { success: false, message: 'End Mileage is required.' };
      if (Number(data['End Mileage']) < Number(data['Start Mileage']))
        return { success: false, message: 'End Mileage cannot be less than Start Mileage.' };
      data['Distance Travelled'] = Number(data['End Mileage']) - Number(data['Start Mileage']);
    }

    if (data['Vehicle ID']) {
      var v = getVehicleById(data['Vehicle ID']);
      if (v) data['Plate Number'] = v['Plate Number'];
    }
    if (data['Requestor Employee ID']) {
      var req = getEmployeeById(data['Requestor Employee ID']);
      if (req) data['Requestor Name'] = req['Employee Name'];
    }
    if (data['Driver Employee ID']) {
      var drv = getEmployeeById(data['Driver Employee ID']);
      if (drv) data['Driver Name'] = drv['Employee Name'];
    }

    var allData = sheet.getDataRange().getValues();
    var headers = allData[0];

    if (data['Trip ID']) {
      var rowIdx = -1;
      for (var r = 1; r < allData.length; r++) {
        if (String(allData[r][0]) === String(data['Trip ID'])) { rowIdx = r; break; }
      }
      if (rowIdx === -1) return { success: false, message: 'Trip not found.' };

      var curr = allData[rowIdx][headers.indexOf('Status')];
      if (!_isValidTransition(curr, status))
        return { success: false, message: 'Cannot change status from ' + curr + ' to ' + status + '.' };

      for (var h = 0; h < headers.length; h++) {
        if (data[headers[h]] !== undefined && headers[h] !== 'Trip ID' && headers[h] !== 'Request Date') {
          sheet.getRange(rowIdx + 1, h + 1).setValue(data[headers[h]]);
        }
      }
      var updAt  = headers.indexOf('Updated At');
      var updBy  = headers.indexOf('Updated By');
      if (updAt > -1) sheet.getRange(rowIdx + 1, updAt + 1).setValue(now);
      if (updBy > -1) sheet.getRange(rowIdx + 1, updBy + 1).setValue(data['Updated By'] || '');
      if (status === 'Approved' || status === 'Rejected') {
        var appIdx = headers.indexOf('Approval/Rejection Date');
        if (appIdx > -1) sheet.getRange(rowIdx + 1, appIdx + 1).setValue(now);
      }
      return { success: true, message: 'Trip ' + status + ' successfully.' };
    } else {
      var tripId = generateId('T', CONFIG.SHEETS.TRIPS);
      var row    = [];
      for (var k = 0; k < headers.length; k++) {
        if      (headers[k] === 'Trip ID')       row.push(tripId);
        else if (headers[k] === 'Request Date')  row.push(now);
        else if (headers[k] === 'Status')        row.push('Draft');
        else if (headers[k] === 'Updated At')    row.push(now);
        else if (headers[k] === 'Updated By')    row.push(data['Updated By'] || '');
        else row.push(data[headers[k]] !== undefined ? data[headers[k]] : '');
      }
      sheet.appendRow(row);
      return { success: true, message: 'Trip created successfully.', tripId: tripId };
    }
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function _isValidTransition(from, to) {
  var allowed = {
    'Draft':     ['Draft','Submitted','Cancelled'],
    'Submitted': ['Submitted','Approved','Rejected','Cancelled'],
    'Approved':  ['Approved','Completed','Cancelled'],
    'Rejected':  ['Rejected'],
    'Cancelled': ['Cancelled'],
    'Completed': ['Completed']
  };
  var list = allowed[from] || [];
  for (var i = 0; i < list.length; i++) { if (list[i] === to) return true; }
  return false;
}

// ── RENEWAL ALERTS ───────────────────────────────────────────
function refreshRenewalAlerts() {
  try {
    var sheet    = getSheet(CONFIG.SHEETS.RENEWAL_ALERTS);
    var vehicles = getVehicles();
    var now      = new Date(); now.setHours(0,0,0,0);
    var lastRow  = sheet.getLastRow();
    if (lastRow > 1) sheet.getRange(2,1,lastRow-1,6).clearContent();

    var rows = [];
    for (var i = 0; i < vehicles.length; i++) {
      var v    = vehicles[i];
      var docs = ['Insurance','LTO'];
      for (var d = 0; d < docs.length; d++) {
        var expVal = v[docs[d] + ' Expiry Date'];
        if (!expVal) continue;
        var expiry   = new Date(expVal); expiry.setHours(0,0,0,0);
        var daysLeft = Math.round((expiry - now) / 86400000);
        var alertSt  = daysLeft < 0 ? 'Expired' : daysLeft <= 30 ? 'Due in 30 Days' : 'OK';
        rows.push([v['Vehicle ID'], v['Plate Number'], docs[d] + ' Expiry', expVal, daysLeft, alertSt]);
      }
    }
    if (rows.length > 0) sheet.getRange(2,1,rows.length,6).setValues(rows);
    return { success: true, data: rows };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function dailyRenewalCheck() { refreshRenewalAlerts(); }

// ── REPORTS ──────────────────────────────────────────────────
function getReportTripsByVehicle() {
  var trips = getTrips(); var map = {};
  for (var i = 0; i < trips.length; i++) {
    if (trips[i]['Status'] !== 'Completed') continue;
    var k = trips[i]['Vehicle ID'] || 'Unknown';
    if (!map[k]) map[k] = { vehicleId: k, plateNumber: trips[i]['Plate Number'] || '', tripCount: 0, totalMileage: 0 };
    map[k].tripCount++;
    map[k].totalMileage += Number(trips[i]['Distance Travelled']) || 0;
  }
  var r = []; for (var k in map) r.push(map[k]); return r;
}

function getReportTripsByDriver() {
  var trips = getTrips(); var map = {};
  for (var i = 0; i < trips.length; i++) {
    if (trips[i]['Status'] !== 'Completed') continue;
    var k = trips[i]['Driver Employee ID'] || 'Unknown';
    if (!map[k]) map[k] = { driverId: k, driverName: trips[i]['Driver Name'] || '', tripCount: 0, totalMileage: 0 };
    map[k].tripCount++;
    map[k].totalMileage += Number(trips[i]['Distance Travelled']) || 0;
  }
  var r = []; for (var k in map) r.push(map[k]); return r;
}

function getReportTripsByType() {
  var trips = getTrips(); var map = {};
  for (var i = 0; i < trips.length; i++) {
    var k = trips[i]['Trip Type'] || 'Unknown';
    if (!map[k]) map[k] = { tripType: k, tripCount: 0 };
    map[k].tripCount++;
  }
  var r = []; for (var k in map) r.push(map[k]); return r;
}

function getReportMileageSummary() {
  var vehicles = getVehicles();
  var trips    = getTrips();
  var result   = [];
  for (var i = 0; i < vehicles.length; i++) {
    var v = vehicles[i]; var totalKm = 0; var lastTrip = null;
    for (var j = 0; j < trips.length; j++) {
      if (trips[j]['Status'] === 'Completed' && String(trips[j]['Vehicle ID']) === String(v['Vehicle ID'])) {
        totalKm += Number(trips[j]['Distance Travelled']) || 0;
        if (!lastTrip || new Date(trips[j]['Actual End DateTime']) > new Date(lastTrip['Actual End DateTime']))
          lastTrip = trips[j];
      }
    }
    result.push({
      vehicleId:        v['Vehicle ID'],
      plateNumber:      v['Plate Number'],
      beginningMileage: v['Beginning Mileage'] || 0,
      latestEndMileage: lastTrip ? lastTrip['End Mileage'] : (v['Beginning Mileage'] || 0),
      totalRecorded:    totalKm
    });
  }
  return result;
}

function getReportExpiringDocuments() {
  refreshRenewalAlerts();
  var data = getSheetData(CONFIG.SHEETS.RENEWAL_ALERTS);
  var r    = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i]['Alert Status'] === 'Expired' || data[i]['Alert Status'] === 'Due in 30 Days') r.push(data[i]);
  }
  return r;
}

// ── BULK DATA LOADER ─────────────────────────────────────────
function getAppData() {
  try {
    var ss = _getSS();
    if (!ss.getSheetByName(CONFIG.SHEETS.EMPLOYEES)) _initializeSpreadsheet(ss);
    if (!ss.getSheetByName(CONFIG.SHEETS.SETTINGS))  _setupSettingsSheet(ss);
    if (!ss.getSheetByName(CONFIG.SHEETS.VEHICLES))  _setupVehiclesSheet(ss);
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
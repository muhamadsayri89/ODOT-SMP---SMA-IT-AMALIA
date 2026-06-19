// ============================================================================
// APLIKASI SMP-SMA IT AMALIA - Google Apps Script
// Main Backend Logic
// ============================================================================

const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || 'YOUR_SPREADSHEET_ID';
const SHEET_NAMES = {
  USERS: 'Users',
  STUDENTS: 'Students',
  CLASSES: 'Classes',
  TEACHERS: 'Teachers',
  SCHEDULE: 'Schedule',
  ODOT: 'ODOT',
  TAHFIDZ: 'Tahfidz',
  MODULES: 'Modules',
  NOTIFICATIONS: 'Notifications',
  SETTINGS: 'Settings'
};

// ============================================================================
// INITIALIZATION & SPREADSHEET SETUP
// ============================================================================

function initializeSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // Create sheets if they don't exist
  createSheetIfNotExists(ss, SHEET_NAMES.USERS, [
    'UserId', 'Username', 'Password', 'NamaLengkap', 'Role', 'Email', 'Kelas', 'Mapel', 'Status', 'CreatedAt'
  ]);
  
  createSheetIfNotExists(ss, SHEET_NAMES.STUDENTS, [
    'StudentId', 'NamaLengkap', 'NISN', 'NIPD', 'TTL', 'JenisKelamin', 'Kelas', 'Target', 'CreatedAt'
  ]);
  
  createSheetIfNotExists(ss, SHEET_NAMES.CLASSES, [
    'ClassId', 'NamaKelas', 'Tingkat', 'TahunAjaran', 'Semester', 'Walas', 'CreatedAt'
  ]);
  
  createSheetIfNotExists(ss, SHEET_NAMES.TEACHERS, [
    'TeacherId', 'NamaGuru', 'UserId', 'Mengajar', 'Mapel', 'Email', 'CreatedAt'
  ]);
  
  createSheetIfNotExists(ss, SHEET_NAMES.SCHEDULE, [
    'ScheduleId', 'Mapel', 'Hari', 'Kelas', 'Jam', 'Guru', 'GuruId', 'CreatedAt'
  ]);
  
  createSheetIfNotExists(ss, SHEET_NAMES.ODOT, [
    'OdotId', 'Kelas', 'Tanggal', 'StudentId', 'Status', 'Jumlah', 'Keterangan', 'TotalHarian', 'CreatedAt'
  ]);
  
  createSheetIfNotExists(ss, SHEET_NAMES.TAHFIDZ, [
    'TahfidzId', 'StudentId', 'Tanggal', 'JumlahAyat', 'Target', 'Tipe', 'CreatedAt'
  ]);
  
  createSheetIfNotExists(ss, SHEET_NAMES.MODULES, [
    'ModuleId', 'Mapel', 'GuruId', 'GuruNama', 'Tanggal', 'FileUrl', 'FileName', 'CreatedAt'
  ]);
  
  createSheetIfNotExists(ss, SHEET_NAMES.NOTIFICATIONS, [
    'NotifId', 'UserId', 'Type', 'Title', 'Message', 'IsRead', 'CreatedAt'
  ]);
  
  createSheetIfNotExists(ss, SHEET_NAMES.SETTINGS, [
    'Setting', 'Value'
  ]);
}

function createSheetIfNotExists(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
  }
}

// ============================================================================
// AUTHENTICATION
// ============================================================================

function authenticateUser(username, password) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAMES.USERS);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  for (let row of data) {
    if (row[1] === username && row[2] === password) {
      return {
        success: true,
        userId: row[0],
        username: row[1],
        namaLengkap: row[3],
        role: row[4],
        email: row[5],
        kelas: row[6],
        mapel: row[7]
      };
    }
  }
  
  return { success: false, message: 'Username atau password salah' };
}

// ============================================================================
// USER MANAGEMENT (Admin Only)
// ============================================================================

function addUser(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  
  const userId = generateId('USR');
  const newRow = [
    userId,
    data.username,
    data.password,
    data.namaLengkap,
    data.role,
    data.email,
    data.kelas || '',
    data.mapel || '',
    'Active',
    new Date().toLocaleString('id-ID')
  ];
  
  sheet.appendRow(newRow);
  return { success: true, userId: userId };
}

function editUser(userId, data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  for (let i = 0; i < allData.length; i++) {
    if (allData[i][0] === userId) {
      const row = i + 2;
      sheet.getRange(row, 2).setValue(data.username);
      sheet.getRange(row, 3).setValue(data.password);
      sheet.getRange(row, 4).setValue(data.namaLengkap);
      sheet.getRange(row, 5).setValue(data.role);
      sheet.getRange(row, 6).setValue(data.email);
      sheet.getRange(row, 7).setValue(data.kelas || '');
      sheet.getRange(row, 8).setValue(data.mapel || '');
      return { success: true };
    }
  }
  
  return { success: false, message: 'User tidak ditemukan' };
}

function deleteUser(userId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  for (let i = allData.length - 1; i >= 0; i--) {
    if (allData[i][0] === userId) {
      sheet.deleteRow(i + 2);
      return { success: true };
    }
  }
  
  return { success: false };
}

function getAllUsers() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAMES.USERS);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  return data.map(row => ({
    userId: row[0],
    username: row[1],
    namaLengkap: row[3],
    role: row[4],
    email: row[5],
    kelas: row[6],
    mapel: row[7],
    status: row[8]
  }));
}

// ============================================================================
// CLASS MANAGEMENT (Admin Only)
// ============================================================================

function addClass(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.CLASSES);
  
  const classId = generateId('CLS');
  const newRow = [
    classId,
    data.namaKelas,
    data.tingkat,
    data.tahunAjaran,
    data.semester,
    data.walas || '',
    new Date().toLocaleString('id-ID')
  ];
  
  sheet.appendRow(newRow);
  return { success: true, classId: classId };
}

function editClass(classId, data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.CLASSES);
  const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  for (let i = 0; i < allData.length; i++) {
    if (allData[i][0] === classId) {
      const row = i + 2;
      sheet.getRange(row, 2).setValue(data.namaKelas);
      sheet.getRange(row, 3).setValue(data.tingkat);
      sheet.getRange(row, 4).setValue(data.tahunAjaran);
      sheet.getRange(row, 5).setValue(data.semester);
      sheet.getRange(row, 6).setValue(data.walas || '');
      return { success: true };
    }
  }
  
  return { success: false };
}

function deleteClass(classId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.CLASSES);
  const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  for (let i = allData.length - 1; i >= 0; i--) {
    if (allData[i][0] === classId) {
      sheet.deleteRow(i + 2);
      return { success: true };
    }
  }
  
  return { success: false };
}

function getAllClasses() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAMES.CLASSES);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  return data.map(row => ({
    classId: row[0],
    namaKelas: row[1],
    tingkat: row[2],
    tahunAjaran: row[3],
    semester: row[4],
    walas: row[5]
  }));
}

// ============================================================================
// STUDENT MANAGEMENT
// ============================================================================

function addStudent(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.STUDENTS);
  
  const studentId = generateId('STD');
  const newRow = [
    studentId,
    data.namaLengkap,
    data.nisn,
    data.nipd,
    data.ttl,
    data.jenisKelamin,
    data.kelas,
    data.target || 'Reguler',
    new Date().toLocaleString('id-ID')
  ];
  
  sheet.appendRow(newRow);
  return { success: true, studentId: studentId };
}

function editStudent(studentId, data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.STUDENTS);
  const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  for (let i = 0; i < allData.length; i++) {
    if (allData[i][0] === studentId) {
      const row = i + 2;
      sheet.getRange(row, 2).setValue(data.namaLengkap);
      sheet.getRange(row, 3).setValue(data.nisn);
      sheet.getRange(row, 4).setValue(data.nipd);
      sheet.getRange(row, 5).setValue(data.ttl);
      sheet.getRange(row, 6).setValue(data.jenisKelamin);
      sheet.getRange(row, 7).setValue(data.kelas);
      sheet.getRange(row, 8).setValue(data.target || 'Reguler');
      return { success: true };
    }
  }
  
  return { success: false };
}

function deleteStudent(studentId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.STUDENTS);
  const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  for (let i = allData.length - 1; i >= 0; i--) {
    if (allData[i][0] === studentId) {
      sheet.deleteRow(i + 2);
      return { success: true };
    }
  }
  
  return { success: false };
}

function getStudentsByClass(kelas) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAMES.STUDENTS);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  return data.filter(row => row[6] === kelas).map(row => ({
    studentId: row[0],
    namaLengkap: row[1],
    nisn: row[2],
    nipd: row[3],
    ttl: row[4],
    jenisKelamin: row[5],
    kelas: row[6],
    target: row[7]
  }));
}

function getAllStudents() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAMES.STUDENTS);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  return data.map(row => ({
    studentId: row[0],
    namaLengkap: row[1],
    nisn: row[2],
    nipd: row[3],
    ttl: row[4],
    jenisKelamin: row[5],
    kelas: row[6],
    target: row[7]
  }));
}

// ============================================================================
// SCHEDULE MANAGEMENT (Admin Only)
// ============================================================================

const MAPEL_LIST = ['PAIBP', 'PKN', 'B. Indo', 'MTK', 'IPA', 'IPS', 'B. Ing', 'PJOK', 'TIK', 'SBdP', 'B. Sunda', 'B. Arab', 'Al-Qur\'an', 'Siroh'];
const JAM_LIST = [
  { jam: 1, waktu: '09.00-09.40' },
  { jam: 2, waktu: '09.40-10.40' },
  { jam: 3, waktu: '10.40-11.20' },
  { jam: 4, waktu: '11.20-12.00' },
  { jam: 5, waktu: '13.30-14.10' },
  { jam: 6, waktu: '14.10-14.50' }
];

function addSchedule(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.SCHEDULE);
  
  const scheduleId = generateId('SCH');
  const newRow = [
    scheduleId,
    data.mapel,
    data.hari,
    data.kelas,
    data.jam,
    data.guru,
    data.guruId || '',
    new Date().toLocaleString('id-ID')
  ];
  
  sheet.appendRow(newRow);
  return { success: true, scheduleId: scheduleId };
}

function editSchedule(scheduleId, data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.SCHEDULE);
  const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  for (let i = 0; i < allData.length; i++) {
    if (allData[i][0] === scheduleId) {
      const row = i + 2;
      sheet.getRange(row, 2).setValue(data.mapel);
      sheet.getRange(row, 3).setValue(data.hari);
      sheet.getRange(row, 4).setValue(data.kelas);
      sheet.getRange(row, 5).setValue(data.jam);
      sheet.getRange(row, 6).setValue(data.guru);
      sheet.getRange(row, 7).setValue(data.guruId || '');
      return { success: true };
    }
  }
  
  return { success: false };
}

function deleteSchedule(scheduleId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.SCHEDULE);
  const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  for (let i = allData.length - 1; i >= 0; i--) {
    if (allData[i][0] === scheduleId) {
      sheet.deleteRow(i + 2);
      return { success: true };
    }
  }
  
  return { success: false };
}

function getScheduleByClass(kelas) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAMES.SCHEDULE);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  return data.filter(row => row[3] === kelas).map(row => ({
    scheduleId: row[0],
    mapel: row[1],
    hari: row[2],
    kelas: row[3],
    jam: row[4],
    guru: row[5]
  }));
}

// ============================================================================
// ODOT MANAGEMENT
// ============================================================================

function addODOT(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.ODOT);
  
  const odotId = generateId('ODT');
  const newRow = [
    odotId,
    data.kelas,
    new Date().toLocaleDateString('id-ID'),
    data.studentId || '',
    data.status || 'Bayar',
    data.jumlah || 1000,
    data.keterangan || '',
    0,
    new Date().toLocaleString('id-ID')
  ];
  
  sheet.appendRow(newRow);
  return { success: true, odotId: odotId };
}

function getODOTByClass(kelas, startDate, endDate) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAMES.ODOT);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  const filtered = data.filter(row => {
    const rowDate = new Date(row[2]);
    const start = new Date(startDate);
    const end = new Date(endDate);
    return row[1] === kelas && rowDate >= start && rowDate <= end;
  });
  
  return {
    data: filtered.map(row => ({
      odotId: row[0],
      kelas: row[1],
      tanggal: row[2],
      studentId: row[3],
      status: row[4],
      jumlah: row[5],
      keterangan: row[6]
    })),
    total: filtered.reduce((sum, row) => sum + (row[4] === 'Bayar' ? row[5] : 0), 0)
  };
}

// ============================================================================
// TAHFIDZ MANAGEMENT
// ============================================================================

function addTahfidz(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.TAHFIDZ);
  
  const tahfidzId = generateId('THF');
  const newRow = [
    tahfidzId,
    data.studentId,
    new Date().toLocaleDateString('id-ID'),
    data.jumlahAyat || 0,
    data.target || 0,
    data.tipe || 'Reguler',
    new Date().toLocaleString('id-ID')
  ];
  
  sheet.appendRow(newRow);
  return { success: true, tahfidzId: tahfidzId };
}

function editTahfidz(tahfidzId, data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.TAHFIDZ);
  const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  for (let i = 0; i < allData.length; i++) {
    if (allData[i][0] === tahfidzId) {
      const row = i + 2;
      sheet.getRange(row, 4).setValue(data.jumlahAyat);
      sheet.getRange(row, 5).setValue(data.target);
      return { success: true };
    }
  }
  
  return { success: false };
}

function getTahfidzByStudent(studentId, startDate, endDate) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAMES.TAHFIDZ);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  const filtered = data.filter(row => {
    const rowDate = new Date(row[2]);
    const start = new Date(startDate);
    const end = new Date(endDate);
    return row[1] === studentId && rowDate >= start && rowDate <= end;
  });
  
  return filtered.map(row => ({
    tahfidzId: row[0],
    studentId: row[1],
    tanggal: row[2],
    jumlahAyat: row[3],
    target: row[4],
    tipe: row[5]
  }));
}

// ============================================================================
// MODULE MANAGEMENT
// ============================================================================

function addModule(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.MODULES);
  
  const moduleId = generateId('MOD');
  const newRow = [
    moduleId,
    data.mapel,
    data.guruId,
    data.guruNama,
    new Date().toLocaleDateString('id-ID'),
    data.fileUrl || '',
    data.fileName || '',
    new Date().toLocaleString('id-ID')
  ];
  
  sheet.appendRow(newRow);
  return { success: true, moduleId: moduleId };
}

function getModulesByMapel(mapel) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAMES.MODULES);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  return data.filter(row => row[1] === mapel).map(row => ({
    moduleId: row[0],
    mapel: row[1],
    guruId: row[2],
    guruNama: row[3],
    tanggal: row[4],
    fileUrl: row[5],
    fileName: row[6]
  }));
}

function getAllModules() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAMES.MODULES);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  return data.map(row => ({
    moduleId: row[0],
    mapel: row[1],
    guruId: row[2],
    guruNama: row[3],
    tanggal: row[4],
    fileUrl: row[5],
    fileName: row[6]
  }));
}

// ============================================================================
// NOTIFICATION MANAGEMENT
// ============================================================================

function addNotification(userId, type, title, message) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.NOTIFICATIONS);
  
  const notifId = generateId('NOT');
  const newRow = [
    notifId,
    userId,
    type,
    title,
    message,
    false,
    new Date().toLocaleString('id-ID')
  ];
  
  sheet.appendRow(newRow);
  return { success: true, notifId: notifId };
}

function getNotificationsByUser(userId) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAMES.NOTIFICATIONS);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  return data.filter(row => row[1] === userId).map(row => ({
    notifId: row[0],
    userId: row[1],
    type: row[2],
    title: row[3],
    message: row[4],
    isRead: row[5],
    createdAt: row[6]
  }));
}

function markNotificationAsRead(notifId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.NOTIFICATIONS);
  const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  for (let i = 0; i < allData.length; i++) {
    if (allData[i][0] === notifId) {
      sheet.getRange(i + 2, 6).setValue(true);
      return { success: true };
    }
  }
  
  return { success: false };
}

// ============================================================================
// SETTINGS MANAGEMENT
// ============================================================================

function setSetting(key, value) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 2, 2).setValue(value);
      return { success: true };
    }
  }
  
  sheet.appendRow([key, value]);
  return { success: true };
}

function getSetting(key) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAMES.SETTINGS);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  for (let row of data) {
    if (row[0] === key) {
      return row[1];
    }
  }
  
  return null;
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

function generateId(prefix) {
  return prefix + '_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
}

function doGet() {
  return HtmlService.createHtmlOutput(getHtml()).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getHtml() {
  return HtmlService.createTemplateFromFile('HTML').evaluate().getContent();
}

// ============================================================================
// EXPOSED FUNCTIONS FOR FRONTEND (Callable via google.script.run)
// ============================================================================

// These are called from the HTML frontend
function initApp() {
  initializeSheets();
  return { success: true, message: 'Aplikasi sudah diinisialisasi' };
}

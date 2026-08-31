function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Kontrol Perkembangan 7 KAIH - SMPN 1 Wangi-Wangi')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetUsers = ss.getSheetByName('Users');
  let sheetData = ss.getSheetByName('DataKaih');
  let sheetSettings = ss.getSheetByName('Settings');

  if (!sheetUsers) {
    sheetUsers = ss.insertSheet('Users');
    sheetUsers.appendRow(['Role', 'Nama', 'NIP_NIS', 'Kelas', 'GuruWali', 'Password']);
  }
  if (!sheetData) {
    sheetData = ss.insertSheet('DataKaih');
    sheetData.appendRow(['Nama', 'Kelas', 'Bulan', 'Tahun', 'Tgl', 'BangunPagi', 'Subuh', 'Dzuhur', 'Ashar', 'Maghrib', 'Isya', 'BacaAlQuran', 'Berolahraga', 'MakanSehat', 'GemarBelajar', 'Bermasyarakat', 'TidurCepat', 'Ket']);
  }
  if (!sheetSettings) {
    sheetSettings = ss.insertSheet('Settings');
    sheetSettings.appendRow(['NamaKepsek', 'NipKepsek']);
  }
  return ss;
}

function getDaftarUserByRole(role) {
  getDatabase();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getDisplayValues();
  const listUser = [];
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === role) {
      listUser.push({
        nama: String(data[i][1]),
        nipNis: String(data[i][2]),
        kelas: String(data[i][3]),
        guruWali: String(data[i][4])
      });
    }
  }
  return listUser;
}

function getDaftarGuru() {
  return getDaftarUserByRole('Guru');
}

function registerUser(userData) {
  getDatabase();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (sheet.getLastRow() >= 2) {
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getDisplayValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][1]).trim().toLowerCase() === String(userData.nama).trim().toLowerCase() && data[i][0] === userData.role) {
        return { success: false, message: 'Nama sudah terdaftar sebagai ' + userData.role + '!' };
      }
    }
  }
  
  sheet.appendRow([
    userData.role,
    String(userData.nama).trim(),
    String(userData.nipNis).trim(),
    userData.kelas ? String(userData.kelas).trim() : '',
    userData.guruWali ? String(userData.guruWali).trim() : '',
    String(userData.password)
  ]);
  return { success: true, message: 'Registrasi berhasil! Silakan login.' };
}

function loginUserByName(role, nama, password) {
  getDatabase();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (sheet.getLastRow() < 2) return { success: false, message: 'Belum ada pengguna terdaftar!' };
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getDisplayValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === role && String(data[i][1]).trim() === String(nama).trim() && String(data[i][5]) === String(password)) {
      return {
        success: true,
        user: { role: data[i][0], nama: String(data[i][1]), nipNis: String(data[i][2]), kelas: String(data[i][3]), guruWali: String(data[i][4]) }
      };
    }
  }
  return { success: false, message: 'Password salah!' };
}

function simpanDataKaih(formData) {
  getDatabase();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DataKaih');
  
  sheet.appendRow([
    String(formData.nama).trim(),
    String(formData.kelas).trim(),
    String(formData.bulan).trim(),
    String(formData.tahun).trim(),
    String(formData.tgl),
    String(formData.bangunPagi),
    formData.subuh ? 'Ya' : 'Tidak',
    formData.dzuhur ? 'Ya' : 'Tidak',
    formData.ashar ? 'Ya' : 'Tidak',
    formData.maghrib ? 'Ya' : 'Tidak',
    formData.isya ? 'Ya' : 'Tidak',
    String(formData.bacaAlQuran),
    formData.berolahraga ? 'Ya' : 'Tidak',
    formData.makanSehat ? 'Ya' : 'Tidak',
    formData.gemarBelajar ? 'Ya' : 'Tidak',
    formData.bermasyarakat ? 'Ya' : 'Tidak',
    formData.tidurCepat ? 'Ya' : 'Tidak',
    String(formData.ket)
  ]);
  return { success: true, message: 'Data harian berhasil disimpan!' };
}

function getSettings() {
  getDatabase();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  if (sheet.getLastRow() < 2) return { namaKepsek: '', nipKepsek: '' };
  
  const data = sheet.getRange(2, 1, 1, 2).getDisplayValues()[0];
  return { 
    namaKepsek: String(data[0] || '').trim(), 
    nipKepsek: String(data[1] || '').trim() 
  };
}

function simpanSettings(settings) {
  getDatabase();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  sheet.clear();
  sheet.appendRow(['NamaKepsek', 'NipKepsek']);
  sheet.appendRow([String(settings.namaKepsek).trim(), String(settings.nipKepsek).trim()]);
  return { success: true, message: 'Pengaturan Kepala Sekolah berhasil disimpan!' };
}

function getDaftarMuridByGuru(namaGuru) {
  getDatabase();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (sheet.getLastRow() < 2) return [];
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getDisplayValues();
  const listMurid = [];
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === 'Murid' && String(data[i][4]).trim() === String(namaGuru).trim()) {
      listMurid.push({ nama: String(data[i][1]), kelas: String(data[i][3]) });
    }
  }
  return listMurid;
}

// Fungsi serbaguna untuk mengambil laporan berdasarkan nama, bulan, dan tahun
function getLaporanMurid(nama, bulan, tahun) {
  getDatabase();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DataKaih');
  if (sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 18).getDisplayValues();
  const result = [];
  
  const targetNama = String(nama).trim().toLowerCase();
  const targetBulan = String(bulan).trim().toLowerCase();
  const targetTahun = String(tahun).trim();

  for (let i = 0; i < data.length; i++) {
    const rowNama = String(data[i][0]).trim().toLowerCase();
    const rowBulan = String(data[i][2]).trim().toLowerCase();
    const rowTahun = String(data[i][3]).trim();

    if (rowNama === targetNama && rowBulan === targetBulan && (rowTahun === targetTahun || rowTahun === '')) {
      result.push({
        tgl: data[i][4],
        bangunPagi: data[i][5],
        subuh: data[i][6],
        dzuhur: data[i][7],
        ashar: data[i][8],
        maghrib: data[i][9],
        isya: data[i][10],
        bacaAlQuran: data[i][11],
        berolahraga: data[i][12],
        makanSehat: data[i][13],
        gemarBelajar: data[i][14],
        bermasyarakat: data[i][15],
        tidurCepat: data[i][16],
        ket: data[i][17]
      });
    }
  }
  return result;
}
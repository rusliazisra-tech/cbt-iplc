/* ======================================================================
   KODE CBT ULTIMATE v3.0: ADMIN PANEL + MANAJEMEN DATA + TOOLS
   ====================================================================== */

const SPREADSHEET_ID = '1jitT60VBe-mDe9_LNOsTv_cCJTkPtRmiWV1o90JTgp8'; 

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('CBT - ANBK SYSTEM')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* --- MENU ALAT BANTU SPREADSHEET (TETAP ADA) --- */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🟢 ALAT BANTU GAMBAR')
      .addItem('1. Ambil Link Folder (Massal)', 'getDriveLinksThumbnail')
      .addSeparator()
      .addItem('2. Konversi Link (Kolom G)', 'convertSingleLinkRaw')
      .addItem('3. Konversi HTML (Kolom C)', 'convertSingleLinkHTML')
      .addToUi();

  ui.createMenu('📝 ALAT BANTU SOAL')
      .addItem('📅 Tabel HTML', 'insertTableTemplate')
      .addSeparator()
      .addItem('✅ Opsi: Benar-Salah', 'insertOptionBS')
      .addItem('👍 Opsi: Setuju-Tidak', 'insertOptionAgree')
      .addItem('🔠 Opsi: A-E Kosong', 'insertOptionABCDE')
      .addToUi();

  ui.createMenu('Vx RUMUS LATEX')
      .addItem('➕ Bungkus ($...$)', 'wrapInLatex')
      .addSeparator()
      .addItem('½ Pecahan', 'insertFraction')
      .addItem('x² Pangkat', 'insertPower')
      .addItem('√ Akar', 'insertRoot')
      .addToUi();
}

/* --- FUNGSI ADMIN API (MANAJEMEN PESERTA & SOAL) --- */

// 1. Ambil Data Peserta untuk Tabel Admin
function adminGetUsers() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  // Skip header, ambil username, password, nama, role, prodi
  return data.slice(1).map(r => ({ u: r[0], p: r[1], n: r[2], role: r[3], prodi: r[4] }));
}

// 2. Tambah Peserta Baru
function adminAddUser(u, p, n, prodi) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Users');
    // Cek duplikat
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++) {
      if(String(data[i][0]) == String(u)) return {status:'error', message:'Username sudah ada!'};
    }
    sheet.appendRow([u, p, n, 'siswa', prodi]);
    return {status:'success'};
  } catch(e) { return {status:'error', message:e.message}; }
}

// 3. Hapus Peserta
function adminDeleteUser(username) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  for(let i=0; i<data.length; i++) {
    if(String(data[i][0]) == String(username)) {
      sheet.deleteRow(i+1);
      return {status:'success'};
    }
  }
  return {status:'error', message:'User tidak ditemukan'};
}

// 4. Reset Login (Hapus dari Monitoring & Results agar bisa ujian lagi)
function adminResetLogin(username) {
  return resetLoginSiswa(username); // Memanggil fungsi yg sudah ada
}

// 5. Ambil Daftar Soal Ringkas
function adminGetSoalList() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Soal');
  if (sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow()-1, 6).getValues();
  // Ambil ID, Tipe, Mapel, Soal(potongan)
  return data.map((r, i) => ({ 
    row: i+2, id: r[0], tipe: r[1], soal: String(r[3]).substring(0, 50) + "...", mapel: r[5] 
  }));
}

// 6. Hapus Soal
function adminDeleteSoal(rowCbt) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Soal');
    sheet.deleteRow(parseInt(rowCbt));
    return {status:'success'};
  } catch(e) { return {status:'error', message:e.message}; }
}

/* --- LOGIKA CBT UTAMA (LOGIN & UJIAN) --- */
// (Bagian ini sama dengan sebelumnya, hanya dirapikan)
function getConfigData() { /* ...Kode Config Lama... */
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Config');
    if (!sheet) return { mapel: ["UJIAN UMUM"], prodi: ["UMUM"] };
    const lastRow = Math.max(sheet.getLastRow(), 2);
    const dataMapel = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(String);
    const rawProdi = sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat().filter(String);
    let uniqueProdi = new Set();
    rawProdi.forEach(p => { p.split(',').forEach(item => { let cleanItem = item.trim().toUpperCase(); if (cleanItem !== "SEMUA") uniqueProdi.add(cleanItem); }); });
    let listProdi = uniqueProdi.size > 0 ? Array.from(uniqueProdi).sort() : ["UMUM"];
    return { mapel: dataMapel.length ? dataMapel : ["UJIAN UMUM"], prodi: listProdi };
  } catch (e) { return { mapel: ["UJIAN UMUM"], prodi: ["UMUM"] }; }
}

function loginUser(username, password, mapel) { 
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetUsers = ss.getSheetByName('Users');
    const dataUsers = sheetUsers.getDataRange().getValues();
    let userFound = null; let userProdiDB = ""; 
    
    // Cek Login
    for (let i = 1; i < dataUsers.length; i++) {
      if (String(dataUsers[i][0]).toLowerCase() == String(username).toLowerCase() && String(dataUsers[i][1]) == String(password)) {
        userFound = { username: dataUsers[i][0], nama: dataUsers[i][2], role: dataUsers[i][3] };
        userProdiDB = String(dataUsers[i][4] || "UMUM").toUpperCase().trim();
        userFound.prodi = userProdiDB; break;
      }
    }
    if (!userFound) return { status: 'error', message: 'Username atau Password salah.' };
    
    // JIKA ADMIN, LANGSUNG LOLOS
    if (String(userFound.role).toLowerCase() === 'admin') {
      return { status: 'success', data: userFound, isAdmin: true };
    }

    // Validasi Mapel (Strict Mode)
    const sheetConfig = ss.getSheetByName('Config');
    if (sheetConfig) {
      const dataConfig = sheetConfig.getDataRange().getValues();
      let mapelFound = false; let isProdiAllowed = false;
      for (let c = 1; c < dataConfig.length; c++) {
        let configMapel = String(dataConfig[c][0]).trim(); 
        let allowedProdis = String(dataConfig[c][1] || "SEMUA").toUpperCase(); 
        if (configMapel === String(mapel).trim()) {
          mapelFound = true;
          let prodiList = allowedProdis.split(',').map(p => p.trim());
          if (prodiList.includes("SEMUA") || prodiList.includes(userProdiDB)) { isProdiAllowed = true; }
          break; 
        }
      }
      if (!mapelFound) return { status: 'error', message: 'Konfigurasi Mapel tidak ditemukan.' };
      if (!isProdiAllowed) { return { status: 'error', message: `AKSES DITOLAK! Anda jurusan ${userProdiDB}.` }; }
    }
    
    // Cek Status Selesai
    const sheetResults = ss.getSheetByName('Results');
    if (sheetResults && sheetResults.getLastRow() > 1) {
      const dataResults = sheetResults.getDataRange().getValues();
      for (let j = 1; j < dataResults.length; j++) {
        if (String(dataResults[j][0]) == String(username) && String(dataResults[j][6]) == String(mapel)) { 
          return { status: 'error', message: 'Anda sudah menyelesaikan ujian ini.' }; 
        }
      }
    }
    
    // Cek Resume
    let savedAnswers = {};
    const sheetMon = ss.getSheetByName('Monitoring');
    if (sheetMon && sheetMon.getLastRow() > 1) {
      const dataMon = sheetMon.getDataRange().getValues();
      for (let k = dataMon.length - 1; k >= 1; k--) {
        if (String(dataMon[k][0]) == String(username) && String(dataMon[k][8]) == String(mapel)) { try { savedAnswers = JSON.parse(dataMon[k][6]); } catch(e){} break; }
      }
    }
    return { status: 'success', data: userFound, savedAnswers: savedAnswers, isAdmin: false };

  } catch (e) { return { status: 'error', message: 'Error DB: ' + e.message }; }
}

// --- FUNGSI PENDUKUNG LAINNYA (SAMA SEPERTI SEBELUMNYA) ---
function getSoal(targetMapel) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Soal');
    if (sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues(); 
    if (!targetMapel) return [];
    const filteredData = data.filter(row => String(row[5]).trim() === String(targetMapel).trim());
    return filteredData.map(row => {
      let parsedOpsi = []; try { parsedOpsi = JSON.parse(row[4]); } catch (e) { }
      return { id: row[0], tipe: row[1], wacana: row[2], soal: row[3], opsi: parsedOpsi, gambar: row[6] };
    });
  } catch (e) { return []; }
}
function syncProgress(user, progress, total, status, currentAns, prodi, mapel) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Monitoring');
    if (!sheet) { sheet = ss.insertSheet('Monitoring'); sheet.appendRow(['Username','Nama','Terjawab','Total','Status','Update','Draft_Jawaban','Prodi','Mapel']); }
    const data = sheet.getDataRange().getValues();
    let found = false; let jsonAns = JSON.stringify(currentAns);
    for(let i=1; i<data.length; i++) {
      if(String(data[i][0]) == String(user.username)) {
        sheet.getRange(i+1, 3, 1, 7).setValues([[progress, total, status, new Date(), jsonAns, prodi, mapel]]); found = true; break;
      }
    }
    if(!found && total > 0) sheet.appendRow([user.username, user.nama, progress, total, status, new Date(), jsonAns, prodi, mapel]);
  } catch(e) {}
}
function submitUjian(jawaban, user, mapel, prodi) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetSoal = ss.getSheetByName('Soal');
    const dataSoal = sheetSoal.getDataRange().getValues();

    let score = 0;

    const mapelSoal = dataSoal
      .slice(1)
      .filter(r => String(r[5]).trim() === String(mapel).trim());

    mapelSoal.forEach(row => {
      const idSoal = row[0];
      const tipe = String(row[1]).toLowerCase().trim();
      const kunci = String(row[7]).trim().toUpperCase();
      const jawab = (jawaban[idSoal] || '').toUpperCase().trim();

      if (!kunci) return;

      if (tipe === 'ganda') {
        const k = kunci.split(',').map(x => x.trim()).sort().join(',');
        const j = jawab.split(',').map(x => x.trim()).sort().join(',');
        if (k === j) score++;
      } else {
        if (jawab === kunci) score++;
      }
    });

    const total = mapelSoal.length;
    const finalScore = total > 0 ? ((score / total) * 100).toFixed(2) : 0;

    let sheet = ss.getSheetByName('Results');
    if (!sheet) {
      sheet = ss.insertSheet('Results');
      sheet.appendRow(['Username', 'Nama', 'Waktu', 'Jawaban', 'Skor', 'Prodi', 'Mapel']);
    }

    sheet.appendRow([
      user.username,
      user.nama,
      new Date(),
      JSON.stringify(jawaban),
      finalScore,
      prodi,
      mapel
    ]);

    // 🔒 JANGAN BIARKAN syncProgress GAGALKAN RETURN
    try {
      syncProgress(user, Object.keys(jawaban).length, 0, 'Selesai', {}, prodi, mapel);
    } catch (err) {
      Logger.log('syncProgress error: ' + err.message);
    }

    // ✅ WAJIB RETURN
    return { status: 'success' };

  } catch (err) {
    return { status: 'error', message: err.message };
  }
}

function getReportData(filterMapel, filterProdi) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Results'); if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
    let filtered = data.map(r => ({ username: r[0], nama: r[1], waktu: new Date(r[2]).toLocaleString(), skor: r[4], prodi: r[5], mapel: r[6] }));
    if (filterMapel && filterMapel !== "") filtered = filtered.filter(item => item.mapel === filterMapel);
    if (filterProdi && filterProdi !== "") filtered = filtered.filter(item => item.prodi === filterProdi);
    return filtered;
  } catch (e) { return []; }
}
function resetLoginSiswa(username) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Results');
    if (sheet) { const data = sheet.getDataRange().getValues(); for (let i = data.length - 1; i >= 1; i--) { if (String(data[i][0]) == String(username)) sheet.deleteRow(i + 1); } }
    const sheetMon = ss.getSheetByName('Monitoring');
    if (sheetMon) { const dataMon = sheetMon.getDataRange().getValues(); for (let j = dataMon.length - 1; j >= 1; j--) { if (String(dataMon[j][0]) == String(username)) sheetMon.deleteRow(j + 1); } }
    return { status: 'success', message: `Data ${username} berhasil di-reset.` };
  } catch (e) { return { status: 'error', message: e.message }; }
}
function getMonitoringData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); const sheet = ss.getSheetByName('Monitoring'); if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow()-1, 9).getValues(); const now = new Date();
    return data.filter(r => r[0] !== '').map(r => ({ username: r[0], nama: r[1], terjawab: r[2], total: r[3], status: (r[4]!=='Selesai' && (now - new Date(r[5]) > 120000)) ? 'Offline' : r[4], mapel: r[8] }));
  } catch (e) { return []; }
}
// Fungsi Pendukung Alat Bantu
function appendToCell(text) { var c = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell(); c.setValue((String(c.getValue()) + " " + text).trim()); }
function wrapInLatex() { var c = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell(); var v = String(c.getValue()).trim(); if(v && !v.startsWith("$")) c.setValue("$ "+v+" $"); }
function insertTableTemplate() { appendToCell('<table border="1" style="width:100%;border-collapse:collapse;text-align:center;"><tr style="background:#f0f0f0;"><th>J1</th><th>J2</th></tr><tr><td>A</td><td>B</td></tr></table>'); }
function insertOptionBS() { SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell().setValue('[{"id":"A","text":"Benar"},{"id":"B","text":"Salah"}]'); }
function insertOptionAgree() { SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell().setValue('[{"id":"A","text":"Setuju"},{"id":"B","text":"Tidak Setuju"}]'); }
function insertOptionABCDE() { SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell().setValue('[{"id":"A","text":"..."},{"id":"B","text":"..."},{"id":"C","text":"..."},{"id":"D","text":"..."},{"id":"E","text":"..."}]'); }
function insertFraction() { appendToCell("$\\frac{a}{b}$"); } function insertPower() { appendToCell("$x^{2}$"); } function insertRoot() { appendToCell("$\\sqrt{x}$"); }
function getDriveLinksThumbnail() { /* ...Kode Sama spt sblmnya... */ var ui = SpreadsheetApp.getUi(); var r = ui.prompt('ID Folder:', ui.ButtonSet.OK_CANCEL); if(r.getSelectedButton()==ui.Button.OK) { try { var f=DriveApp.getFolderById(r.getResponseText().trim()); var files=f.getFiles(); var d=[["NAMA","LINK"]]; while(files.hasNext()){var x=files.next(); if(x.getMimeType().includes("image")) d.push([x.getName(),"https://drive.google.com/thumbnail?id="+x.getId()+"&sz=w1000"]);} var s=SpreadsheetApp.getActiveSpreadsheet(); var ts=s.getSheetByName("HASIL_GAMBAR_BARU")||s.insertSheet("HASIL_GAMBAR_BARU"); ts.clear(); ts.getRange(1,1,d.length,2).setValues(d); ui.alert("Berhasil!"); } catch(e){ui.alert("Gagal: "+e.message);} } }
function convertSingleLinkRaw() { processLinkConversion(false); } function convertSingleLinkHTML() { processLinkConversion(true); }
function processLinkConversion(html) { var ui=SpreadsheetApp.getUi(); var r=ui.prompt('Link:', ui.ButtonSet.OK_CANCEL); if(r.getSelectedButton()==ui.Button.OK) { var m=r.getResponseText().match(/\/d\/([a-zA-Z0-9_-]+)/); if(m){ var l="https://drive.google.com/thumbnail?id="+m[1]+"&sz=w1000"; SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell().setValue(html ? "<b>Gbr:</b><br><img src=\""+l+"\" style=\"width:100%;max-width:500px;\">" : l); } else { ui.alert("Link Salah"); } } }

// Fungsi untuk memproses data menjadi CSV agar bisa diunduh
function downloadReportCSV(filterMapel, filterProdi) {
  try {
    const data = getReportData(filterMapel, filterProdi);
    if (data.length === 0) return { status: 'error', message: 'Tidak ada data untuk diunduh' };

    let csvContent = "Username,Nama,Waktu,Skor,Prodi,Mapel\n";
    data.forEach(r => {
      // Membersihkan koma agar tidak merusak format CSV
      let row = [
        `"${r.username}"`,
        `"${r.nama}"`,
        `"${r.waktu}"`,
        `"${r.skor}"`,
        `"${r.prodi}"`,
        `"${r.mapel}"`
      ];
      csvContent += row.join(",") + "\n";
    });

    return { status: 'success', csv: csvContent };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

// Fungsi untuk mengambil data statistik di Dashboard Admin
function getAdminStats() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const u = ss.getSheetByName('Users').getLastRow() - 1;
  const s = ss.getSheetByName('Soal').getLastRow() - 1;
  const h = ss.getSheetByName('Results') ? ss.getSheetByName('Results').getLastRow() - 1 : 0;
  return { users: u, soal: s, hasil: h };
}

// Fungsi untuk mengambil data soal khusus tampilan tabel admin
function getSoalDataForAdmin() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Soal');
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  return data.map(r => ({ id: r[0], tipe: r[1], soal: String(r[3]).substring(0, 50) + "..." }));
}

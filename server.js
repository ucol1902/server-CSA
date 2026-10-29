/**
 * SelesaiSurvei - Local Express Server dengan Database Excel (.xlsx)
 * Menjalankan dashboard CSA secara offline/lokal dengan database file Excel.
 * 
 * Dependensi yang dibutuhkan:
 *   npm install express xlsx
 * Jalankan:
 *   node server.js
 */

import express from 'express';
import xlsxModule from 'xlsx';
const XLSX = xlsxModule.default || xlsxModule;
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;
const DB_FILE = path.resolve(__dirname, 'database.xlsx');

// Middleware
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// 20 Master Units Default Wilayah Jambi
const INITIAL_MASTER_UNITS = [
  { id: '1', name: 'UP3 MUARA BUNGO', parent: 'UP3 MUARA BUNGO', target: 0 },
  { id: '2', name: 'ULP SAROLANGUN', parent: 'UP3 MUARA BUNGO', target: 0 },
  { id: '3', name: 'ULP SNGAI PNUH', parent: 'UP3 MUARA BUNGO', target: 0 },
  { id: '4', name: 'ULP MUARA TEBO', parent: 'UP3 MUARA BUNGO', target: 0 },
  { id: '5', name: 'ULP KERSIK TUO', parent: 'UP3 MUARA BUNGO', target: 0 },
  { id: '6', name: 'ULP BANGKO', parent: 'UP3 MUARA BUNGO', target: 0 },
  { id: '7', name: 'ULP RMB BUJANG', parent: 'UP3 MUARA BUNGO', target: 0 },
  { id: '8', name: 'ULP KOTA BUNGO', parent: 'UP3 MUARA BUNGO', target: 0 },
  { id: '9', name: 'UP3 JAMBI', parent: 'UP3 JAMBI', target: 0 },
  { id: '10', name: 'ULP KOTA BARU', parent: 'UP3 JAMBI', target: 0 },
  { id: '11', name: 'ULP KUALA TNGKL', parent: 'UP3 JAMBI', target: 0 },
  { id: '12', name: 'ULP MRA BULIAN', parent: 'UP3 JAMBI', target: 0 },
  { id: '13', name: 'ULP MUARA SABAK', parent: 'UP3 JAMBI', target: 0 },
  { id: '14', name: 'ULP SBRNG KOTA', parent: 'UP3 JAMBI', target: 0 },
  { id: '15', name: 'ULP TLANAI PURA', parent: 'UP3 JAMBI', target: 0 },
  { id: '16', name: 'UPT JAMBI', parent: 'UPT JAMBI', target: 0 },
  { id: '17', name: 'ULTG AUR DURI', parent: 'UPT JAMBI', target: 0 },
  { id: '18', name: 'ULTG MUARABUNGO', parent: 'UPT JAMBI', target: 0 },
  { id: '19', name: 'UP2K JAMBI', parent: 'UP2K JAMBI', target: 0 },
  { id: '20', name: 'UPP SUMBAGTENG3', parent: 'UPP SUMBAGTENG3', target: 0 }
];

const INDONESIAN_MONTHS = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

function getCurrentMonthLabel() {
  const now = new Date();
  return INDONESIAN_MONTHS[now.getMonth()] + " " + now.getFullYear();
}

function cleanNip(v) {
  return String(v || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function isNipValid(nipClean) {
  return nipClean && nipClean !== '0' && nipClean !== '-' && nipClean !== 'TIDAKADA' && nipClean !== 'BELUMADA' && nipClean.length >= 3;
}

// ==========================================
// 1. MANAJEMEN DATABASE EXCEL (.xlsx)
// ==========================================

/**
 * Inisialisasi database.xlsx jika file belum tersedia
 */
function initDatabase() {
  if (fs.existsSync(DB_FILE)) {
    console.log(`[Database] File ${DB_FILE} terdeteksi dan siap digunakan.`);
    return;
  }

  console.log(`[Database] Membuat file baru: ${DB_FILE}`);
  const wb = XLSX.utils.book_new();
  const currentMonth = getCurrentMonthLabel();

  // 1. Sheet Daftar_Belum_Survei
  const defaultBelumSurvei = [
    { No: 1, Nama: 'Muhammad Rian', NIP: '199504122020021008', Unit: 'UP3 MUARA BUNGO', Mengikuti_Awareness: 'Ya', Jawaban_Benar: '8/10', Persentase: 80 },
    { No: 2, Nama: 'Putri Handayani', NIP: '199711032021012014', Unit: 'ULP SAROLANGUN', Mengikuti_Awareness: 'Tidak', Jawaban_Benar: 'Tidak Quiz', Persentase: 0 },
    { No: 3, Nama: 'Bayu Saputra', NIP: '199108212015031003', Unit: 'UP3 JAMBI', Mengikuti_Awareness: 'Tidak', Jawaban_Benar: 'Tidak Quiz', Persentase: 0 },
    { No: 4, Nama: 'Lestari Indah', NIP: '199307112018042011', Unit: 'ULTG AUR DURI', Mengikuti_Awareness: 'Ya', Jawaban_Benar: '9/10', Persentase: 90 },
    { No: 5, Nama: 'Slamet Riadi', NIP: '198901152011121004', Unit: 'UPP SUMBAGTENG3', Mengikuti_Awareness: 'Tidak', Jawaban_Benar: 'Tidak Quiz', Persentase: 0 }
  ];
  const wsBelumSurvei = XLSX.utils.json_to_sheet(defaultBelumSurvei, {
    header: ['No', 'Nama', 'NIP', 'Unit', 'Mengikuti_Awareness', 'Jawaban_Benar', 'Persentase']
  });
  XLSX.utils.book_append_sheet(wb, wsBelumSurvei, 'Daftar_Belum_Survei');

  // 2. Sheet Target_Unit
  const defaultTargets = INITIAL_MASTER_UNITS.map(u => ({
    ID: `${u.id}_${currentMonth}`,
    Nama_Unit: u.name,
    Kelompok_Induk: u.parent,
    Target: 0,
    Bulan: currentMonth,
    'Realisasi (Target - Belum Survey)': 0,
    'Tanggal Update': ''
  }));
  const wsTargets = XLSX.utils.json_to_sheet(defaultTargets, {
    header: ['ID', 'Nama_Unit', 'Kelompok_Induk', 'Target', 'Bulan', 'Realisasi (Target - Belum Survey)', 'Tanggal Update']
  });
  XLSX.utils.book_append_sheet(wb, wsTargets, 'Target_Unit');

  // 3. Sheet Referensi_Pegawai
  const defaultRef = [
    { No: 1, Nama: 'Muhammad Rian', NIP: '199504122020021008', Email_Korporat: 'm.rian@pln.co.id', Unit_Pelayanan: 'UP3 MUARA BUNGO', Sub_Unit_Pelayanan: 'UP3 MUARA BUNGO' },
    { No: 2, Nama: 'Putri Handayani', NIP: '199711032021012014', Email_Korporat: 'putri.h@pln.co.id', Unit_Pelayanan: 'UP3 MUARA BUNGO', Sub_Unit_Pelayanan: 'ULP SAROLANGUN' },
    { No: 3, Nama: 'Bayu Saputra', NIP: '199108212015031003', Email_Korporat: 'bayu.s@pln.co.id', Unit_Pelayanan: 'UP3 JAMBI', Sub_Unit_Pelayanan: 'UP3 JAMBI' },
    { No: 4, Nama: 'Lestari Indah', NIP: '199307112018042011', Email_Korporat: 'lestari.i@pln.co.id', Unit_Pelayanan: 'UPT JAMBI', Sub_Unit_Pelayanan: 'ULTG AUR DURI' },
    { No: 5, Nama: 'Slamet Riadi', NIP: '198901152011121004', Email_Korporat: 'slamet.r@pln.co.id', Unit_Pelayanan: 'UPP SUMBAGTENG3', Sub_Unit_Pelayanan: 'UPP SUMBAGTENG3' }
  ];
  const wsRef = XLSX.utils.json_to_sheet(defaultRef, {
    header: ['No', 'Nama', 'NIP', 'Email_Korporat', 'Unit_Pelayanan', 'Sub_Unit_Pelayanan']
  });
  XLSX.utils.book_append_sheet(wb, wsRef, 'Referensi_Pegawai');

  // 4. Sheet Riwayat_Bulanan
  const defaultHist = [
    { Bulan: 'April 2026', Pencapaian: 100, Total_Pegawai: 327, Snapshot_Data: '' }
  ];
  const wsHist = XLSX.utils.json_to_sheet(defaultHist, {
    header: ['Bulan', 'Pencapaian', 'Total_Pegawai', 'Snapshot_Data']
  });
  XLSX.utils.book_append_sheet(wb, wsHist, 'Riwayat_Bulanan');

  // 5. Sheet Config
  const defaultConfig = [
    { Key: 'ADMIN_PASSWORD', Value: 'st1Jambi' }
  ];
  const wsConfig = XLSX.utils.json_to_sheet(defaultConfig, { header: ['Key', 'Value'] });
  XLSX.utils.book_append_sheet(wb, wsConfig, 'Config');

  XLSX.writeFile(wb, DB_FILE);
  console.log(`[Database] database.xlsx berhasil diinisialisasi.`);
}

function readWorkbook() {
  initDatabase();
  return XLSX.readFile(DB_FILE);
}

function writeWorkbook(wb) {
  XLSX.writeFile(wb, DB_FILE);
}

function getSheetData(sheetName) {
  const wb = readWorkbook();
  const ws = wb.Sheets[sheetName];
  if (!ws) return [];
  return XLSX.utils.sheet_to_json(ws, { defval: '' });
}

function saveSheetData(sheetName, jsonData, headers) {
  const wb = readWorkbook();
  const ws = XLSX.utils.json_to_sheet(jsonData, { header: headers });
  wb.Sheets[sheetName] = ws;
  if (!wb.SheetNames.includes(sheetName)) {
    wb.SheetNames.push(sheetName);
  }
  writeWorkbook(wb);
}

// ==========================================
// 2. LOGIKA BISNIS (SINKRON DENGAN CODE.GS)
// ==========================================

function getReferensiPegawaiBackend() {
  try {
    const raw = getSheetData('Referensi_Pegawai');
    return raw.map((r, idx) => ({
      no: Number(r.No) || (idx + 1),
      nama: String(r.Nama || '').trim(),
      nip: String(r.NIP || '').trim(),
      email: String(r.Email_Korporat || '').trim(),
      unit: String(r.Unit_Pelayanan || '').trim(),
      subUnit: String(r.Sub_Unit_Pelayanan || '').trim()
    })).filter(r => r.nama !== '');
  } catch (e) {
    console.error('getReferensiPegawaiBackend error:', e);
    return [];
  }
}

function saveReferensiPegawaiBackend(refList) {
  try {
    const rows = (refList || []).map((e, idx) => ({
      No: idx + 1,
      Nama: String(e.nama || '').trim(),
      NIP: String(e.nip || '').trim(),
      Email_Korporat: String(e.email || '').trim(),
      Unit_Pelayanan: String(e.unit || '').trim(),
      Sub_Unit_Pelayanan: String(e.subUnit || '').trim()
    }));
    saveSheetData('Referensi_Pegawai', rows, ['No', 'Nama', 'NIP', 'Email_Korporat', 'Unit_Pelayanan', 'Sub_Unit_Pelayanan']);
    return { success: true, count: rows.length };
  } catch (e) {
    console.error('saveReferensiPegawaiBackend error:', e);
    return { success: false, error: e.toString() };
  }
}

function deduplicateEmployeesBackend(empList) {
  if (!empList || !Array.isArray(empList)) return [];

  const refEmployees = getReferensiPegawaiBackend();
  const refNipMap = {};
  if (refEmployees.length > 0) {
    refEmployees.forEach(ref => {
      const cNip = cleanNip(ref.nip);
      if (isNipValid(cNip)) {
        refNipMap[cNip] = ref;
      }
    });
  }
  const hasReference = Object.keys(refNipMap).length > 0;

  const keyMap = {};
  empList.forEach((emp, index) => {
    if (!emp) return;
    const nipClean = cleanNip(emp.nip);

    // Filter ketat: jika ada referensi, pegawai wajib terdaftar di referensi!
    if (hasReference) {
      if (!isNipValid(nipClean) || !refNipMap[nipClean]) {
        return; // Lewati karena NIP tidak ada di Referensi Pegawai
      }
      const refItem = refNipMap[nipClean];
      if (refItem && refItem.nama) {
        emp.nama = refItem.nama;
        emp.nip = refItem.nip;
      }
    }

    const key = isNipValid(nipClean) ? ('NIP_' + nipClean) : ('NAME_' + cleanNip(emp.nama) + '_UNIT_' + cleanNip(emp.unit));

    if (!keyMap[key]) {
      keyMap[key] = { ...emp };
    } else {
      const existing = keyMap[key];
      const p1 = parseFloat(existing.persentase) || 0;
      const p2 = parseFloat(emp.persentase) || 0;
      if (p2 > p1) {
        existing.persentase = p2;
        existing.jawabanBenar = emp.jawabanBenar;
        existing.mengikutiAwareness = emp.mengikutiAwareness;
      }
      const existingNipClean = cleanNip(existing.nip);
      if (!isNipValid(existingNipClean) && isNipValid(nipClean)) {
        existing.nip = emp.nip;
      }
    }
  });

  const results = Object.values(keyMap);
  results.forEach((e, idx) => { e.no = idx + 1; });
  return results;
}

function getUncompletedEmployeesBackend() {
  try {
    const raw = getSheetData('Daftar_Belum_Survei');
    const employees = raw.map((r, idx) => ({
      no: Number(r.No) || (idx + 1),
      nama: String(r.Nama || '').trim(),
      nip: String(r.NIP || '').trim(),
      unit: String(r.Unit || 'UP3 MUARA BUNGO').trim(),
      mengikutiAwareness: String(r.Mengikuti_Awareness || 'Ya').trim(),
      jawabanBenar: String(r.Jawaban_Benar || 'Tidak Quiz').trim(),
      persentase: parseFloat(r.Persentase) || 0
    })).filter(r => r.nama !== '');
    return deduplicateEmployeesBackend(employees);
  } catch (e) {
    console.error('getUncompletedEmployeesBackend error:', e);
    return [];
  }
}

function updateAllEmployeesOnSheetBackend(employeesList) {
  try {
    const deduplicated = deduplicateEmployeesBackend(employeesList || []);
    const rows = deduplicated.map((e, idx) => ({
      No: idx + 1,
      Nama: e.nama,
      NIP: e.nip,
      Unit: e.unit,
      Mengikuti_Awareness: e.mengikutiAwareness,
      Jawaban_Benar: e.jawabanBenar,
      Persentase: e.persentase
    }));
    saveSheetData('Daftar_Belum_Survei', rows, ['No', 'Nama', 'NIP', 'Unit', 'Mengikuti_Awareness', 'Jawaban_Benar', 'Persentase']);
    return { success: true, count: rows.length };
  } catch (e) {
    console.error('updateAllEmployeesOnSheetBackend error:', e);
    return { success: false, error: e.toString() };
  }
}

function getUnitTargetsBackend() {
  try {
    const raw = getSheetData('Target_Unit');
    if (raw.length === 0) {
      const currentMonth = getCurrentMonthLabel();
      return INITIAL_MASTER_UNITS.map(u => ({
        id: `${u.id}_${currentMonth}`,
        name: u.name,
        parent: u.parent,
        target: 0,
        bulan: currentMonth,
        realisasi: 0,
        tanggalUpdate: ''
      }));
    }

    return raw.map((r, idx) => ({
      id: String(r.ID || (idx + 1)),
      name: String(r.Nama_Unit || '').trim(),
      parent: String(r.Kelompok_Induk || 'UP3 JAMBI').trim(),
      target: Number(r.Target) || 0,
      bulan: String(r.Bulan || getCurrentMonthLabel()).trim(),
      realisasi: Number(r['Realisasi (Target - Belum Survey)']) || 0,
      tanggalUpdate: String(r['Tanggal Update'] || '').trim()
    })).filter(r => r.name !== '');
  } catch (e) {
    console.error('getUnitTargetsBackend error:', e);
    return [];
  }
}

function updateUnitTargetsBackend(targetsList) {
  try {
    const rows = (targetsList || []).map(t => {
      const uBulan = String(t.bulan || getCurrentMonthLabel()).trim();
      let compoundId = String(t.id || '');
      if (compoundId.includes('_')) {
        compoundId = compoundId.split('_')[0] + '_' + uBulan;
      } else {
        compoundId = compoundId + '_' + uBulan;
      }
      return {
        ID: compoundId,
        Nama_Unit: t.name,
        Kelompok_Induk: t.parent,
        Target: Number(t.target) || 0,
        Bulan: uBulan,
        'Realisasi (Target - Belum Survey)': Number(t.realisasi) || 0,
        'Tanggal Update': String(t.tanggalUpdate || '')
      };
    });

    saveSheetData('Target_Unit', rows, ['ID', 'Nama_Unit', 'Kelompok_Induk', 'Target', 'Bulan', 'Realisasi (Target - Belum Survey)', 'Tanggal Update']);
    return { success: true, count: rows.length };
  } catch (e) {
    console.error('updateUnitTargetsBackend error:', e);
    return { success: false, error: e.toString() };
  }
}

function getMonthlyHistoryBackend() {
  try {
    const raw = getSheetData('Riwayat_Bulanan');
    return raw.map(r => ({
      bulan: String(r.Bulan || ''),
      pencapaian: parseFloat(r.Pencapaian) || 0,
      totalPegawai: Number(r.Total_Pegawai) || 0,
      snapshotData: String(r.Snapshot_Data || '')
    }));
  } catch (e) {
    return [];
  }
}

function updateMonthlyHistoryBackend(historyList) {
  try {
    const rows = (historyList || []).map(h => ({
      Bulan: h.bulan,
      Pencapaian: h.pencapaian,
      Total_Pegawai: h.totalPegawai,
      Snapshot_Data: typeof h.snapshotData === 'object' ? JSON.stringify(h.snapshotData) : String(h.snapshotData || '')
    }));
    saveSheetData('Riwayat_Bulanan', rows, ['Bulan', 'Pencapaian', 'Total_Pegawai', 'Snapshot_Data']);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function verifyPasswordBackend(password) {
  const cfg = getSheetData('Config');
  const row = cfg.find(c => c.Key === 'ADMIN_PASSWORD');
  const validPass = row ? String(row.Value).trim() : 'st1Jambi';
  return String(password).trim() === validPass;
}

// ==========================================
// 3. RPC ADAPTER (PENGGANTI GOOGLE.SCRIPT.RUN)
// ==========================================

app.post('/api/gas', (req, res) => {
  const { action, args = [] } = req.body;
  try {
    switch (action) {
      case 'getUncompletedEmployees':
        return res.json(getUncompletedEmployeesBackend());

      case 'updateAllEmployeesOnSheet':
        return res.json(updateAllEmployeesOnSheetBackend(args[0]));

      case 'getUnitTargets':
        return res.json(getUnitTargetsBackend());

      case 'updateUnitTargets':
        return res.json(updateUnitTargetsBackend(args[0]));

      case 'getReferensiPegawai':
        return res.json(getReferensiPegawaiBackend());

      case 'saveReferensiPegawai':
        return res.json(saveReferensiPegawaiBackend(args[0]));

      case 'getCSAMasterData':
        return res.json({
          employees: getReferensiPegawaiBackend(),
          uncompleted: getUncompletedEmployeesBackend(),
          targets: getUnitTargetsBackend(),
          history: getMonthlyHistoryBackend()
        });

      case 'updateMonthlyHistory':
        return res.json(updateMonthlyHistoryBackend(args[0]));

      case 'verifyPasswordOnServer':
        return res.json(verifyPasswordBackend(args[0]));

      default:
        console.warn(`[API] Action tidak dikenal: ${action}`);
        return res.status(400).json({ error: `Action '${action}' tidak dikenali di server lokal.` });
    }
  } catch (err) {
    console.error(`[API] Error menjalankan action ${action}:`, err);
    return res.status(500).json({ error: err.toString() });
  }
});

// REST API Endpoints alternatif
app.get('/api/employees', (req, res) => res.json(getUncompletedEmployeesBackend()));
app.post('/api/employees', (req, res) => res.json(updateAllEmployeesOnSheetBackend(req.body)));

app.get('/api/targets', (req, res) => res.json(getUnitTargetsBackend()));
app.post('/api/targets', (req, res) => res.json(updateUnitTargetsBackend(req.body)));

app.get('/api/reference', (req, res) => res.json(getReferensiPegawaiBackend()));
app.post('/api/reference', (req, res) => res.json(saveReferensiPegawaiBackend(req.body)));

app.post('/api/verify-password', (req, res) => {
  const { password } = req.body;
  res.json({ valid: verifyPasswordBackend(password) });
});

// Unduh file database.xlsx langsung dari browser
app.get('/api/download-database', (req, res) => {
  initDatabase();
  res.download(DB_FILE, 'database_csa_selesaisurvei.xlsx');
});

// ==========================================
// 4. SERVING STATIC FILES & TEMPLATE LOADER
// ==========================================

// Serve folder assets (jika ada file icon atau library offline)
if (fs.existsSync(path.resolve(__dirname, 'assets'))) {
  app.use('/assets', express.static(path.resolve(__dirname, 'assets')));
}

// Serve halaman utama (Index.html)
app.get(['/', '/index.html'], (req, res) => {
  try {
    let htmlContent = '';
    const candidates = [
      path.resolve(__dirname, 'index.html'),
      path.resolve(__dirname, 'Index.html'),
      path.resolve(__dirname, 'index.htm'),
      path.resolve(__dirname, 'Index.htm'),
      path.resolve(__dirname, 'index.html.txt'),
      path.resolve(__dirname, 'Index.html.txt'),
      path.resolve(__dirname, 'google-apps-script/index.html'),
      path.resolve(__dirname, 'google-apps-script/Index.html')
    ];

    let foundPath = null;
    for (const p of candidates) {
      if (fs.existsSync(p)) {
        foundPath = p;
        break;
      }
    }

    if (foundPath) {
      htmlContent = fs.readFileSync(foundPath, 'utf-8');
    } else {
      return res.status(404).send('File index.html tidak ditemukan di direktori.');
    }

    // Set header anti-cache agar browser selalu membaca konten terbaru
    res.setHeader('Content-Type', 'text/html; charset=utf-8');
    res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate, max-age=0');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Expires', '0');

    // Validasi apakah file index.html yang dibaca benar-benar file HTML tampilan dashboard
    const hasHtmlTag = htmlContent.toLowerCase().includes('<html') || htmlContent.toLowerCase().includes('<!doctype html');
    const isServerCodeWithoutHtml = !hasHtmlTag && (
      htmlContent.includes('import express') || 
      htmlContent.includes('require(\'express\')') || 
      htmlContent.includes('require("express")') || 
      htmlContent.includes('app.listen(')
    );

    if (!hasHtmlTag || isServerCodeWithoutHtml) {
      console.warn(`[PERINGATAN] File "${foundPath}" terdeteksi berisi kode JavaScript/server, BUKAN kode HTML tampilan dashboard!`);
      const errorPage = `<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8">
  <title>Peringatan: File index.html Belum Sesuai</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: #0f172a; color: #f8fafc; padding: 40px 20px; display: flex; justify-content: center; }
    .card { max-width: 650px; background: #1e293b; border: 1px solid #334155; border-radius: 16px; padding: 32px; box-shadow: 0 10px 25px rgba(0,0,0,0.5); }
    h1 { color: #f59e0b; font-size: 22px; margin-bottom: 16px; display: flex; align-items: center; gap: 10px; }
    p { color: #cbd5e1; line-height: 1.6; margin-bottom: 16px; }
    .badge { background: #334155; color: #38bdf8; padding: 4px 8px; border-radius: 6px; font-family: monospace; font-size: 13px; word-break: break-all; }
    .steps { background: #0f172a; border-radius: 8px; padding: 16px 20px; margin: 20px 0; border-left: 4px solid #3b82f6; }
    .steps ol { margin-left: 20px; color: #e2e8f0; }
    .steps li { margin-bottom: 8px; }
    .btn { display: inline-block; background: #2563eb; color: white; padding: 10px 20px; border-radius: 8px; text-decoration: none; font-weight: 600; cursor: pointer; border: none; font-size: 15px; }
    .btn:hover { background: #1d4ed8; }
  </style>
</head>
<body>
  <div class="card">
    <h1>⚠️ File index.html Masih Berisi Kode Server</h1>
    <p>Server membaca file: <br><span class="badge">${foundPath}</span></p>
    <p>File tersebut terdeteksi <strong>berisi kode JavaScript (server.js)</strong>, sehingga tampilan dashboard belum dapat dimuat.</p>
    
    <div class="steps">
      <h3 style="color:#38bdf8; margin-bottom:10px;">Langkah Memperbaikinya:</h3>
      <ol>
        <li>Buka file <b>index.html</b> di folder server Anda dengan Notepad atau VS Code.</li>
        <li>Hapus seluruh isinya (tekan Ctrl+A lalu Delete).</li>
        <li>Copy seluruh isi file <b>index.html</b> yang benar (yang diawali <code>&lt;!DOCTYPE html&gt;</code>), lalu Paste ke file tersebut.</li>
        <li>Simpan file (Ctrl+S).</li>
        <li>Klik tombol <b>Muat Ulang Halaman</b> di bawah ini.</li>
      </ol>
    </div>

    <button class="btn" onclick="window.location.reload(true)">🔄 Muat Ulang Halaman</button>
  </div>
</body>
</html>`;
      return res.send(errorPage);
    }

    // Resolusi include Apps Script jika ada (<?!= include('CSS') ?>)
    const includeRegex = /<\?[=!]\s*include\(['"]([^'"]+)['"]\);\s*\?>/g;
    htmlContent = htmlContent.replace(includeRegex, (match, filename) => {
      const candidates = [
        path.resolve(__dirname, `${filename}.html`),
        path.resolve(__dirname, `google-apps-script/${filename}.html`)
      ];
      for (const p of candidates) {
        if (fs.existsSync(p)) {
          return fs.readFileSync(p, 'utf-8');
        }
      }
      return `<!-- Include ${filename} tidak ditemukan -->`;
    });

    // Injeksi otomatis Bridge Google Apps Script jika belum terpasang
    const bridgeScript = `
    <script>
      // Otomatisasi Bridge Client untuk Server Lokal Express
      if (typeof window.google === 'undefined' || !window.google.script || !window.google.script.run) {
        window.google = window.google || {};
        window.google.script = window.google.script || {};
        
        function createGasRunner(successCb, errorCb) {
          return new Proxy({}, {
            get(target, prop) {
              if (prop === 'withSuccessHandler') {
                return (cb) => createGasRunner(cb, errorCb);
              }
              if (prop === 'withFailureHandler') {
                return (cb) => createGasRunner(successCb, cb);
              }
              return async (...args) => {
                try {
                  const res = await fetch('/api/gas', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ action: prop, args: args })
                  });
                  const data = await res.json();
                  if (data && data.error) {
                    if (errorCb) errorCb(new Error(data.error));
                  } else {
                    if (successCb) successCb(data.result !== undefined ? data.result : data);
                  }
                } catch (err) {
                  console.error('[GAS Bridge Error]', err);
                  if (errorCb) errorCb(err);
                }
              };
            }
          });
        }
        window.google.script.run = createGasRunner();
        console.log('[SelesaiSurvei] Server Lokal Mode aktif - Database tersambung ke database.xlsx');
      }
    </script>
    `;

    // Sisipkan bridge sebelum tag penutup </head> atau <body>
    if (htmlContent.includes('</head>')) {
      htmlContent = htmlContent.replace('</head>', `${bridgeScript}\n</head>`);
    } else {
      htmlContent = bridgeScript + htmlContent;
    }

    return res.send(htmlContent);
  } catch (err) {
    console.error('Error saat melayani Index.html:', err);
    return res.status(500).send('Terjadi kesalahan pada server saat memuat Index.html: ' + err.message);
  }
});

// Jalankan server
const server = app.listen(PORT, '0.0.0.0', () => {
  initDatabase();
  console.log(`====================================================`);
  console.log(`🚀 SelesaiSurvei Local Server Berjalan!`);
  console.log(`🌐 Akses: http://localhost:${PORT}`);
  console.log(`📊 Database: ${DB_FILE}`);
  console.log(`====================================================`);
});

server.on('error', (err) => {
  if (err.code === 'EADDRINUSE') {
    const fallbackPort = Number(PORT) + 1;
    console.warn(`[Server] Port ${PORT} sedang dipakai, beralih ke port ${fallbackPort}...`);
    app.listen(fallbackPort, '0.0.0.0', () => {
      initDatabase();
      console.log(`====================================================`);
      console.log(`🚀 SelesaiSurvei Local Server Berjalan!`);
      console.log(`🌐 Akses: http://localhost:${fallbackPort}`);
      console.log(`📊 Database: ${DB_FILE}`);
      console.log(`====================================================`);
    });
  } else {
    console.error('[Server Error]', err);
  }
});

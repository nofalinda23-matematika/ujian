/**
 * KONFIGURASI PARAMETER APLIKASI CBT
 * Sesuaikan parameter di bawah ini sesuai kebutuhan ujian Anda.
 */
const APP_CONFIG = {
  JUDUL_UJIAN: "Sumatif Akhir Semester (SAS)",
  MATA_PELAJARAN: "Literasi & Numerasi Terpadu",
  SATUAN_PENDIDIKAN: "SMPN 01 Konoha",
  KELAS_FASE: "Kelas 7 / Fase D",
  DURASI_MENIT: 90,
  PASSWORD_ADMIN: "123", // Ubah ke "admincbt123" atau sesuai kebutuhan
  IZINKAN_LIHAT_PEMBAHASAN: true // Ubah ke false jika siswa tidak boleh melihat pembahasan setelah selesai
};

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_CONFIG.JUDUL_UJIAN + ' - ' + APP_CONFIG.SATUAN_PENDIDIKAN)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * FUNGSI OTOMATISASI DATABASE GOOGLE SHEETS
 * Jalankan fungsi ini SEKALI dari editor Apps Script untuk membuat sheet & data awal.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Setup Sheet Master_Soal
  let sheetSoal = ss.getSheetByName("Master_Soal");
  if (!sheetSoal) {
    sheetSoal = ss.insertSheet("Master_Soal");
  } else {
    sheetSoal.clear();
  }
  const headerSoal = [
    "ID_Soal", "Nomor_Soal", "Jenis_Soal", "Teks_Stimulus", "Gambar_Stimulus_URL",
    "Teks_Soal", "Gambar_Soal_URL", "Pilihan_A", "Pilihan_B", "Pilihan_C", "Pilihan_D",
    "Kunci_Jawaban", "Pembahasan", "Bobot"
  ];
  sheetSoal.appendRow(headerSoal);
  sheetSoal.getRange("A1:N1").setFontWeight("bold").setBackground("#1e3a8a").setFontColor("#ffffff");

  const sampleSoal = [
    [
      "SOAL-001", 1, "PG",
      "Bacalah kutipan teks berikut dengan saksama!\n\nEnergi terbarukan seperti tenaga surya dan angin semakin diminati sebagai pengganti bahan bakar fosil. Selain ramah lingkungan karena tidak menghasilkan emisi karbon secara langsung, biaya instalasi panel surya dalam satu dekade terakhir turun hingga 80%. Namun, tantangan utama energi terbarukan terletak pada sifatnya yang intermiten (bergantung pada cuaca) sehingga membutuhkan teknologi baterai penyimpan daya yang efisien.",
      "https://unsplash.com",
      "Berdasarkan bacaan di samping, apa tantangan utama dalam pemanfaatan energi terbarukan?",
      "",
      "Biaya instalasi yang terus meningkat setiap tahun",
      "Menghasilkan emisi karbon dalam jumlah besar",
      "Ketergantungan pada kondisi cuaca sehingga butuh baterai penyimpan",
      "Tidak dapat digunakan untuk kebutuhan rumah tangga",
      "C",
      "Pada kalimat terakhir bacaan disebutkan bahwa tantangan utama energi terbarukan terletak pada sifatnya yang intermiten (bergantung pada cuaca) sehingga membutuhkan baterai penyimpan daya.",
      20
    ],
    [
      "SOAL-002", 2, "PG_Kompleks",
      "Perhatikan data kandungan gizi per 100 gram bahan makanan berikut:\n• Tempe: Protein 19g, Lemak 4g, Karbohidrat 9g\n• Telur Ayam: Protein 13g, Lemak 11g, Karbohidrat 1.1g\n• Dada Ayam: Protein 31g, Lemak 3.6g, Karbohidrat 0g\n• Nasi Putih: Protein 2.7g, Lemak 0.3g, Karbohidrat 28g",
      "",
      "Manakah pernyataan yang BENAR berdasarkan tabel kandungan gizi tersebut? (Pilih lebih dari satu jawaban benar)",
      "",
      "Dada ayam memiliki kandungan protein paling tinggi di antara keempat bahan.",
      "Tempe mengandung karbohidrat lebih tinggi dibandingkan nasi putih.",
      "Telur ayam memiliki kandungan lemak paling tinggi di antara keempat bahan.",
      "Nasi putih adalah sumber protein utama dengan kadar 28g.",
      "A,C",
      "Dada ayam memiliki protein tertinggi (31g) dan telur ayam memiliki lemak tertinggi (11g). Jadi jawaban yang tepat adalah A dan C.",
      40
    ],
    [
      "SOAL-003", 3, "Benar_Salah",
      "Sebuah akuarium berbentuk balok memiliki ukuran panjang 60 cm, lebar 30 cm, dan tinggi 40 cm. Akuarium tersebut akan diisi air hingga 3/4 bagian dari tinggi totalnya. (Catatan: 1 liter = 1.000 cm³)",
      "",
      "Tentukan kebenaran pernyataan berikut: Volume air yang dibutuhkan untuk mengisi 3/4 bagian akuarium tersebut adalah 54 Liter.",
      "",
      "Benar",
      "Salah",
      "-",
      "-",
      "A",
      "Volume total = 60 x 30 x 40 = 72.000 cm³ = 72 Liter. Volume 3/4 bagian = 3/4 x 72 Liter = 54 Liter. Maka pernyataan tersebut BENAR (Pilihan A).",
      40
    ]
  ];
  sampleSoal.forEach(row => sheetSoal.appendRow(row));

  // 2. Setup Sheet Master_Siswa
  let sheetSiswa = ss.getSheetByName("Master_Siswa");
  if (!sheetSiswa) {
    sheetSiswa = ss.insertSheet("Master_Siswa");
  } else {
    sheetSiswa.clear();
  }
  const headerSiswa = ["Nama_Siswa", "Kelas", "Token_Akses"];
  sheetSiswa.appendRow(headerSiswa);
  sheetSiswa.getRange("A1:C1").setFontWeight("bold").setBackground("#065f46").setFontColor("#ffffff");

  const sampleSiswa = [
    ["Naruto Uzumaki", "7A", "KONOHA01"],
    ["Sasuke Uchiha", "7A", "KONOHA02"],
    ["Sakura Haruno", "7A", "KONOHA03"],
    ["Shikamaru Nara", "7B", "KONOHA04"],
    ["Hinata Hyuga", "7B", "KONOHA05"]
  ];
  sampleSiswa.forEach(row => sheetSiswa.appendRow(row));

  // 3. Setup Sheet Hasil_Ujian
  let sheetHasil = ss.getSheetByName("Hasil_Ujian");
  if (!sheetHasil) {
    sheetHasil = ss.insertSheet("Hasil_Ujian");
  } else {
    sheetHasil.clear();
  }
  const headerHasil = [
    "Timestamp", "Nama_Siswa", "Kelas", "Nilai_Akhir",
    "Total_Benar", "Total_Salah", "Pelanggaran_Tab", "Detail_Jawaban_JSON"
  ];
  sheetHasil.appendRow(headerHasil);
  sheetHasil.getRange("A1:H1").setFontWeight("bold").setBackground("#7c2d12").setFontColor("#ffffff");

  // Hapus Sheet1 bawaan jika kosong
  const defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }
}

/**
 * Mengambil konfigurasi umum untuk ditampilkan di UI
 */
function getAppConfig() {
  return {
    judulUjian: APP_CONFIG.JUDUL_UJIAN,
    mataPelajaran: APP_CONFIG.MATA_PELAJARAN,
    satuanPendidikan: APP_CONFIG.SATUAN_PENDIDIKAN,
    kelasFase: APP_CONFIG.KELAS_FASE,
    durasiMenit: APP_CONFIG.DURASI_MENIT,
    izinkanPembahasan: APP_CONFIG.IZINKAN_LIHAT_PEMBAHASAN
  };
}

/**
 * Validasi Login Siswa
 */
function loginSiswa(nama, token) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Master_Siswa");
  if (!sheet) return { success: false, message: "Database Master_Siswa belum dibuat!" };

  const data = sheet.getDataRange().getValues();
  const inputNama = String(nama).trim().toLowerCase();
  const inputToken = String(token).trim().toUpperCase();

  for (let i = 1; i < data.length; i++) {
    const dbNama = String(data[i][0]).trim().toLowerCase();
    const dbKelas = String(data[i][1]).trim();
    const dbToken = String(data[i][2]).trim().toUpperCase();

    if (dbNama === inputNama && dbToken === inputToken) {
      return {
        success: true,
        siswa: {
          nama: data[i][0],
          kelas: dbKelas,
          token: dbToken
        }
      };
    }
  }
  return { success: false, message: "Nama Siswa atau Token Akses tidak cocok!" };
}

/**
 * Validasi Login Admin / Guru
 */
function loginAdmin(password) {
  if (String(password).trim() === String(APP_CONFIG.PASSWORD_ADMIN)) {
    return { success: true };
  }
  return { success: false, message: "Password Admin salah!" };
}

/**
 * KEAMANAN: Mengambil daftar soal untuk siswa TANPA Kunci_Jawaban dan Pembahasan
 */
function getQuestionsForStudent() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Master_Soal");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const questions = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    questions.push({
      id: data[i][0],
      nomor: Number(data[i][1]),
      jenis: String(data[i][2]).trim(), // PG, PG_Kompleks, Benar_Salah
      teksStimulus: String(data[i][3] || ""),
      gambarStimulus: String(data[i][4] || ""),
      teksSoal: String(data[i][5] || ""),
      gambarSoal: String(data[i][6] || ""),
      pilihanA: String(data[i][7] || ""),
      pilihanB: String(data[i][8] || ""),
      pilihanC: String(data[i][9] || ""),
      pilihanD: String(data[i][10] || ""),
      bobot: Number(data[i][13] || 1)
      // Kolom index 11 (Kunci_Jawaban) & 12 (Pembahasan) SENGAJA DIHILANGKAN
    });
  }

  questions.sort((a, b) => a.nomor - b.nomor);
  return questions;
}

/**
 * AUTO-SCORING ENGINE DI BACKEND
 * Menerima jawaban siswa, menghitung skor dengan aman di server, menyimpan ke Hasil_Ujian
 */
function submitUjian(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSoal = ss.getSheetByName("Master_Soal");
  const sheetHasil = ss.getSheetByName("Hasil_Ujian");

  const dataSoal = sheetSoal.getDataRange().getValues();
  const jawabanSiswa = payload.jawaban || {}; // Format: { "SOAL-001": ["C"], "SOAL-002": ["A", "C"] }

  let totalBobotMaksimal = 0;
  let bobotDiperoleh = 0;
  let totalBenar = 0;
  let totalSalah = 0;
  const detailReview = [];

  for (let i = 1; i < dataSoal.length; i++) {
    const idSoal = String(dataSoal[i][0]);
    if (!idSoal) continue;

    const nomor = Number(dataSoal[i][1]);
    const jenis = String(dataSoal[i][2]).trim();
    const teksSoal = String(dataSoal[i][5]);
    const kunciRaw = String(dataSoal[i][11]).trim().toUpperCase();
    const pembahasan = String(dataSoal[i][12] || "Tidak ada pembahasan.");
    const bobot = Number(dataSoal[i][13] || 1);

    totalBobotMaksimal += bobot;

    // Normalisasi kunci jawaban menjadi array terurut
    const kunciArr = kunciRaw.split(",").map(k => k.trim()).filter(Boolean).sort();
    // Normalisasi jawaban siswa menjadi array terurut
    const jwbArr = (jawabanSiswa[idSoal] || []).map(j => String(j).trim().toUpperCase()).sort();


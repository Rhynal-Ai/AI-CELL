// =====================
// PENGATURAN UTAMA
// =====================
const ZONA_WAKTU = "Asia/Jakarta";

// Fungsi WAJIB agar halaman bisa dibuka
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('TOKO AI CELL')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =====================
// DATA BARANG
// =====================
function ambilSemuaBarang() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data Barang");
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1);
}

// =====================
// MANAJEMEN DATA BARANG (TAMBAH & EDIT)
// =====================
function ambilDataBarangLengkap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data Barang");
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const hasil = [];
  
  for (let i = 1; i < data.length; i++) {
    hasil.push({
      row: i + 1, // Baris asli di spreadsheet untuk keperluan edit
      kode: data[i][0] || "",
      nama: data[i][1] || "",
      harga: data[i][2] || 0,
      stok: data[i][3] || 0,
      modal: data[i][4] || 0
    });
  }
  return hasil;
}

function tambahBarang(itemBaru) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data Barang");
  if (!sheet) throw new Error("Sheet Data Barang tidak ditemukan");

  const data = sheet.getDataRange().getValues();
  
  // Validasi Cek Duplikat Kode/ID Barang
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === String(itemBaru.kode).trim().toLowerCase()) {
      return {
        status: "error",
        pesan: "⚠️ Gagal! Kode/ID barang '" + itemBaru.kode + "' sudah terdaftar. Gunakan kode lain agar tidak ganda."
      };
    }
  }

  // Jika belum ada, masukkan baris baru
  sheet.appendRow([
    itemBaru.kode,
    itemBaru.nama,
    Number(itemBaru.harga) || 0,
    Number(itemBaru.stok) || 0,
    Number(itemBaru.modal) || 0
  ]);

  return {
    status: "success",
    pesan: "✅ Barang baru berhasil ditambahkan!"
  };
}

function updateBarang(itemEdit) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data Barang");
  if (!sheet) throw new Error("Sheet Data Barang tidak ditemukan");

  const rowNum = Number(itemEdit.row);
  if (!rowNum || rowNum < 2) throw new Error("Baris data tidak valid");

  sheet.getRange(rowNum, 1).setValue(itemEdit.kode);
  sheet.getRange(rowNum, 2).setValue(itemEdit.nama);
  sheet.getRange(rowNum, 3).setValue(Number(itemEdit.harga) || 0);
  sheet.getRange(rowNum, 4).setValue(Number(itemEdit.stok) || 0);
  if (sheet.getLastColumn() >= 5) {
    sheet.getRange(rowNum, 5).setValue(Number(itemEdit.modal) || 0);
  }

  return {
    status: "success",
    pesan: "✅ Data barang berhasil diperbarui!"
  };
}

function hapusBarang(rowNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data Barang");
  if (!sheet) throw new Error("Sheet Data Barang tidak ditemukan");

  sheet.deleteRow(Number(rowNum));
  return {
    status: "success",
    pesan: "✅ Barang berhasil dihapus!"
  };
}

// =====================
// SIMPAN PENJUALAN BARANG
// =====================
function simpanPenjualan(daftarItem, uangBayar, uangKembali) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Laporan Penjualan");
  if (!sheet) throw new Error("Sheet Laporan Penjualan tidak ditemukan");

  const tglSekarang = new Date();
  const tglStr = Utilities.formatDate(tglSekarang, ZONA_WAKTU, "yyyy-MM-dd HH:mm:ss");
  const noTrx = "TRX-" + Utilities.formatDate(tglSekarang, ZONA_WAKTU, "yyyyMMdd-HHmmss");

  daftarItem.forEach(item => {
    const modalSatuan = Number(item.modal) || 0;
    const laba = Math.round((item.harga - modalSatuan) * item.jumlah);

    sheet.appendRow([
      tglStr,
      noTrx,
      item.kode,
      item.nama,
      item.harga,
      item.jumlah,
      item.sub,
      modalSatuan,
      laba,
      uangKembali,
      uangBayar
    ]);

    // Kurangi stok barang
    const sheetBarang = ss.getSheetByName("Data Barang");
    const dataBarang = sheetBarang.getDataRange().getValues();
    for (let i = 1; i < dataBarang.length; i++) {
      if (String(dataBarang[i][0]) === String(item.kode)) {
        const stokBaru = Number(dataBarang[i][3]) - item.jumlah;
        sheetBarang.getRange(i + 1, 4).setValue(stokBaru >= 0 ? stokBaru : 0);
        break;
      }
    }
  });

  return true;
}

// =====================
// SIMPAN LAYANAN / PULSA
// =====================
function simpanLayanan(jenis, keterangan, modal, jual) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Laporan Cell");
  if (!sheet) throw new Error("Sheet Laporan Cell tidak ditemukan");

  const tglSekarang = new Date();
  const tglStr = Utilities.formatDate(tglSekarang, ZONA_WAKTU, "yyyy-MM-dd HH:mm:ss");
  const laba = Math.round(Number(jual) - Number(modal));

  sheet.appendRow([
    tglStr,
    jenis,
    keterangan || "-",
    Number(modal) || 0,
    Number(jual) || 0,
    laba
  ]);

  return true;
}

// =====================
// AMBIL LAPORAN PENJUALAN
// =====================
function ambilLaporanJual(tipe, nilai) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Laporan Penjualan");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const hasil = [];

  for (let i = 1; i < data.length; i++) {
    const baris = data[i];
    const tglAsli = new Date(baris[0]);
    if (isNaN(tglAsli)) continue;

    let cocok = false;
    if (tipe === "hari") {
      const str = Utilities.formatDate(tglAsli, ZONA_WAKTU, "yyyy-MM-dd");
      cocok = str === nilai;
    } else if (tipe === "bulan") {
      const str = Utilities.formatDate(tglAsli, ZONA_WAKTU, "yyyy-MM");
      cocok = str === nilai;
    } else if (tipe === "tahun") {
      const str = Utilities.formatDate(tglAsli, ZONA_WAKTU, "yyyy");
      cocok = str === nilai;
    }

    if (cocok) {
      hasil.push({
        tgl: Utilities.formatDate(tglAsli, ZONA_WAKTU, "dd/MM/yyyy HH:mm"),
        noTransaksi: baris[1] || "-",
        idBarang: baris[2] || "-",
        nama: baris[3] || "-",
        harga: Math.round(Number(baris[4]) || 0),
        jumlah: Number(baris[5]) || 0,
        subtotal: Math.round(Number(baris[6]) || 0),
        laba: Math.round(Number(baris[8]) || 0)
      });
    }
  }

  return hasil;
}

// =====================
// AMBIL LAPORAN LAYANAN
// =====================
function ambilLaporanLayanan(tipe, nilai) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Laporan Cell");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const hasil = [];

  for (let i = 1; i < data.length; i++) {
    const baris = data[i];
    const tglAsli = new Date(baris[0]);
    if (isNaN(tglAsli)) continue;

    let cocok = false;
    if (tipe === "hari") {
      const str = Utilities.formatDate(tglAsli, ZONA_WAKTU, "yyyy-MM-dd");
      cocok = str === nilai;
    } else if (tipe === "bulan") {
      const str = Utilities.formatDate(tglAsli, ZONA_WAKTU, "yyyy-MM");
      cocok = str === nilai;
    } else if (tipe === "tahun") {
      const str = Utilities.formatDate(tglAsli, ZONA_WAKTU, "yyyy");
      cocok = str === nilai;
    }

    if (cocok) {
      hasil.push({
        tgl: Utilities.formatDate(tglAsli, ZONA_WAKTU, "dd/MM/yyyy HH:mm"),
        jenis: baris[1] || "-",
        modal: Math.round(Number(baris[3]) || 0),
        jual: Math.round(Number(baris[4]) || 0),
        laba: Math.round(Number(baris[5]) || 0)
      });
    }
  }

  return hasil;
}

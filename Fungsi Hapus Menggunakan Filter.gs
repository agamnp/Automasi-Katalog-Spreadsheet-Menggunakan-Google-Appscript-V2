  // ======== Konfigurasi Filter Options Default ========
      const filterOptionsDefault = {
        kodeRef: { aktif: false, nilaiDiperbolehkan: [''] },   // ['001.01','001.02','001.03']},
        tahun: { aktif: false, min: 0, max: 9999 },            // min: 2024 max:2026
        halaman: { aktif: false, min: 0, max: 99999 },           // min: 2024 max:2026
        harga: { aktif: false, min: 0, max: 99999999 },         // min: 20000 max:25000
        kategori: { aktif: false, nilaiDiperbolehkan: [''] },  // ['fiksi', 'bahasa', 'puisi','cerpen','novel']}
      };
  // ======== Konfigurasi Filter Options Default ======== 
  
  //     ========     MengHapus Data Menggunakan Filter 1 Sheet    ========
    function HapusDataDenganKriteria() {
      modulFungsiHapusDataDenganKriteria() ;
    }
  //     ========     MengHapus Data Menggunakan Filter 1 Sheet    ========
  
  //     ========     MengHapus Data Menggunakan Filter Semua Sheet    ========
    function HapusDataDenganKriteriaALLSHEET() {
    //PERLU DIRUBAH ============================================================================== 
    const sheetMulai = 0; 
    //PERLU DIRUBAH ==============================================================================

      modulFungsiHapusDataDenganKriteriaALLSHEET(sheetMulai) ;

    }

  //     ========     MengHapus Data Menggunakan Filter Semua Sheet    ========

  //     ========     MengHapus Data Menggunakan Filter Dengan Batch    ========
    function HapusDataDenganKriteriaBATCH() {
      jalankanSemuaSheetPenerbitBatch_HapusDataDenganKriteria();
    }
  //     ========     MengHapus Data Menggunakan Filter Dengan Batch    ========
 
 
 



  //     ========     Fungsi Utama Hapus Data Menggunakan Filter  1 Sheet   ========
    function modulFungsiHapusDataDenganKriteria(sheet) {     
          const ss = SpreadsheetApp.getActiveSpreadsheet();    
      if (!sheet) sheet = ss.getActiveSheet();
          const name = sheet.getName();
          const namaPenerbit = name.replace(/^\d+\.\s*/, "").trim();
          const skipSheets = ['Form Pengadaan', 'Hasil Seleksi', 'Referensi', 'DaftarISBN', 'DaftarUUID'];

          // Lewati jika termasuk sheet yang di-skip
          if (skipSheets.includes(name)) {
            Logger.log("⏭ Melewati sheet: " + name);
            return;
          }

          Logger.log("🔍 Memproses sheet: " + name);

          // pakai default filter
          const filterOptions = JSON.parse(JSON.stringify(filterOptionsDefault));

          const startRow = 10;
          const headerRow = 9;
          const lastRow = sheet.getLastRow();
          const lastCol = sheet.getLastColumn();

          if (lastRow < startRow) {
            Logger.log("ℹ️ Tidak ada data yang diproses.");
            return;
          }

          // Ambil header & data
          const headers = sheet.getRange(headerRow, 1, 1, lastCol)
            .getValues()[0]
            .map(h => String(h).trim().toLowerCase());
          const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues();

          // Index kolom
          const getIndex = (colName) => headers.indexOf(colName.toLowerCase());
          const idx = {
            kodeRef: getIndex("kode referensi"),
            kategori: getIndex("kategori*"),
            tahun: getIndex("tahun terbit digital*"),
            halaman: getIndex("jumlah halaman*"),
            harga: getIndex("harga satuan")
          };

          // Fungsi pengecekan baris lolos filter
          const isLolosFilter = (row) => {
            if (filterOptions.kodeRef.aktif && idx.kodeRef !== -1) {
              const val = String(row[idx.kodeRef]).trim();
              if (!filterOptions.kodeRef.nilaiDiperbolehkan.includes(val)) return false;
            }
            if (filterOptions.kategori.aktif && idx.kategori !== -1) {
              const val = String(row[idx.kategori]).trim().toLowerCase();
            const allowed = filterOptions.kategori.nilaiDiperbolehkan.map(v => v.toLowerCase());
            if (!allowed.includes(val)) return false;
          }
            if (filterOptions.tahun.aktif && idx.tahun !== -1) {
              const val = Number(row[idx.tahun]);
              if (isNaN(val) || val < filterOptions.tahun.min || val > filterOptions.tahun.max) return false;
            }
            if (filterOptions.halaman.aktif && idx.halaman !== -1) {
              const val = Number(row[idx.halaman]);
              if (isNaN(val) || val < filterOptions.halaman.min || val > filterOptions.halaman.max) return false;
            }
            if (filterOptions.harga.aktif && idx.harga !== -1) {
              const val = Number(row[idx.harga]);
              if (isNaN(val) || val < filterOptions.harga.min || val > filterOptions.harga.max) return false;
            }
            return true;
          };

          // Filter data
          const dataLolos = data.filter(isLolosFilter);

          // Jika semua baris gagal → hapus sheet
          if (dataLolos.length === 0) {
            const allSheets = ss.getSheets();
            if (allSheets.length > 1) {
              Logger.log("🗑 Semua data tidak memenuhi syarat. Menghapus sheet: " + name);
              modulHapusBarisDariHasilSeleksi(namaPenerbit);
              ss.deleteSheet(sheet);
            } else {
              Logger.log("⚠️ Tidak bisa menghapus sheet karena hanya ada satu sheet.");
            }
            return;
          }

          // Kosongkan isi data lama
          sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol).clearContent();

          // Tulis ulang data yang lolos
          sheet.getRange(startRow, 1, dataLolos.length, lastCol).setValues(dataLolos);

          // Bersihkan baris kosong & format di bawah data
          const lastDataRow = startRow + dataLolos.length - 1;
          if (lastDataRow < sheet.getMaxRows()) {
            const rangeToClear = sheet.getRange(lastDataRow + 1, 1, sheet.getMaxRows() - lastDataRow, lastCol);
            rangeToClear.clear({ contentsOnly: false }).setBorder(false, false, false, false, false, false);
            Logger.log(`🧹 Menghapus format & isi dari ${sheet.getMaxRows() - lastDataRow} baris di bawah data.`);
          }
            modulFungsiTampilanSheetPenerbit(sheet,true);
          Logger.log(`✅ Selesai. Dihapus ${data.length - dataLolos.length} baris, tersisa ${dataLolos.length}.`);
    }

    function modulHapusBarisDariHasilSeleksi(namaPenerbit) {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hasil Seleksi");
        const data = sheet.getDataRange().getValues();

        const barisUntukHapus = [];

        for (let i = 0; i < data.length; i++) {
          const nama = String(data[i][1] || "").replace(/^\d+\.\s*/, "").trim();
          if (nama.toLowerCase() === namaPenerbit.toLowerCase()) {
            barisUntukHapus.push(i + 1);
          }
        }

        barisUntukHapus.reverse().forEach(row => sheet.deleteRow(row));
    }
  //     ========     Fungsi Utama Hapus Data Menggunakan Filter  1 Sheet   ========
  
  //     ========     Fungsi Utama Hapus Data Menggunakan Filter Semua Sheet    ========
    function modulFungsiHapusDataDenganKriteriaALLSHEET(sheetMulai) {
      const spreadsheet = SpreadsheetApp.getActive();
      const semuaSheet = spreadsheet.getSheets();
      const mulai = sheetMulai + 2;
      const sheetDikecualikan = ['Form Pengadaan', 'Hasil Seleksi', 'Referensi', 'DaftarISBN', 'DaftarUUID'];

      if (mulai < 0 || mulai >= semuaSheet.length) {
        Logger.log("Nomor sheet tidak valid.");
        return;
      }
      semuaSheet.slice(mulai).forEach(sheet => {
        const namaSheet = sheet.getName();
        if (sheetDikecualikan.includes(namaSheet)) return;
        spreadsheet.setActiveSheet(sheet);
        modulFungsiHapusDataDenganKriteria(sheet);
      });
    }
  //     ========     Fungsi Utama Hapus Data Menggunakan Filter Semua Sheet    ========

  //     ========     Fungsi Utama Hapus Data Menggunakan Filter Dengan Batch    ========
    function jalankanSemuaSheetPenerbitBatch_HapusDataDenganKriteria() {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const semuaSheet = ss.getSheets()
        .map(s => s.getName())
        .filter(n => !SHEET_KECUALI.includes(n));

      const props = PropertiesService.getScriptProperties();
      props.setProperty("DAFTAR_SHEET_PENERBIT", JSON.stringify(semuaSheet));
      props.setProperty("TOTAL_SHEET", semuaSheet.length.toString());
      props.setProperty("BATCH_INDEX", "1");

      jalankanBatchBerikutnya_Hapus_Kriteria();
    }
    
    function jalankanBatchBerikutnya_Hapus_Kriteria() {
      const props = PropertiesService.getScriptProperties();
      let daftar = JSON.parse(props.getProperty("DAFTAR_SHEET_PENERBIT") || "[]");
      let batchIndex = parseInt(props.getProperty("BATCH_INDEX") || "1");

      if (daftar.length === 0) {
        kirimNotifikasiSelesaiAkhir_HapusDataDenganKriteria(parseInt(props.getProperty("TOTAL_SHEET") || "0"));
        props.deleteAllProperties();
        hapusSemuaTrigger_HapusDataDenganKriteria();
        Logger.log("✅ Semua sheet selesai diproses Fungsi Hapus Data Menggunakan Kriteria!");
        return;
      }

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const batchSheets = daftar.splice(0, JUMLAH_PER_BATCH); // ambil 5 sheet
      const hasilBatch_HapusDataDenganKriteria = [];

      Logger.log(`▶ Memproses Fungsi Hapus Data Menggunakan Kriteria batch #${batchIndex}: ${batchSheets.join(", ")}`);

      batchSheets.forEach(nama => {
        const sheet = ss.getSheetByName(nama);
        if (!sheet) {
          hasilBatch_HapusDataDenganKriteria.push({ sheet: nama, status: "❌ Tidak ditemukan" });
          return;
        }

        try {
          modulFungsiHapusDataDenganKriteria(sheet);  
      
          hasilBatch_HapusDataDenganKriteria.push({ sheet: nama, status: "✅ Berhasil" });
        } catch (err) {
          hasilBatch_HapusDataDenganKriteria.push({ sheet: nama, status: `⚠️ Gagal: ${err.message}` });
        }
      });

      // kirim notifikasi per batch (berisi hasil sukses/gagal)
      kirimNotifikasiBatch_HapusDataDenganKriteria(hasilBatch_HapusDataDenganKriteria, batchIndex);

      // simpan sisa & increment batch
      props.setProperty("DAFTAR_SHEET_PENERBIT", JSON.stringify(daftar));
      props.setProperty("BATCH_INDEX", (batchIndex + 1).toString());

      if (daftar.length > 0) {
        const sisa = daftar.length;
        const jedaDetik = sisa > 30 ? 20 : sisa > 10 ? 10 : 5;

        hapusSemuaTrigger_HapusDataDenganKriteria();
        ScriptApp.newTrigger("jalankanBatchBerikutnya_Hapus_Kriteria")
          .timeBased()
          .after(jedaDetik * 1000)
          .create();

        Logger.log(`⏱ Menjadwalkan Fungsi Hapus Data Menggunakan Kriteria batch berikutnya (${jedaDetik} detik)...`);
      } else {
        kirimNotifikasiSelesaiAkhir_HapusDataDenganKriteria(parseInt(props.getProperty("TOTAL_SHEET") || "0"));
        props.deleteAllProperties();
        hapusSemuaTrigger_HapusDataDenganKriteria();
        Logger.log("✅ Semua sheet selesai diproses Fungsi Hapus Data Menggunakan Kriteria !");
      }
    }
  //     ========     Fungsi Utama Hapus Data Menggunakan Filter Dengan Batch    ========

  //     ========     Fungsi Utama Mengirim Noftifikasi Ke Telegram    ========
   // region
  /** 🧩 Notifikasi batch ke Telegram */
    function kirimNotifikasiBatch_HapusDataDenganKriteria(hasilBatch, batchIndex) {
      const waktu = new Date().toLocaleString("id-ID", { timeZone: "Asia/Jakarta" });
      const namaSpreadsheet = SpreadsheetApp.getActive().getName();

      const daftarSheet = hasilBatch
        .map(item => `     • *${item.status} * ${item.sheet}`)
        .join('\n');

      const pesan =
`
📘 *Spreadsheet :*
${namaSpreadsheet}

🗑️ *Batch #${batchIndex}-Penghapusan Data Menggunakan Kriteria *

${daftarSheet}

🕒 *Waktu:* ${waktu}

`;

      kirimPesanTelegram_HapusDataDenganKriteria(pesan);
    }
  /** 🧩 Notifikasi batch ke Telegram */

  /** 🧩 Notifikasi selesai semua batch */
    function kirimNotifikasiSelesaiAkhir_HapusDataDenganKriteria(totalSheet) {
      const waktu = new Date().toLocaleString("id-ID", { timeZone: "Asia/Jakarta" });
      const namaSpreadsheet = SpreadsheetApp.getActive().getName();

      const pesan =
`
🎉 *Fungsi Hapus Data Menggunakan Kriteria Selesai*

📘 *Spreadsheet:*
${namaSpreadsheet}

📊 *Total Sheet:* ${totalSheet}

🕒 *Selesai pada:* ${waktu}

`;

      kirimPesanTelegram_HapusDataDenganKriteria(pesan);
    }
  /** 🧩 Notifikasi selesai semua batch */

  /** 📬 Kirim pesan ke Telegram */
    function kirimPesanTelegram_HapusDataDenganKriteria(pesan) {
      const url = `https://api.telegram.org/bot${BOT_TOKEN}/sendMessage`;

      CHAT_IDS.forEach(id => {
        try {
          UrlFetchApp.fetch(url, {
            method: "post",
            contentType: "application/json",
            payload: JSON.stringify({
              chat_id: id,
              text: pesan,
              parse_mode: "Markdown"
            }),
            muteHttpExceptions: true
          });

          Logger.log(`📨 Notifikasi dikirim ke ${id}`);
        } catch (e) {
          Logger.log(`⚠️ Gagal mengirim notifikasi ke ${id}: ${e}`);
        }
      });
    }
  /** 📬 Kirim pesan ke Telegram */

  /** 🧩 Hapus semua trigger batch lama */
    function hapusSemuaTrigger_HapusDataDenganKriteria() {
      ScriptApp.getProjectTriggers().forEach(t => {
        if (t.getHandlerFunction() === "jalankanBatchBerikutnya_Hapus_Kriteria") {
          ScriptApp.deleteTrigger(t);
        }
      });
    }
  /** 🧩 Hapus semua trigger batch lama */

   // endregion 
  //     ========     Fungsi Utama Mengirim Noftifikasi Ke Telegram    ========

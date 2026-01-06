  //     ========     MengHapus Pembelian Sebelumnya berdasar UUID 1 Sheet     ========
    function HapusPembeliansebelumdenganUUID(){
      ModulHapusSudahdibeli("UUID");
    }
  //     ========     MengHapus Pembelian Sebelumnya berdasar UUID 1 Sheet     ======== 

  //     ========     MengHapus Pembelian Sebelumnya berdasar UUID Semua Sheet     ========
    function HapusPembeliansebelumdenganUUIDALLSHEET() {
    //PERLU DIRUBAH ============================================================================== 
    const sheetMulai = 0; 
    //PERLU DIRUBAH ==============================================================================
      ModulHapusSudahdibeliALLSHEET("UUID",sheetMulai);
    }
  //     ========     MengHapus Pembelian Sebelumnya berdasar UUID Semua Sheet     ========  

  //     ========     MengHapus Pembelian Sebelumnya berdasar ISBN & e-ISBN 1 Sheet     ========
    function HapusPembeliansebelumdenganISBN(){
      ModulHapusSudahdibeli("ISBN");
    }
  //     ========     MengHapus Pembelian Sebelumnya berdasar ISBN & e-ISBN 1 Sheet     ========  

  //     ========     MengHapus Pembelian Sebelumnya berdasar ISBN & e-ISBN Semua Sheet     ========
    function HapusPembeliansebelumdenganISBNALLSHEET() {
    //PERLU DIRUBAH ============================================================================== 
    const sheetMulai = 0; 
    //PERLU DIRUBAH ==============================================================================
    
      ModulHapusSudahdibeliALLSHEET("ISBN",sheetMulai);
    }
  //     ========     MengHapus Pembelian Sebelumnya berdasar ISBN & e-ISBN Semua Sheet     ========  

  //     ========     MengHapus Pembelian Sebelumnya Dengan Batch UUID    ========
    function HapusPembelian_UUID_BATCH() {
      jalankanSemuaSheetPenerbitBatch_HapusPembelian("UUID");
    }
  //     ========     MengHapus Pembelian Sebelumnya Dengan Batch UUID    ========

 //     ========     MengHapus Pembelian Sebelumnya Dengan Batch ISBN    ========
    function HapusPembelian_ISBN_BATCH() {
      jalankanSemuaSheetPenerbitBatch_HapusPembelian("ISBN");
    }
 //     ========     MengHapus Pembelian Sebelumnya Dengan Batch ISBN    ========






  //     ========     Fungsi Utama MengHapus Pembelian Sebelumnya 1 Sheet    ========
    function ModulHapusSudahdibeli(type, sheetData) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!sheetData) sheetData = ss.getActiveSheet();
      const nama = sheetData.getName();

      // ✅ Skip sheet non-penerbit
      const sheetExcluded = ['Form Pengadaan', 'Hasil Seleksi', 'Referensi', 'DaftarISBN', 'DaftarUUID'];
      if (sheetExcluded.includes(nama)) {
        Logger.log('➡️ Sheet dilewati: ' + nama);
        return;
      }

      Logger.log('▶ Menjalankan HapusPembelianSebelumnya pada: ' + nama);

      const headerRow = 9;
      const startRow = 10;
      const lastRow = sheetData.getLastRow();
      const numRows = lastRow - startRow + 1;

      if (numRows < 1) {
        Logger.log("⚠️ Tidak ada data untuk diproses.");
        return;
      }

      // ✅ Ambil header
      const headers = sheetData.getRange(headerRow, 1, 1, sheetData.getLastColumn())
        .getValues()[0]
        .map(h => String(h).toLowerCase().trim());
      const data = sheetData.getRange(startRow, 1, numRows, sheetData.getLastColumn()).getValues();

      let filteredData = [];

      // ===== Hapus Berdasarkan UUID =====
      if (type === "UUID") {
        const sheetUUID = ss.getSheetByName("DaftarUUID");
        if (!sheetUUID) throw new Error("Sheet 'DaftarUUID' tidak ditemukan.");
        const uuidRef = sheetUUID.getRange("A:A").getValues().flat().filter(Boolean);
        const uuidSet = new Set(uuidRef.map(val => String(val).trim()));

        const uuidIndex = headers.findIndex(h => h.includes("uuid"));
        if (uuidIndex === -1) throw new Error(`Kolom UUID tidak ditemukan di baris ${headerRow}`);

        filteredData = data.filter(row => {
          const uuid = String(row[uuidIndex]).trim();
          return !uuidSet.has(uuid);
        });
      }

      // ===== Hapus Berdasarkan ISBN =====
      if (type === "ISBN") {
        const sheetISBN = ss.getSheetByName("DaftarISBN");
        if (!sheetISBN) throw new Error("Sheet 'DaftarISBN' tidak ditemukan.");
        const rawISBNData = sheetISBN.getRange("A:B").getValues();
        const isbnSet = new Set();
        rawISBNData.forEach(([isbnCetak, isbnElektronik]) => {
          const clean1 = String(isbnCetak).replace(/\D/g, "");
          const clean2 = String(isbnElektronik).replace(/\D/g, "");
          if (clean1) isbnSet.add(clean1);
          if (clean2) isbnSet.add(clean2);
        });

        const isbnCetakIndex = headers.findIndex(h => h.includes("isbn cetak"));
        const isbnElektronikIndex = headers.findIndex(h => h.includes("isbn elektronik"));
        if (isbnCetakIndex === -1 && isbnElektronikIndex === -1)
          throw new Error("Kolom ISBN tidak ditemukan di baris " + headerRow);

        filteredData = data.filter(row => {
          const valCetak = isbnCetakIndex !== -1 ? String(row[isbnCetakIndex]).replace(/\D/g, "") : "";
          const valElektronik = isbnElektronikIndex !== -1 ? String(row[isbnElektronikIndex]).replace(/\D/g, "") : "";
          return !(isbnSet.has(valCetak) || isbnSet.has(valElektronik));
        });
      }

      // Kosongkan area lama
      sheetData.getRange(startRow, 1, numRows, sheetData.getLastColumn()).clearContent();

      // ✅ Tulis ulang data hasil filter
      if (filteredData.length > 0) {
        sheetData.getRange(startRow, 1, filteredData.length, sheetData.getLastColumn())
          .setValues(filteredData);
      }

      const deletedCount = data.length - filteredData.length;

      // ✅ Jika sheet kosong, hapus
      if (filteredData.length === 0) {
        const namaPenerbit = sheetData.getName().replace(/^\d+\.\s*/, "").trim();
        ss.deleteSheet(sheetData);
        ModulHapusBarisDariHasilSeleksi(namaPenerbit);
        Logger.log(`🗑 Sheet '${nama}' dihapus karena kosong.`);
        return;
      }

      Logger.log(`✅ [${type}] Selesai. Total: ${data.length}, dihapus: ${deletedCount}, tersisa: ${filteredData.length}`);

      // ✅ Atur tampilan khusus sheet ini
      modulFungsiTampilanSheetPenerbit(sheetData,true);
    }

    function ModulHapusBarisDariHasilSeleksi(namaPenerbit) {
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

  //     ========     Fungsi Utama MengHapus Pembelian Sebelumnya 1 Sheet    ========  

  //     ========     Fungsi Utama MengHapus Pembelian Sebelumnya Semua Sheet    ========  
    function ModulHapusSudahdibeliALLSHEET(mode, sheetMulai) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const semuaSheet = ss.getSheets();
      const mulai = sheetMulai + 2;
      const sheetExcluded = [
        'Form Pengadaan',
        'Hasil Seleksi',
        'Referensi',
        'DaftarISBN',
        'DaftarUUID'
      ];

      if (sheetMulai < 0 || sheetMulai >= semuaSheet.length) {
        Logger.log("❌ Nomor sheet tidak valid.");
        return;
      }

      let totalDiproses = 0;
      semuaSheet.slice(mulai).forEach(sheet => {
        const nama = sheet.getName();
        if (sheetExcluded.includes(nama)) {
          Logger.log("➡️ Skip sheet: " + nama);
          return;
        }

        try {
          Logger.log(`▶ Menjalankan HapusPembelian pada: ${nama}`);
          ModulHapusSudahdibeli(mode, sheet);
          totalDiproses++;
        } catch (err) {
          Logger.log(`⚠️ Gagal di sheet ${nama}: ${err.message}`);
        }
      });

      Logger.log(`✅ Selesai Hapus Data di ${totalDiproses} sheet (mode: ${mode})`);
    }
  //     ========     Fungsi Utama MengHapus Pembelian Sebelumnya Semua Sheet    ========
  
  //     ========     Fungsi Utama MengHapus Pembelian Sebelumnya Dengan Batch    ========
    function jalankanSemuaSheetPenerbitBatch_HapusPembelian(mode = "UUID") {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const semuaSheet = ss.getSheets()
        .map(s => s.getName())
        .filter(n => !SHEET_KECUALI.includes(n));
      const props = PropertiesService.getScriptProperties();
      props.setProperty("MODE_BATCH", mode);
      props.setProperty("DAFTAR_SHEET_PENERBIT", JSON.stringify(semuaSheet));
      props.setProperty("TOTAL_SHEET", semuaSheet.length.toString());
      props.setProperty("BATCH_INDEX", "1");

      jalankanBatchBerikutnya_HapusPembelian();
    }

    function jalankanBatchBerikutnya_HapusPembelian() {
      const props = PropertiesService.getScriptProperties();
      const mode = props.getProperty("MODE_BATCH");
      let daftar = JSON.parse(props.getProperty("DAFTAR_SHEET_PENERBIT") || "[]");
      let batchIndex = parseInt(props.getProperty("BATCH_INDEX") || "1");

      if (daftar.length === 0) {
        kirimNotifikasiSelesaiAkhir_HapusPembelian(parseInt(props.getProperty("TOTAL_SHEET") || "0"));
        props.deleteAllProperties();
        hapusSemuaTrigger_HapusPembelian();
        Logger.log("✅ Semua sheet selesai diproses Fungsi MengHapus Pembelian Sebelumnya!");
        return;
      }

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const batchSheets = daftar.splice(0, JUMLAH_PER_BATCH);
      const hasilBatch_HapusPembelian = [];

      Logger.log(`▶ Memproses Fungsi MengHapus Pembelian Sebelumnya batch #${batchIndex}: ${batchSheets.join(", ")}`);

      batchSheets.forEach(nama => {
        const sheet = ss.getSheetByName(nama);
        if (!sheet) {
          hasilBatch_HapusPembelian.push({ sheet: nama, status: "❌ Tidak ditemukan" });
          return;
        }

        try {
          ModulHapusSudahdibeli(mode, sheet);
          hasilBatch_HapusPembelian.push({ sheet: nama, status: "✅ Berhasil" });
        } catch (err) {
          hasilBatch_HapusPembelian.push({ sheet: nama, status: `⚠️ Gagal: ${err.message}` });
        }
      });

      kirimNotifikasiBatch_HapusPembelian(hasilBatch_HapusPembelian, batchIndex);
      props.setProperty("DAFTAR_SHEET_PENERBIT", JSON.stringify(daftar));
      props.setProperty("BATCH_INDEX", (batchIndex + 1).toString());

      if (daftar.length > 0) {
        const sisa = daftar.length;
        const jedaDetik = sisa > 30 ? 20 : sisa > 10 ? 10 : 5;
        hapusSemuaTrigger_HapusPembelian();
        ScriptApp.newTrigger("jalankanBatchBerikutnya_HapusPembelian")
          .timeBased()
          .after(jedaDetik * 1000)
          .create();
        Logger.log(`⏱ Menjadwalkan Fungsi MengHapus Pembelian Sebelumnya batch berikutnya (${jedaDetik} detik)...`);
      } else {
        kirimNotifikasiSelesaiAkhir_HapusPembelian(parseInt(props.getProperty("TOTAL_SHEET") || "0"));
        props.deleteAllProperties();
        hapusSemuaTrigger_HapusPembelian();
        Logger.log("✅ Semua sheet selesai diproses Fungsi MengHapus Pembelian Sebelumnya!");
      }
    }
  //     ========     Fungsi Utama MengHapus Pembelian Sebelumnya Dengan Batch    ========  

//     ========     Fungsi Utama Mengirim Noftifikasi Ke Telegram    ========
   // region

  /** 🧩 Notifikasi batch ke Telegram */
      function kirimNotifikasiBatch_HapusPembelian(hasilBatch_HapusPembelian, batchIndex) {
        const waktu = new Date().toLocaleString("id-ID", { timeZone: "Asia/Jakarta" });
        const namaSpreadsheet = SpreadsheetApp.getActiveSpreadsheet().getName();
        const props = PropertiesService.getScriptProperties();
        const mode = props.getProperty("MODE_BATCH"); 
        const daftarSheet = hasilBatch_HapusPembelian
        .map(item => `     • *${item.status}*  ${item.sheet}`)
        .join('\n');

        const pesan =
`
📘 *Spreadsheet:* 
${namaSpreadsheet}

✅ *Batch #${batchIndex} - Penghapusan Data Pembelian Dengan ${mode} Selesai ! *

${daftarSheet}

🕒 *Waktu selesai:* ${waktu}

`;

        kirimPesanTelegram_HapusPembelian(pesan);
      }
  /** 🧩 Notifikasi batch ke Telegram */

  /** 🧩 Notifikasi selesai semua batch */
    function kirimNotifikasiSelesaiAkhir_HapusPembelian(totalSheet) {
      const waktu = new Date().toLocaleString("id-ID", { timeZone: "Asia/Jakarta" });
      const namaSpreadsheet = SpreadsheetApp.getActiveSpreadsheet().getName();
      const props = PropertiesService.getScriptProperties();
      const mode = props.getProperty("MODE_BATCH"); 
      const pesan =
`
🎉 *Fungsi Hapus Data Pembelian Dengan ${mode} Selesai !*

📘 *Spreadsheet:* 
${namaSpreadsheet}

📊 Total sheet: *${totalSheet}*

🕒 Waktu selesai: ${waktu}

`;

      kirimPesanTelegram_HapusPembelian(pesan);
    }
  /** 🧩 Notifikasi selesai semua batch */

  /** 📬 Kirim pesan ke Telegram */
    function kirimPesanTelegram_HapusPembelian(pesan) {
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
          Logger.log(`⚠️ Gagal kirim notifikasi ke ${id}: ${e}`);
        }
      });
    }
  /** 📬 Kirim pesan ke Telegram */

  /** 🧩 Hapus semua trigger batch lama */
    function hapusSemuaTrigger_HapusPembelian() {
      ScriptApp.getProjectTriggers().forEach(t => {
        if (t.getHandlerFunction() === "jalankanBatchBerikutnya_HapusPembelian") {
          ScriptApp.deleteTrigger(t);
        }
      });
    }
  /** 🧩 Hapus semua trigger batch lama */  
  // endregion 
//     ========     Fungsi Utama Mengirim Noftifikasi Ke Telegram    ========
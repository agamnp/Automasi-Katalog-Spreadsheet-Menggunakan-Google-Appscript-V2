//     ========     Atur Tampilan 1 Sheet    ========
  function AturTampilan1Sheet(sheet){
    hapusSemuaTrigger_Tampilan_Sheet_Penerbit(); // ✅ optional safety
    modulFungsiTampilanSheetPenerbit(sheet); }
//     ========     Atur Tampilan 1 Sheet    ========     

//     ========     Atur Tampilan Semua Sheet    ========
  function AturTampilanSemuaSheet(){
    //PERLU DIRUBAH ============================================================================== 
      const sheetMulai = 0; 
    //PERLU DIRUBAH ==============================================================================
    modulFungsiTampilanSheetPenerbitALLSHEET(sheetMulai);  }
//     ========     Atur Tampilan Semua Sheet    ========

//     ========     Fungsi Mengatur Tampilan Sheet Penerbit Batch     ========
   function AturTampilanSheetBatch(){
    jalankanSemuaSheetPenerbitBatch_Tampilan_Sheet_Penerbit();
    }
//     ========     Fungsi Mengatur Tampilan Sheet Penerbit Batch     ========






//     ========     Fungsi Utama Mengatur Tampilan Sheet Penerbit    ========
  function modulFungsiTampilanSheetPenerbit(sheet, fromBatch = false) {
    // kalau dijalankan manual, hapus trigger
  if (!fromBatch) {
    try {
      hapusSemuaTrigger_Tampilan_Sheet_Penerbit();
    } catch (e) {
      Logger.log("⚠️ Tidak bisa hapus trigger: " + e);
    }
  }

  if (!sheet) {
    try { sheet = SpreadsheetApp.getActiveSheet(); }
    catch (e) {
      Logger.log("❌ Tidak ada sheet aktif. Fungsi dihentikan.");
      return;
    }
  }
    // ✅ Memeriksa Sheet Yang Aktif
      if (!sheet) {
      try { 
      sheet = SpreadsheetApp.getActiveSheet(); 
      } catch (e) {
      Logger.log("❌ Tidak ada sheet aktif. Fungsi dihentikan.");
      return;
      }} if (!sheet) return;
    //  

     // Pastikan sheet valid
  if (!sheet || typeof sheet.getName !== "function") {
  Logger.log("❌ Parameter sheet tidak valid, fungsi dihentikan.");
  return;
  }

    // ✅ Variabel yang Banyak Di Gunakan
      const nama = sheet.getSheetName();
      const startRow = 10;
      const spreadsheet = sheet.getParent();
      const lastRow = sheet.getLastRow();

      Logger.log('▶ Menjalankan FungsiTampilanSheetPenerbit pada: ' + nama);

    //

    // ✅ Melewati Sheet Form Pengadaan , Hasil Seleksi , Referensi , DaftarISBN , DaftarUUID
      const sheetExcluded = ['Form Pengadaan', 'Hasil Seleksi', 'Referensi', 'DaftarISBN', 'DaftarUUID'];
      if (sheetExcluded.includes(nama)) {
        Logger.log('Sheet dilewati: ' + nama);
        return;
      }
    //  

    // ✅ Melepas Freeze & Filter jika ada
      sheet.setFrozenRows(0);
      sheet.setFrozenColumns(0);
      const filter = sheet.getFilter();
      if (filter) filter.remove();
    //

    // ✅ Menghapusformat seluruh area data dulu
      if (lastRow >= startRow) {
        const dataRange = sheet.getRange(`A10:AC${lastRow}`);
        dataRange.clearFormat();                          // hapus format  
      }
    //   

    //✅ Mengatur ukuran kolom
      const ukuranKolom = [
        44, 119, 369, 129, 127, 134, 124, 125, 109, 109,109,
        100, 100, 100, 100, 100, 100, 100, 100, 100, 100,
        100, 100, 100, 100, 150, 80, 115, 266
      ];

      ukuranKolom.forEach((width, i) => {
        sheet.setColumnWidth(i + 1, width);
      });
    // 
   
    // ✅ Mengisi Rumus Tabel Ketersediaan Katalog dan Hasil Seleksi
      const formulaCells = [
        ['G2', `=COUNTA(C10:C${lastRow})`],
        ['G3', `=SUM(Z10:Z${lastRow})`],
        ['G4', `=AVERAGE(Z10:Z${lastRow})`],
        ['J2', `=COUNTA(AA10:AA${lastRow})`],
        ['J3', `=SUM(AA10:AA${lastRow})`],
        ['J4', `=SUM(AB10:AB${lastRow})`],
        ['J5', `=AVERAGEIF(AA10:AA${lastRow}; ">0"; AB10:AB${lastRow})`]
      ];
      formulaCells.forEach(([cell, formula]) => sheet.getRange(cell).setFormula(formula));
    //

    // ✅ Mengisi autofill nomor Urut , Preview Konten ,Referensi,Sub Referensi ,Total Harga
      function modulclearAndAutoFillColumn(colLetter, formulaOrValue) {
          const col = sheet.getRange(colLetter + '1').getColumn();
          const range = sheet.getRange(startRow, col, lastRow - startRow + 1);
          range.clear({ contentsOnly: true, skipFilteredRows: true });

          const firstCell = sheet.getRange(startRow, col);
          if (formulaOrValue.startsWith('=')) {
            firstCell.setFormula(formulaOrValue);
          } else {
            firstCell.setValue(formulaOrValue);
          }
          if (lastRow > startRow) {
            firstCell.autoFill(
              sheet.getRange(startRow, col, lastRow - startRow + 1),
              SpreadsheetApp.AutoFillSeries.ALTERNATE_SERIES
            );
          }
      }
      modulclearAndAutoFillColumn('A', '1');
      modulclearAndAutoFillColumn('B', '=HYPERLINK("https://mocostore.moco.co.id/catalog/"&AC10;"Klik Disini")');
      modulclearAndAutoFillColumn('AB', '=Z10*AA10');
      modulclearAndAutoFillColumn('K', '=IFERROR(VLOOKUP(J10; Referensi!A:B; 2; FALSE); "")');
      modulclearAndAutoFillColumn('I','=IFERROR(VLOOKUP(LEFT(J10;3); Referensi!A:B; 2; FALSE); "")')
    //

    // ✅ Hapus baris kosong 
      // === HAPUS BARIS KOSONG SETELAH DATA (berdasarkan kolom C) ===
    
      const totalRowshapus = sheet.getMaxRows();
      const values = sheet.getRange(`C10:C${totalRowshapus}`).getValues();
      let lastDataRowHapus = 10;

      // cari baris terakhir yang berisi data (mulai dari bawah)
      for (let i = values.length - 1; i >= 0; i--) {
        if (values[i][0] !== "" && values[i][0] !== null) {
          lastDataRowHapus = i + 10;  // offset, karena mulai dari row 10
          break;
        }
      }

      // hapus semua baris setelah baris terakhir yang berisi data
      if (lastDataRowHapus < totalRowshapus) {
        const jumlahDihapus = totalRowshapus - lastDataRowHapus;
        sheet.deleteRows(lastDataRowHapus + 1, jumlahDihapus);
        Logger.log(`🗑 Menghapus ${jumlahDihapus} baris kosong (mulai row ${lastDataRowHapus + 1}).`);
      }
      
    //

    // ✅ Mengatur Format border , Font , alignment
        const borderStyle = SpreadsheetApp.BorderStyle.SOLID;
        sheet.getRange('F1:J5').setBorder(false, false, false, false, false, false);
        sheet.getRangeList(['F1:G4', 'I1:J5']).setBorder(true, true, true, true, true, true, '#000', borderStyle);
        sheet.getRange(`A10:AC${lastRow}`).setBorder(true, true, true, true, true, true, '#000', borderStyle);
        sheet.getRange('F1:J5').setHorizontalAlignment('center').setVerticalAlignment('middle');
        sheet.getRangeList([
          `A10:B${lastRow}`, `F10:G${lastRow}`, `J10:J${lastRow}`, `Y10:AC${lastRow}`
        ]).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontStyle('normal').setFontWeight('normal');
        sheet.getRangeList([
          `C10:E${lastRow}`, `H10:I${lastRow}`, `K10:X${lastRow}`
        ]).setHorizontalAlignment('left').setVerticalAlignment('middle').setFontStyle('normal').setFontWeight('normal');
        sheet.getRangeList(['J4', 'G3', 'G4']).setNumberFormat('[$Rp-421] #,##0');
    //

    // ✅ Update alternating color range
      function updateAlternatingColor(sheet, lastRow) {
        sheet.getBandings().forEach(b => b.remove());
        const banding = sheet.getRange(`A9:AC${lastRow}`)
                            .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);

      // banding.setHeaderRowColor(null); // biar baris 9 tetap polos
          }
      updateAlternatingColor(sheet, lastRow);
    //  

    // ✅ Pasang ulang freeze
      sheet.setFrozenColumns(10)
      sheet.setFrozenRows(9);
    //

    // ✅ Pasang filter
      const dataRange = sheet.getRange(`A9:AC${lastRow}`);
      if (!dataRange.getFilter()) dataRange.createFilter();
    //

    // ✅ Mengatur tingg Baris
      if (lastRow >= startRow) {
        sheet.setRowHeightsForced(startRow, lastRow - startRow + 1, 20);
      }
    //

    // ✅ Ganti named range
      const cleanNamerange = nama.replace(/[0-9().\-]/g, '').replace(/\s/g, '');
      spreadsheet.setNamedRange(cleanNamerange, sheet.getRange(`J10:J${lastRow}`));
    //

    // ✅ Ganti nama sheet
      const cleanNamesheet = nama.replace(/[0-9.]/g, '');
      const sheetIndex = sheet.getIndex() - 3;
      const newName = `${sheetIndex}.${cleanNamesheet}`;
      if (spreadsheet.getSheets().every(s => s.getName() !== newName)) {
        sheet.setName(newName);
      }
    //

    Logger.log('✅ Selesai Menjalankan FungsiTampilanSheetPenerbit pada: ' + nama);
  }
  
//     ========     Fungsi Utama Mengatur Tampilan Sheet Penerbit    ========

//     ========     Fungsi Utama Mengatur Tampilan Sheet Penerbit semua Sheet     ========
  function modulFungsiTampilanSheetPenerbitALLSHEET(sheetMulai) {
    const spreadsheet = SpreadsheetApp.getActive();
    const semuaSheet = spreadsheet.getSheets();
    const mulai = sheetMulai + 2;
    const sheetDikecualikan = ['Form Pengadaan', 'Hasil Seleksi', 'Referensi', 'DaftarISBN', 'DaftarUUID'];
    semuaSheet.slice(mulai).forEach(sheet => {
      if (!sheetDikecualikan.includes(sheet.getName())) {
        spreadsheet.setActiveSheet(sheet); // 🔹 pindah ke sheet yang diproses
        modulFungsiTampilanSheetPenerbit(sheet);
      }
    });
  }
//     ========     Fungsi Utama Mengatur Tampilan Sheet Penerbit semua Sheet     ========

//     ========     Fungsi Utama Mengatur Tampilan Sheet Penerbit Dengan Batch    ========
  function jalankanSemuaSheetPenerbitBatch_Tampilan_Sheet_Penerbit() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const semuaSheet = ss.getSheets()
      .map(s => s.getName())
      .filter(n => !SHEET_KECUALI.includes(n));

    const props = PropertiesService.getScriptProperties();
    props.setProperty("DAFTAR_SHEET_PENERBIT", JSON.stringify(semuaSheet));
    props.setProperty("TOTAL_SHEET", semuaSheet.length.toString());
    props.setProperty("BATCH_INDEX", "1");

    jalankanBatchBerikutnya_Tampilan_Sheet_Penerbit();
  }

  function jalankanBatchBerikutnya_Tampilan_Sheet_Penerbit() {
    const props = PropertiesService.getScriptProperties();
    let daftar = JSON.parse(props.getProperty("DAFTAR_SHEET_PENERBIT") || "[]");
    let batchIndex = parseInt(props.getProperty("BATCH_INDEX") || "1");

    if (daftar.length === 0) {
      kirimNotifikasiSelesaiAkhir_Tampilan_Sheet_Penerbit(parseInt(props.getProperty("TOTAL_SHEET") || "0"));
      props.deleteAllProperties();
      hapusSemuaTrigger_Tampilan_Sheet_Penerbit();
      Logger.log("✅ Semua sheet selesai diproses Fungsi Tampilan Sheet Penerbit !");
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const batchSheets = daftar.splice(0, JUMLAH_PER_BATCH); // ambil 5 sheet
    const hasilBatchTampilan_Sheet_Penerbit = [];
    Logger.log(`▶ Memproses Fungsi Tampilan Sheet Penerbit ! batch #${batchIndex}: ${batchSheets.join(", ")}`);
    batchSheets.forEach(nama => {
      const sheet = ss.getSheetByName(nama);
      if (!sheet) {
        hasilBatchTampilan_Sheet_Penerbit.push({ sheet: nama, status: "❌ Tidak ditemukan" });
        return;
      }
      try {
        modulFungsiTampilanSheetPenerbit(sheet,true);  //===================================== Fungsi Utama
    
        hasilBatchTampilan_Sheet_Penerbit.push({ sheet: nama, status: "✅ Berhasil" });
      } catch (err) {
        hasilBatchTampilan_Sheet_Penerbit.push({ sheet: nama, status: `⚠️ Gagal: ${err.message}` });
      }
    });

    // kirim notifikasi per batch (berisi hasil sukses/gagal)
    kirimNotifikasiBatch_Tampilan_Sheet_Penerbit(hasilBatchTampilan_Sheet_Penerbit, batchIndex);

    // simpan sisa & increment batch
    props.setProperty("DAFTAR_SHEET_PENERBIT", JSON.stringify(daftar));
    props.setProperty("BATCH_INDEX", (batchIndex + 1).toString());

    if (daftar.length > 0) {
      const sisa = daftar.length;
      const jedaDetik = sisa > 30 ? 20 : sisa > 10 ? 10 : 5;

      hapusSemuaTrigger_Tampilan_Sheet_Penerbit();
      ScriptApp.newTrigger("jalankanBatchBerikutnya_Tampilan_Sheet_Penerbit")
        .timeBased()
        .after(jedaDetik * 1000)
        .create();

      Logger.log(`⏱ Menjadwalkan Fungsi Tampilan Sheet Penerbit ! batch berikutnya (${jedaDetik} detik)...`);
    } else {
      kirimNotifikasiSelesaiAkhir_Tampilan_Sheet_Penerbit(parseInt(props.getProperty("TOTAL_SHEET") || "0"));
      props.deleteAllProperties();
      hapusSemuaTrigger_Tampilan_Sheet_Penerbit();
      Logger.log("✅ Semua sheet selesai diproses Fungsi Tampilan Sheet Penerbit !");
    }
  }
//     ========     Fungsi Utama Mengatur Tampilan Sheet Penerbit Dengan Batch    ========

//     ========     Fungsi Utama Mengirim Noftifikasi Ke Telegram    ========
  // region
  /** 🧩 Notifikasi batch ke Telegram */
    function kirimNotifikasiBatch_Tampilan_Sheet_Penerbit(hasilBatch, batchIndex) {
      const waktu = new Date().toLocaleString("id-ID", { timeZone: "Asia/Jakarta" });
      const namaSpreadsheet = SpreadsheetApp.getActive().getName();
      const daftarSheet = hasilBatch
        .map(item => `     • *${item.status}*  ${item.sheet}`)
        .join('\n');

const pesan =
`
📘 *Spreadsheet :* 
${namaSpreadsheet}

✅ *Batch #${batchIndex} - Fungsi Tampilan Sheet Selesai !*

${daftarSheet}

🕒 *Waktu:* ${waktu}

`;

      kirimPesanTelegram_Tampilan_Sheet_Penerbit(pesan);
    }
  /** 🧩 Notifikasi batch ke Telegram */

  /** 🧩 Notifikasi ketika seluruh batch selesai */
    function kirimNotifikasiSelesaiAkhir_Tampilan_Sheet_Penerbit(totalSheet) {
      const waktu = new Date().toLocaleString("id-ID", { timeZone: "Asia/Jakarta" });
      const namaSpreadsheet = SpreadsheetApp.getActive().getName();

      const pesan =
`
🎉 *Fungsi Tampilan Sheet Semua Sheet Selesai !*

📘 *Spreadsheet :* 
${namaSpreadsheet}

📊 *Total Sheet :* ${totalSheet}

🕒 *Selesai pada :* ${waktu}

`;

      kirimPesanTelegram_Tampilan_Sheet_Penerbit(pesan);
    }
  /** 🧩 Notifikasi ketika seluruh batch selesai */

  /** 📬 Kirim pesan ke Telegram */
    function kirimPesanTelegram_Tampilan_Sheet_Penerbit(pesan) {
      const url = `https://api.telegram.org/bot${BOT_TOKEN}/sendMessage`;

      CHAT_IDS.forEach(chatId => {
        try {
          UrlFetchApp.fetch(url, {
            method: "post",
            contentType: "application/json",
            payload: JSON.stringify({
              chat_id: chatId,
              text: pesan,
              parse_mode: "Markdown"
            }),
            muteHttpExceptions: true
          });

          Logger.log(`📨 Notifikasi terkirim ke ${chatId}`);
        } catch (error) {
          Logger.log(`⚠️ Gagal mengirim notifikasi ke ${chatId}: ${error}`);
        }
      });
    }
  /** 📬 Kirim pesan ke Telegram */ 

   /** 🧩 Hapus semua trigger batch lama */
    function hapusSemuaTrigger_HapusPembelian() {
      ScriptApp.getProjectTriggers().forEach(t => {
        if (t.getHandlerFunction() === "jalankanBatchBerikutnya_Tampilan_Sheet_Penerbit") {
          ScriptApp.deleteTrigger(t);
        }
      });
    }
  /** 🧩 Hapus semua trigger batch lama */ 
// endregion 
//     ========     Fungsi Utama Mengirim Noftifikasi Ke Telegram    ========
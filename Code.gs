function onOpen() {
  SpreadsheetApp.getUi().createMenu('Dashboard SIPARPAS')
    .addItem('Buka Dashboard (Popup)', 'showModal')
    .addItem('Buka di Tab Baru (Web App)', 'openInNewTab')
    .addToUi();
}

// ponytail: doGet enables Apps Script deployment as a standalone web app URL
function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Dashboard SIPARPAS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function showModal() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createTemplateFromFile('Index').evaluate()
      .setTitle('Dashboard SIPARPAS').setWidth(1280).setHeight(850),
    'Dashboard SIPARPAS Pasaman'
  );
}

// ponytail: opens web app in new browser tab via lightweight modal trigger
function openInNewTab() {
  var url = ScriptApp.getService().getUrl();
  if (!url) {
    SpreadsheetApp.getUi().alert(
      'Web App Belum Di-deploy',
      'Silakan klik "Deploy" -> "New deployment" -> pilih tipe "Web app" (Execute as: Me, Who has access: Anyone) agar link tab baru aktif.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  var html = '<script>window.open("' + url + '", "_blank"); google.script.host.close();</script>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(200).setHeight(60), 'Membuka Tab Baru...');
}

// ID Spreadsheet Google Form Responses 1
var SPREADSHEET_ID = "1DjFcM9yoQ8Lr8QTlsSmaEWl84TJyYsxo2fU0wGDaMTw";

// ponytail: automatic column frequency analyzer to support dynamic chart selection for any header
function getDashboardData() {
  var ss;
  try {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  var sheet = ss.getSheetByName("Form Responses 1") || ss.getSheetByName("Formulir Respon 1") || ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();
  
  if (!data || data.length === 0) {
    return { summary: {}, rawData: [], headers: [], columnStats: {} };
  }

  var headers = data[0];
  var rows = [];
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() !== '') {
      rows.push(data[i]);
    }
  }
  
  var summary = {
    total: rows.length,
    kategori: {},
    kecamatan: {},
    operasionalYa: 0,
    pengunjungDomestik: 0,
    pengunjungMancanegara: 0,
    pendapatanTiket: 0,
    kamarTersedia: 0,
    kamarTerisi: 0
  };

  var columnStats = {};
  for (var c = 0; c < headers.length; c++) {
    var colName = headers[c].toString().trim();
    if (colName) {
      columnStats[colName] = {};
    }
  }

  var idxKecamatan = -1;
  var idxOperasional = -1;
  var idxJenisUsaha = -1;
  var idxPengunjungDalam = -1;
  var idxPengunjungLuar = -1;
  var idxPendapatan = -1;
  var idxKamarTersedia = -1;
  var idxKamarTerisi = -1;

  for (var h = 0; h < headers.length; h++) {
    var hdr = headers[h].toString().toLowerCase().trim();
    if (idxKecamatan === -1 && (hdr.indexOf("kecamatan") !== -1 || hdr.indexOf("5. di kecamatan") !== -1)) idxKecamatan = h;
    if (idxOperasional === -1 && (hdr.indexOf("beroperasi") !== -1 || hdr.indexOf("6. apakah usaha") !== -1)) idxOperasional = h;
    if (idxJenisUsaha === -1 && (hdr.indexOf("jenis usaha") !== -1 || hdr.indexOf("jenis destinasi") !== -1 || hdr.indexOf("jenis akomodasi") !== -1)) idxJenisUsaha = h;
    if (idxPengunjungDalam === -1 && hdr.indexOf("dalam negeri") !== -1) idxPengunjungDalam = h;
    if (idxPengunjungLuar === -1 && hdr.indexOf("luar negeri") !== -1) idxPengunjungLuar = h;
    if (idxPendapatan === -1 && (hdr.indexOf("pendapatan") !== -1 || hdr.indexOf("retribusi") !== -1)) idxPendapatan = h;
    if (idxKamarTersedia === -1 && (hdr.indexOf("jumlah kamar tersedia") !== -1 || hdr.indexOf("seluruh kamar") !== -1)) idxKamarTersedia = h;
    if (idxKamarTerisi === -1 && (hdr.indexOf("kamar yang terjual") !== -1 || hdr.indexOf("terjual") !== -1)) idxKamarTerisi = h;
  }

  var processedData = [];

  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];
    var obj = {};

    for (var col = 0; col < headers.length; col++) {
      var headerName = headers[col].toString().trim();
      var rawVal = row[col];
      var strVal = (rawVal !== null && rawVal !== undefined) ? rawVal.toString().trim() : '';

      obj[headerName] = strVal || '-';

      // Build column statistics for charts
      if (headerName && strVal) {
        // truncate long responses for clean charting
        var chartLabel = strVal.length > 50 ? strVal.substring(0, 47) + '...' : strVal;
        if (!columnStats[headerName][chartLabel]) {
          columnStats[headerName][chartLabel] = 0;
        }
        columnStats[headerName][chartLabel]++;
      }
    }

    // Specific KPI aggregations
    var jenis = idxJenisUsaha !== -1 ? row[idxJenisUsaha] : row[9];
    jenis = (jenis && jenis.toString().trim()) ? jenis.toString().trim() : 'Lainnya';
    if (!summary.kategori[jenis]) summary.kategori[jenis] = 0;
    summary.kategori[jenis]++;

    var kec = idxKecamatan !== -1 ? row[idxKecamatan] : row[4];
    kec = (kec && kec.toString().trim()) ? kec.toString().trim() : 'Tidak Diketahui';
    if (!summary.kecamatan[kec]) summary.kecamatan[kec] = 0;
    summary.kecamatan[kec]++;

    var op = idxOperasional !== -1 ? row[idxOperasional] : row[5];
    if (op && op.toString().toLowerCase().indexOf('ya') !== -1) {
      summary.operasionalYa++;
    }

    if (idxPengunjungDalam !== -1) summary.pengunjungDomestik += parseInt(row[idxPengunjungDalam]) || 0;
    if (idxPengunjungLuar !== -1) summary.pengunjungMancanegara += parseInt(row[idxPengunjungLuar]) || 0;
    if (idxPendapatan !== -1) summary.pendapatanTiket += parseRupiah(row[idxPendapatan]);
    if (idxKamarTersedia !== -1) summary.kamarTersedia += parseInt(row[idxKamarTersedia]) || 0;
    if (idxKamarTerisi !== -1) summary.kamarTerisi += parseInt(row[idxKamarTerisi]) || 0;

    processedData.push(obj);
  }

  var meOccupancy = (summary.kamarTersedia > 0) 
    ? (summary.kamarTerisi / summary.kamarTersedia * 100) 
    : 0;

  return {
    summary: {
      total: summary.total,
      kategori: summary.kategori,
      kecamatan: summary.kecamatan,
      operasional: {
        "Ya (Aktif)": summary.operasionalYa,
        "Tidak (Tutup)": summary.total - summary.operasionalYa
      },
      pengunjung: {
        domestik: summary.pengunjungDomestik,
        mancanegara: summary.pengunjungMancanegara
      },
      akomodasi: {
        kamarTersedia: summary.kamarTersedia,
        kamarTerisi: summary.kamarTerisi,
        tingkatOccupancy: meOccupancy
      },
      pendapatan: {
        tiket: summary.pendapatanTiket
      }
    },
    columnStats: columnStats,
    rawData: processedData,
    headers: headers
  };
}

function parseRupiah(val) {
  if (!val) return 0;
  var str = val.toString().replace(/[^0-9]/g, '');
  return parseInt(str) || 0;
}

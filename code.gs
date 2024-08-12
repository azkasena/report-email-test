function doGet() {
  kirimLaporanEmail();
  return ContentService.createTextOutput("Laporan telah dikirim melalui email.");
}

function kirimLaporanEmail() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  
  // Mendapatkan range data (dimulai dari B3 untuk menghindari header)
  var dataRange = sheet.getRange("B3:F");
  var data = dataRange.getValues();
  
  var htmlTable = createHtmlTable(data);
  
  var subject = "Laporan Penjualan " + new Date().toLocaleDateString();
  var body = "Berikut adalah laporan penjualan terbaru:<br><br> Silahkan balas email ini jika terdapat hal yang tidak sesuai" + htmlTable;
  
  MailApp.sendEmail({
    to: "azkablack@gmail.com",
    subject: subject,
    htmlBody: body
  });
}

function createHtmlTable(data) {
  var html = '<table border="1" style="border-collapse: collapse;">';
  
  // Menambahkan header secara manual
  html += '<tr style="background-color: #f2f2f2;">';
  html += '<th style="padding: 8px;">Tanggal</th>';
  html += '<th style="padding: 8px;">Nama Produk</th>';
  html += '<th style="padding: 8px;">Jumlah</th>';
  html += '<th style="padding: 8px;">Harga Satuan</th>';
  html += '<th style="padding: 8px;">Total</th>';
  html += '</tr>';
  
  // Menambahkan baris data
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] !== "") {  // Memeriksa apakah baris tidak kosong
      html += '<tr>';
      for (var j = 0; j < data[i].length; j++) {
        var cellValue = data[i][j];
        if (j === 0) {  // Kolom tanggal
          cellValue = Utilities.formatDate(new Date(cellValue), "GMT+7", "yyyy-MM-dd");
        } else if (j === 2 || j === 3 || j === 4) {  // Kolom angka
          cellValue = Number(cellValue).toLocaleString('id-ID');
        }
        html += '<td style="padding: 8px;">' + cellValue + '</td>';
      }
      html += '</tr>';
    }
  }
  
  html += '</table>';
  return html;
}

    // Import library XLSX melalui CDN
    if (typeof XLSX !== 'undefined') {
        console.log('Library XLSX telah berhasil diimpor.');
      } else {
        console.log('Gagal mengimpor library XLSX.');
      }
      
      // Fungsi perhitungan
      class Perhitungan {
        amount(tsi, share) {
          return tsi * share;
        }
      
        bppdan(amount, share) {
          return Math.min(amount * 0.025, 500000000 * share);
        }
      
        maipark(tsi, wilayah, share) {
          if (wilayah === "Jakarta" || wilayah === "Banten" || wilayah === "Jawa Barat") {
            return Math.min(tsi * 0.1, 100000000000 * share);
          } else if (wilayah === "No EQVET") {
            return 0;
          } else {
            return Math.min(tsi * 0.25, 100000000000 * share);
          }
        }
      
        ret(amount, bppdan, maipark) {
          return Math.min(amount - bppdan - maipark, 32000000000);
        }
      
        surplus(amount, bppdan, maipark, qs) {
          return Math.min(amount - bppdan - maipark - qs, 400000000000);
        }
      
        facul(amount, bppdan, maipark, qs, surplus) {
          return amount - bppdan - maipark - qs - surplus;
        }
      }
      
      // Event listener saat input file berubah
document.getElementById("file").addEventListener('change', function(event) {
    var input = event.target.files[0];
  
    var reader = new FileReader();
    reader.onload = function(e) {
      var contents = e.target.result;
      var workbook = XLSX.read(contents, { type: 'binary' });
  
      // Proses workbook dan lakukan perhitungan
      var sheetName = workbook.SheetNames[0]; // Anggap data ada di sheet pertama
      var worksheet = workbook.Sheets[sheetName];
      var data = XLSX.utils.sheet_to_json(worksheet);
  
      var p = new Perhitungan();
      var result = [];
  
      for (var i = 0; i < data.length; i++) {
        //pengambilan data 
        var row = data[i];
        var tsi = row.TSI;
        var wilayah = row.Wilayah;
        var share = row.Share;
        
        //perhitungan 
        var amount = p.amount(tsi, share);
        var bppdan = p.bppdan(amount, share);
        var maipark = p.maipark(tsi, wilayah, share);
        var ret = p.ret(amount, bppdan, maipark); // nilai retensi atau OR dari pengurangan amount, bppdan dan maipark
        var qs = ret * 0.2; // Hitung nilai QS (20% dari ret)
        var surplus = p.surplus(amount, bppdan, maipark, qs); // Hitung surplus dengan pengurangan QS bukan ret
        var facul = p.facul(amount, bppdan, maipark, qs, surplus); // Hitung facul dengan pengurangan QS bukan ret
        
        // Pembuatan Columns baru yang berisi data sebelumnya 
        row['Amount of Share'] = amount;
        row['BPPDAN'] = bppdan;
        row['MAIPARK'] = maipark;
        row['QS'] = qs; 
        row['Surplus'] = surplus;
        row['Facultative'] = facul;
  
        result.push(row);
      }
  
      var newWorkbook = XLSX.utils.book_new();
      var newWorksheet = XLSX.utils.json_to_sheet(result);
      XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);
      var xlsxFile = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });
  
      // Simpan file XLSX dengan nama "hasil.xlsx"
      var blob = new Blob([xlsxFile], { type: 'application/octet-stream' });
      var url = URL.createObjectURL(blob);
      var link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', 'hasil.xlsx');
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    };
        reader.readAsBinaryString(input);
        input = undefined;
      });
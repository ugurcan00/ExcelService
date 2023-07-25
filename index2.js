var Excel = require("exceljs");
const http = require("http");
const request = require("request");
const express = require("express");
require("cross-fetch/polyfill");
require("isomorphic-form-data");
const https = require("https");
const fs = require("fs")
// var options = {
//   key: fs.readFileSync('./keynew.pem'),
//   cert: fs.readFileSync('./certx.pem'),
// };

const istanbulIlceleri = {
  "01": "Adalar",
  "02": "Arnavutköy",
  "03": "Ataşehir",
  "04": "Avcılar",
  "05": "Bağcılar",
  "06": "Bahçelievler",
  "07": "Bakırköy",
  "08": "Başakşehir",
  "09": "Bayrampaşa",
  10: "Beşiktaş",
  11: "Beykoz",
  12: "Beylikdüzü",
  13: "Beyoğlu",
  14: "Büyükçekmece",
  15: "Çatalca",
  16: "Çekmeköy",
  17: "Esenler",
  18: "Esenyurt",
  19: "Eyüp",
  20: "Fatih",
  21: "Gaziosmanpaşa",
  22: "Güngören",
  23: "Kadıköy",
  24: "Kağıthane",
  25: "Kartal",
  26: "Küçükçekmece",
  27: "Maltepe",
  28: "Pendik",
  29: "Sancaktepe",
  30: "Sarıyer",
  31: "Silivri",
  32: "Sultanbeyli",
  33: "Sultangazi",
  34: "Şile",
  35: "Şişli",
  36: "Tuzla",
  37: "Ümraniye",
  38: "Üsküdar",
  39: "Zeytinburnu",
};
function timestampToDate(param) {
  var timestamp = new Date(Number(param));
  var ymdFormat = timestamp.getDate() + '/' + (timestamp.getMonth() + 1) + '/' + timestamp.getFullYear();
  console.log("ymdFormat", ymdFormat)
  return ymdFormat
}
String.prototype.turkishToUpper = function () {
  var string = this;
  var letters = { "i": "I", "ş": "S", "ğ": "Ğ", "ü": "U", "ö": "O", "ç": "C", "ı": "I" };
  string = string.replace(/(([iışğüçö]))/g, function (letter) { return letters[letter]; });
  return string.toUpperCase();
}
const app = express();
// https.createServer(options, app).listen(3000, function(){
//   console.log("Express server listening on port " +  3000);
// });
app.listen(3000,()=>{
  console.log('3000 portyu dinleniyor')
})
app.get("/", (req, res) => {
  console.log("ecrcimisil isteği geldi");
  res.send('deneme')
})

app.get("/ecrimisil", (req, res) => {
  console.log("ecrcimisil isteği geldi");
  var keys = Object.keys(req.query);
  var values = Object.values(req.query);
  for (let i = 0; i < values.length; i++) {
    console.log("index=", i, " ", keys[i], " = ", values[i])

  }


  var workbook = new Excel.Workbook();

  workbook.xlsx.readFile("data/last.xlsx").then(function (result) {

    var worksheet = workbook.getWorksheet("Tutanak Sayfası");
    if (values[0] != null) {
      var row = worksheet.getRow(11);
      row.getCell(6).value = values[0].turkishToUpper();
      //ilce kod
      for (const [key, value] of Object.entries(istanbulIlceleri)) {
        if (value.turkishToUpper() == values[0].turkishToUpper()) {
          var dosyano = worksheet.getRow(7);
          dosyano.getCell(14).value = "34 - " + key + " - ";
          console.log("ilçe kod yazıldı");
          if (values[18] != null) {
            console.log("mahalle kod yazıldı");
            var dosyano = worksheet.getRow(7);
            dosyano.getCell(14).value += values[18] + " - ";
          }
        } else {
          value.toUpperCase() == values[0].toUpperCase();
        }
      }
    } else {
      row.getCell(6).value = "-";
      return;
    }
    if (values[1] != null) {
      var mahalle = worksheet.getRow(13);
      mahalle.getCell(6).value = values[1].toUpperCase();
      // mahalle code
    } else {
      return;
    }
    //mulktipikmsmulk
    var ibbhisse = worksheet.getRow(17);
    var tapuAlani = worksheet.getRow(15);
    if (values[17] == "KMS") {
      ibbhisse.getCell(6).value = "-";
      tapuAlani.getCell(6).value = "-";

    } else if (values[17] == "MULK") {
      ibbhisse.getCell(6).value = "%" + values[12];
      tapuAlani.getCell(6).value = values[31] + " m²"
    }
    if (values[2] != null) {
      var pafta = worksheet.getRow(11);
      pafta.getCell(14).value = values[2].toUpperCase();
    } else {
      return;
    }
    if (values[3] != null) {
      var ada = worksheet.getRow(13);
      ada.getCell(14).value = values[3].toUpperCase();
    } else {
      return;
    }
    if (values[4] != null) {
      var parsel = worksheet.getRow(15);
      parsel.getCell(14).value = values[4].toUpperCase();
    } else {
      return;
    }
    var yonSonuc;
    if (values[5] != null) {
      if (values[5] == "K") {
        yonSonuc = "KUZEY";
      } else if (values[5] == "G") {
        yonSonuc = "GÜNEY";
      } else if (values[5] == "D") {
        yonSonuc = "DOĞU";
      } else if (values[5] == "B") {
        yonSonuc = "BATI";
      } else if (values[5] == "KD") {
        yonSonuc = "KUZEYDOĞU";
      } else if (values[5] == "GD") {
        yonSonuc = "GÜNEYDOĞU";
      } else if (values[5] == "KB") {
        yonSonuc = "KUZEYBATI";
      } else if (values[5] == "GB") {
        yonSonuc = "GÜNEYBATI";
      } else {
        yonSonuc = values[5];
      }
      console.log(yonSonuc)
      var yon = worksheet.getRow(17);
      if (values[17] == "KMS") {
        yon.getCell(14).value = yonSonuc + ", " + "KMS";
      } else if (values[17] == "MULK") {
        var yon = worksheet.getRow(17);
        yon.getCell(14).value = "-";
      }
    }
    else {
      return;
    }

    const kullanimlar = [];
    var alansayisi = values[20];
    console.log("alan sayisi= ", alansayisi);
    forText = [];
    var alanlartoplami = 0;
    if (values[20] != null) {
      var amacialani = worksheet.getRow(21);
      amacialani.getCell(4).value = "";
      for (var i = 1; i < alansayisi * 3 + 1; i++) {
        kullanimlar.push(values[25 + i].turkishToUpper());
        amacialani.getCell(4).value += values[25 + i].turkishToUpper();
        if (i % 3 != 0) {
          amacialani.getCell(4).value += ", ";
        }
        if (i % 3 == 0) {
          amacialani.getCell(4).value += " m²";
          if (i != alansayisi * 3) {
            amacialani.getCell(4).value += "\n";
          }
          var whenNumber = Number(values[25 + i]);
          alanlartoplami += whenNumber;
        }
      }
      //
      console.log("alanlar toplamı", alanlartoplami);
      var kullanimalan = worksheet.getRow(23);
      kullanimalan.getCell(14).value = alanlartoplami + " m²";
      //
      //
      var textmessage;
      const ucLuler = [];
      for (let i = 0; i < kullanimlar.length; i += 3) {
        ucLuler.push(kullanimlar.slice(i, i + 3));
      }
      const yeniDizi = ucLuler.map((grup) => {
        const birinci = `(${grup[0]})`.toLocaleLowerCase();
        const ikinci = grup[1].toLocaleLowerCase();
        const ucuncu = `${grup[2]} m²`;

        return [ucuncu, ikinci, birinci];
      });
      const sonucDizi = yeniDizi.flat();
      textmessage = sonucDizi;
      var sonucMetin = sonucDizi.join(", ");
    } else {
      return;
    }
    if (values[13] != null) {
      var baslangictarihi = worksheet.getRow(21);
      baslangictarihi.getCell(14).value = values[13];
    } else {
      return;
    }
    if (values[6] != null) {
      var isletmeadi = worksheet.getRow(27);
      isletmeadi.getCell(6).value = values[6].toUpperCase();
    } else {
      return;
    }
    if (values[7] != null) {
      var isletmesahibi = worksheet.getRow(28);
      isletmesahibi.getCell(6).value = values[7].toUpperCase();
    } else {
      return;
    }
    if (values[8] != null) {
      var iletisimno = worksheet.getRow(29);
      iletisimno.getCell(6).value = values[8];
    } else {
      return;
    }
    if (values[9] != null) {
      var tcno = worksheet.getRow(30);
      tcno.getCell(6).value = values[9];
    } else {
      return;
    }
    if (values[10] != null) {
      var vergino = worksheet.getRow(31);
      vergino.getCell(6).value = values[10];
    } else {
      return;
    }
    if (values[11] != null) {
      var tebligat = worksheet.getRow(32);
      tebligat.getCell(6).value = values[11].toUpperCase();
    } else {
      return;
    }
    if (values[15] != null) {
      var tespittarihi = worksheet.getRow(55);
      tespittarihi.getCell(4).value = timestampToDate(values[15]);
    } else {
      return;
    }
    if (values[16] && values[24] != null) {
      var tespitKullanici1 = worksheet.getRow(57);
      tespitKullanici1.getCell(2).value = values[16];
      var tespitKullanici2 = worksheet.getRow(57);
      tespitKullanici2.getCell(6).value = values[24];

    }
    //KULLANICI İPTAL EDİLDİ
    // var kullanan_adi_soyadi = worksheet.getRow(57);
    // kullanan_adi_soyadi.getCell(10).value = values[7].toUpperCase();

    // var aciklama = worksheet.getRow(36);
    // if (values[17] == "KMS") {
    //   aciklama.getCell(3).value =
    //     "2022/892833 İKN no’lu “İstanbul İl Sınırları İçerisinde Bulunan Mülkiyeti İBB'ye Ait Olan Tapuda Kayıtlı Taşınmazlar İle İBB'nin Sorumluluk Alanlarında Yer Alan Taşınmazların Güncel Durumlarının Tespit Edilmesi ve Kayıt Altına Alınması İşi” kapsamında İstanbul Büyükşehir Belediyesi Emlak Şube Müdürlüğü tarafından " +
    //     " " +
    //     values[0]
    //       .toLocaleLowerCase()
    //       .split(" ")
    //       .map((kelime) => kelime.slice(0, 1).toUpperCase() + kelime.slice(1))
    //       .join(" ") +
    //     " " +
    //     "İlçesi, " +
    //     " " +
    //     values[1]
    //       .toLocaleLowerCase()
    //       .split(" ")
    //       .map((kelime) => kelime.slice(0, 1).toUpperCase() + kelime.slice(1))
    //       .join(" ") +
    //     " " +
    //     "Mahallesi, " +
    //     " " +
    //     "anaarter" +
    //     values[22]+
    //     "Caddesi'ne cephe, " +
    //     values[3] +
    //     " " +
    //     "ada" +
    //     " " +
    //     values[4] +
    //     " " +
    //     "parselin" +
    //     " " +
    //     yonSonuc.toLocaleLowerCase() +
    //     " " +
    //     " " +
    //     "yönünde kalan " +
    //     values[17] +
    //     " alanının tespiti talep edilmiştir.\n"+
    //     //2.paragraf
    //     "İlgili alanda " +
    //     values[7]
    //       .toLocaleLowerCase()
    //       .split(" ")
    //       .map((kelime) => kelime.slice(0, 1).toUpperCase() + kelime.slice(1))
    //       .join(" ") +
    //     " " +
    //     "tarafından işletilen" +
    //     " " +
    //     values[6]
    //       .toLocaleLowerCase()
    //       .split(" ")
    //       .map((kelime) => kelime.slice(0, 1).toUpperCase() + kelime.slice(1))
    //       .join(" ") +
    //     " " +
    //     "adlı "+ 
    //     values[23]+
    //     "işgal tespit edilmiştir." +
    //     sonucMetin +
    //     " bulunduğu görülmüştür." +
    //     "\n Ek: \n 1-Uydu   Görüntüsü/Hava   Fotoğrafı   (2017)   (1 Adet)\n2-Halihazır/Kroki   (1/1000-2017)   (1 Adet)\n3-Fotoğraf (4 Adet)\n4-Fotoğraf ( x adet)\n5-Fotoğraf (Tapu/Tapu Tahsis Belgesi)\n6-Fotoğraf (Vergi Levhası)\n7-Fotoğraf (Kaşe)\n8-Fotoğraf (İş Yeri Açma ve Çalışma Ruhsatı)\n9-Fotoğraf (Yapı Kayıt Belgesi)\n10-Fotoğraf (Kira Sözleşmesi)\n11-Diğer";
    // }
    // else if(values[17] == "MULK"){
    //   aciklama.getCell(3).value =
    //   "2022/892833 İKN no’lu “İstanbul İl Sınırları İçerisinde Bulunan Mülkiyeti İBB'ye Ait Olan Tapuda Kayıtlı Taşınmazlar İle İBB'nin Sorumluluk Alanlarında Yer Alan Taşınmazların Güncel Durumlarının Tespit Edilmesi ve Kayıt Altına Alınması İşi” kapsamında İstanbul Büyükşehir Belediyesi Emlak Şube Müdürlüğü tarafından " +
    //   " " +
    //   values[0]
    //     .toLocaleLowerCase()
    //     .split(" ")
    //     .map((kelime) => kelime.slice(0, 1).toUpperCase() + kelime.slice(1))
    //     .join(" ") +
    //   " " +
    //   "İlçesi, " +
    //   " " +
    //   values[1]
    //     .toLocaleLowerCase()
    //     .split(" ")
    //     .map((kelime) => kelime.slice(0, 1).toUpperCase() + kelime.slice(1))
    //     .join(" ") +
    //   " " +
    //   "Mahallesi, " +
    //   " " +
    //   "anaarter" +
    //   values[22]+
    //   "Caddesi'ne cephe, " +
    //   values[3] +
    //   " " +
    //   "ada" +
    //   " " +
    //   values[4] +
    //   " " +
    //   "parseldeki mülk alanının tespiti talep edilmiştir.\n"+
    //   //2.paragraf
    //   "İlgili alanda " +
    //   values[7]
    //     .toLocaleLowerCase()
    //     .split(" ")
    //     .map((kelime) => kelime.slice(0, 1).toUpperCase() + kelime.slice(1))
    //     .join(" ") +
    //   " " +
    //   "tarafından işletilen" +
    //   " " +
    //   values[6]
    //     .toLocaleLowerCase()
    //     .split(" ")
    //     .map((kelime) => kelime.slice(0, 1).toUpperCase() + kelime.slice(1))
    //     .join(" ") +
    //   " " +
    //   "adlı "+ 
    //   values[23]+
    //   "işgal tespit edilmiştir." +
    //   sonucMetin +
    //   " bulunduğu görülmüştür." +
    //   "\n Ek: \n 1-Uydu   Görüntüsü/Hava   Fotoğrafı   (2017)   (1 Adet)\n2-Halihazır/Kroki   (1/1000-2017)   (1 Adet)\n3-Fotoğraf (4 Adet)\n4-Fotoğraf ( x adet)\n5-Fotoğraf (Tapu/Tapu Tahsis Belgesi)\n6-Fotoğraf (Vergi Levhası)\n7-Fotoğraf (Kaşe)\n8-Fotoğraf (İş Yeri Açma ve Çalışma Ruhsatı)\n9-Fotoğraf (Yapı Kayıt Belgesi)\n10-Fotoğraf (Kira Sözleşmesi)\n11-Diğer";
    // }


    //sheets 
    for (let i = 1; i < 12; i++) {
      var worksheets = workbook.getWorksheet(`Ek ${i}`);
      var ilce_1 = worksheets.getRow(3);
      ilce_1.getCell(3).value = values[0].toUpperCase();
      var mahalle_1 = worksheets.getRow(3);
      mahalle_1.getCell(5).value = values[1].toUpperCase();
      var ada_1 = worksheets.getRow(3);
      ada_1.getCell(7).value = values[3].toUpperCase();
      var parsel_1 = worksheets.getRow(3);
      parsel_1.getCell(9).value = values[4].toUpperCase();
      var yon_1 = worksheets.getRow(3);
      yon_1.getCell(11).value = yonSonuc;
    }
    // var worksheet6 = workbook.addWorksheet('attachs');
    // //  BASE64 RESİM EKLEME VALUES 17
    // console.log("attach sayi", values[19]);
    // if (values[19] != null) {
    //   var attachsayisi = values[19];
    //   for (var i = 0; i < attachsayisi; i++) {
    //     var url = values[23 + values[20] * 3 + 1 + i];
    //     url.toString();
    //     request.get(url, { encoding: null }, (err, res, body) => {
    //       const base64Image = Buffer.from(body).toString("base64");
    //       // Görüntüyü Excel hücresine ekleyin
    //       const imageId = workbook.addImage({
    //         buffer: Buffer.from(base64Image, "base64"),
    //         extension: "jpeg",
    //       });
    //       worksheet6.addImage(imageId, `B5:K20`);
    //     });
    //   }
    // }

    setTimeout(function () {
      row.commit();
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
      res.setHeader(
        "Content-Disposition",
        "attachment; filename=" + "berkay.xlsx"
      );
      return workbook.xlsx.write(res).then(function () {
        // console.log(res);
        res.status(200);
        console.log("excel gönderildi")
        // workbook.removeWorksheet('attachs');
        // const newsheet = workbook.addWorksheet('attachs');
      });
    }, 1500);
  });
});
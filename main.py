import base64
import datetime

import xlrd as xlrd

from flask import Flask, render_template, request, redirect, session
import pymongo
from bson.objectid import ObjectId

app = Flask(__name__)
app.secret_key = 'bizim cok zor gizli sozcugumuz'

myclient = pymongo.MongoClient("mongodb://localhost:27017/")
mydb = myclient["obsDB"]
kullanicilar_tablosu = mydb["kullanicilar"]
ogrenciler_tablosu = mydb["ogrenciler"]
dersler_tablosu = mydb["dersler"]



@app.route('/')
def baslangic():
    kayit = None
    if 'kullanici' in session:
        kayit = session["kullanici"]

    return render_template("anasayfa.html")


@app.route('/giris', methods=['GET', 'POST'])
def giris():
    if request.method == 'POST':
        kullanici = request.form['kullanici']
        sifre = request.form['sifre']
        kayit = kullanicilar_tablosu.find_one({"_id": kullanici})
        if kayit:
            if sifre == kayit["sifre"]:
                del kayit['sifre']
                session["kullanici"] = kayit
                return redirect("/", code=302)
            else:
                return "Şifre yanlış"
        else:
            return "Kullanıcı bulunamadı"
    else:
        return render_template("giris.html")

@app.route('/dosyayukle')
def dosya_yukle():
    loc = ("data/MYOKazananlar.xls")

    # To open Workbook
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)

    # For row 0 and column 0
    print(sheet.cell_value(0, 0))
    print(sheet.nrows)
    
    for i in range(1, sheet.nrows):
        tckn = str(int(sheet.cell_value(i, 0)))
        adSoyad = sheet.cell_value(i, 1)
        telefon = str(sheet.cell_value(i, 2))
        bolum = sheet.cell_value(i, 3)
        kullanici_adi = sheet.cell_value(i, 4)
        kayit = {"tckn": tckn, "adSoyad": adSoyad, "kayit_tarihi": datetime.datetime.now(), "telefon": telefon, "bolum": bolum, "kullanici_adi":kullanici_adi}
        ogrenciler_tablosu.insert_one(kayit)
        kullanicilar_tablosu.insert_one({"_id":kullanici_adi, "adSoyad":adSoyad, "sifre":tckn})
    return redirect("/", code=302)


@app.route('/cikis')
def cikis():
    session.clear()
    return redirect("/", code=302)

if __name__ == "__main__":
    app.run(debug=True, port=5001)
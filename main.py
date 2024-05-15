import json
import pandas as pd
import smtplib
import requests
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def create_excel_report(data, file_name):
    # Veri içindeki fonlar ve getirilerin isimleri
    getiri_isimleri = ["GETIRI3A", "GETIRI6A", "GETIRI1Y", "GETIRIYB", "GETIRI3Y", "GETIRI5Y"]
    diger_bilgiler = ["FONKODU", "FONUNVAN", "FONTURACIKLAMA"]

    # Excel dosyası oluştur
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')

    # Her bir getiri türü için ayrı bir tablo oluştur
    for getiri_adi in getiri_isimleri:
        # Tüm fonları dolaşarak getirileri al
        getiriler = []
        for fon in data["data"]:
            getiri_degeri = fon.get(getiri_adi)
            if getiri_degeri is not None:
                fon_bilgileri = [fon.get(bilgi, None) for bilgi in diger_bilgiler]
                getiriler.append([getiri_degeri] + fon_bilgileri)

        # None olmayan getirileri filtrele
        if getiriler:
            df = pd.DataFrame(getiriler, columns=[getiri_adi] + diger_bilgiler)
            # En düşük ve en yüksek değerleri hesapla
            en_dusuk = df[getiri_adi].min()
            en_yuksek = df[getiri_adi].max()
            # En düşük ve en yüksek değerleri tabloya ekle
            df.loc[len(df)] = [en_dusuk, f"En düşük {getiri_adi}", "", ""]
            df.loc[len(df)] = [en_yuksek, f"En yüksek {getiri_adi}", "", ""]
            # Excel tablosuna yaz
            df.to_excel(writer, sheet_name=getiri_adi, index=False)

    # Excel dosyasını kaydet
    writer._save()


def send_email_with_attachment(receiver_email, attachment_file):
    sender_email = "CHANGEME!"
    sender_password = "CHANGEME!"

    # E-posta gövdesi oluştur
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = "Excel Raporu"

    # E-posta metni
    body = "Merhaba,\n\nEk olarak Excel raporu gönderiyorum.\n\nİyi günler."
    message.attach(MIMEText(body, "plain"))

    # Dosya ekleme
    with open(attachment_file, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    encoders.encode_base64(part)

    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {attachment_file}",
    )

    message.attach(part)

    # E-posta gönderme
    with smtplib.SMTP("smtp-mail.outlook.com", 587) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        text = message.as_string()
        server.sendmail(sender_email, receiver_email, text)
        print("E-posta gönderildi.")

def main():
    # JSON verisini aç
    cookies = {
        'ASP.NET_SessionId': 'xqgknh3zbkekwubha0xdvzkk',
        'TS01ec1a88': '019fc7fc4d93bbc5dc60e7dcb7cb081aa18ba3fc7b53d84769e2f5dcf39842a983ab74839c604809f0cd280131ff54034bd7e4f89cf046cb7d61ba92092e802b33906cafcd',
        'TSb7d61442027': '08ce165641ab20002be8c729e4207bc710294220ab51fdcbe264235a4e1d91c94a7ae4cb104292e7086f596b90113000d9a491d87c460672a8947ea7d1e8daa4b812ca054b96debdc155181b91808488ed7218365b125b1528821250cbda81f1',
    }

    headers = {
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Accept-Language': 'en-US,en;q=0.6',
        'Connection': 'keep-alive',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        # 'Cookie': 'ASP.NET_SessionId=xqgknh3zbkekwubha0xdvzkk; TS01ec1a88=019fc7fc4d93bbc5dc60e7dcb7cb081aa18ba3fc7b53d84769e2f5dcf39842a983ab74839c604809f0cd280131ff54034bd7e4f89cf046cb7d61ba92092e802b33906cafcd; TSb7d61442027=08ce165641ab20002be8c729e4207bc710294220ab51fdcbe264235a4e1d91c94a7ae4cb104292e7086f596b90113000d9a491d87c460672a8947ea7d1e8daa4b812ca054b96debdc155181b91808488ed7218365b125b1528821250cbda81f1',
        'Origin': 'https://www.tefas.gov.tr',
        'Referer': 'https://www.tefas.gov.tr/FonKarsilastirma.aspx',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-GPC': '1',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.54 Safari/537.36',
        'X-Requested-With': 'XMLHttpRequest',
        'sec-ch-ua': '"Google Chrome";v="101", "Chromium";v="101", "Not=A?Brand";v="24"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
    }

    data = {
        'calismatipi': '2',
        'fontip': 'YAT',
        'sfontur': '',
        'kurucukod': '',
        'fongrup': '',
        'bastarih': 'Başlangıç',
        'bittarih': 'Bitiş',
        'fonturkod': '',
        'fonunvantip': '',
        'strperiod': '1,1,1,1,1,1,1',
        'islemdurum': '1',
    }

    response = requests.post('https://www.tefas.gov.tr/api/DB/BindComparisonFundReturns', cookies=cookies, headers=headers, data=data)

    # verileri cek ve veriler.json dosyasina kaydet.
    fon_veri = open("veriler.json","w")
    fon_veri.write(response.text)
    fon_veri.close()
    with open('veriler.json', 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # Excel raporu oluştur
    excel_file = "rapor.xlsx"
    create_excel_report(data, excel_file)

    # Alıcı e-posta adresini al
    receiver_email = input("Alıcı e-posta adresini girin: ")

    # E-posta gönder
    send_email_with_attachment(receiver_email, excel_file)

if __name__ == "__main__":
    main()

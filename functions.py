import os
import pandas as pd, re
from datetime import datetime


def getcurrentime():
    mevcut_zaman = datetime.now()

    # Saat, dakika ve saniye değerlerini alarak string olarak formatlama
    zaman_str = "{:%H%M%S}".format(mevcut_zaman)
    return zaman_str

def sort_folders(folders):
    def natural_keys(text):
        def atoi(text):
            return int(text) if text.isdigit() else text

        return [atoi(c) for c in re.split(r'(\d+)', text)]

    return sorted(folders, key=natural_keys)

def validate_layer(string):
    pattern = r"^Layer\s\d+$"
    return re.match(pattern, string) is not None

def list_folders(directory: str):
    #directory = directory.replace("file://","")
    folders = [folder for folder in os.listdir(directory) if os.path.isdir(os.path.join(directory, folder))]
    sorted_folders = sort_folders(folders)  # Klasörleri sırala

    return sorted_folders

def list_files(directory: str):
    #directory = directory.replace("file://","")
     #if f.endswith(".xlsx")
    folders = [f for f in os.listdir(directory)]
    return folders

def get_layers(file_path):
    try:
        sheets = []
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names
        for name in sheet_names:
             if validate_layer(name):
                 sheets.append(name)
        return sheets
    except Exception as e:
        print(f"Excel Layer Tespit Edilirken Hata oluştu: {e}")
        return []
    

def check_files(folder, file):
    dosya_yolu = os.path.join(folder, file)
    return os.path.isfile(dosya_yolu)


def excel_sheet_check(excel_yolu, sayfa_adi):
    try:
        # Excel dosyasını oku
        excel = pd.ExcelFile(excel_yolu)

        # Dosyadaki mevcut sayfaları al
        mevcut_sayfalar = excel.sheet_names

        # Belirtilen sayfanın mevcut sayfalar arasında olup olmadığını kontrol et
        if sayfa_adi in mevcut_sayfalar:
            return True
        else:
            print(f"'{sayfa_adi}' adında bir sayfa bulunamadı.")
            return False
    except FileNotFoundError:
        print(f"Hata: {excel_yolu}{sayfa_adi} dosyası bulunamadı veya dosya yolu geçersiz.")
        return False
    except Exception as e:
        print(f"Hata: {e}")
        return False


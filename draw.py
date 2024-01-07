import pandas as pd, os
import matplotlib.pyplot as plt
from PyQt6.QtWidgets import QApplication, QMainWindow, QProgressDialog
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon


def check_sonuclar_folder():
    # Mevcut çalışma dizinini al
    mevcut_dizin = os.getcwd()

    # Belirtilen klasör adını içeren dosya yolunu oluştur
    klasor_yolu = os.path.join(mevcut_dizin, "sonuclar")

    # Klasörün var olup olmadığını kontrol et
    if not os.path.exists(klasor_yolu):
        # Klasör yoksa oluştur
        try:
            os.makedirs(klasor_yolu)
            print(f"Sonuçlar klasörü oluşturuldu.")
        except OSError as e:
            print(f"Hata: Sonuçlar klasörü oluşturulamadı - {e}")
    else:
        print(f"Sonuçlar klasörü zaten var.")

    
def getperiod():
    try:
        excel_data = pd.read_excel("period.xlsx", sheet_name="Sayfa1")
        selected_column = excel_data["Period (sec)"].dropna().tolist()
        return selected_column
    except FileNotFoundError:
        return f"Period dosyası bulunamadı."
    except Exception as e:
        return f"Hata oluştu: {e}"
    
def get_sae(sae: str):
    try:
        excel_data = pd.read_excel("period.xlsx", sheet_name="Sayfa1")
        selected_column = excel_data[sae].dropna().tolist()
        return selected_column
    except FileNotFoundError:
        return f"{sae.capitalize()} dosyası bulunamadı."
    except Exception as e:
        return f"Hata oluştu: {e}"


def multi_analysis_with_progress_bar(folders, deprem, layer, column):
    total_folders = len(folders)
    progress_dialog = QProgressDialog("İşlem devam ediyor...", "İptal", 0, total_folders)
    progress_dialog.setWindowIcon(QIcon("icon.ico"))
    progress_dialog.setWindowTitle("Analiz İlerlemesi")
    progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
    progress_dialog.setMinimumDuration(0)
    progress_dialog.setAutoReset(False)
    progress_dialog.setAutoClose(True)

    all_y_values = []
    input_motion = []
    
    for index, folder in enumerate(folders):
        # İşlem adımını güncelle (progress bar)
        progress_dialog.setValue(index)
        QApplication.processEvents()  # Arayüzün donmamasını sağlamak için güncellemeleri işler
        
        file_path = os.path.join(folder, deprem)

        if os.path.exists(folder):
            try:
                excel_data = pd.read_excel(file_path, sheet_name=layer)
                y_values = excel_data[column].dropna().tolist()
                all_y_values.append(y_values)
                input_data = pd.read_excel(file_path, sheet_name="Input Motion", header=1)
                input = input_data["PSA (g)"].dropna().tolist()
                input_motion.append(input)
                
            except Exception as e:
                print(f"MULTİ ANALYSİS {deprem} dosyasını okuma sırasında hata oluştu: {e}")
                
        else:
            print(f"{deprem} dosyası bulunamadı: {folder}")
    
    # İşlem tamamlandığında progress bar'ı kapat
    progress_dialog.close()

    return all_y_values, input_motion[0]

    

def draw_test(y_values,graphname: str, title: str,input_motion,labels, motion=False):
    label = labels
    x_values = getperiod()
    plt.figure(figsize=(15, 9))
    plt.xscale('log')  # Y eksenini logaritmik ölçekte gösterme
    plt.tight_layout()
    #plt.ylim(bottom=0,top=1.75)
    for idx, y_data in enumerate(y_values,start=0):
        plt.plot(x_values, y_data,label=label[idx])
    if motion:        
        plt.plot(x_values,input_motion,label="Input Motion")    
    plt.xlabel("Periyot (X)")
    plt.ylabel("PSA (g) (y)")
    plt.title(str(title))
    plt.tight_layout()
    plt.legend()
    plt.savefig(f"sonuclar/{graphname}")
    plt.close()



def multi_deprem_with_progress_bar(folder_path, file_names, sheet_name, column_name):
    total_files = len(file_names)
    progress_dialog = QProgressDialog("İşlem devam ediyor...", "İptal", 0, total_files)
    progress_dialog.setWindowTitle("Deprem İşlemi İlerlemesi")
    progress_dialog.setWindowIcon(QIcon("icon.ico"))
    progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
    progress_dialog.setMinimumDuration(0)
    progress_dialog.setAutoReset(False)
    progress_dialog.setAutoClose(True)

    all_values = []

    for index, file_name in enumerate(file_names):
        # İşlem adımını güncelle (progress bar)
        progress_dialog.setValue(index)
        QApplication.processEvents()  # Arayüzün donmamasını sağlamak için güncellemeleri işler
        
        file_path = os.path.join(folder_path, file_name)

        try:
            excel_data = pd.read_excel(file_path, sheet_name=sheet_name)
            if column_name in excel_data.columns:
                values = excel_data[column_name].dropna().tolist()
                all_values.append(values)
                print(f"{file_name} {column_name} başarılı")
        except Exception as e:
            print(f"Hata: {file_name} dosyasında okuma sırasında bir hata oluştu multi_deprem- {e}")
    
    progress_dialog.close()
    return all_values

def multi_layer_with_progress_bar(folder, sheet_name, column_name, motion=False):
    total_sheets = len(sheet_name)
    progress_dialog = QProgressDialog("İşlem devam ediyor...", "İptal", 0, total_sheets)
    progress_dialog.setWindowIcon(QIcon("icon.ico"))

    progress_dialog.setWindowTitle("Layer İşlemi İlerlemesi")
    progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
    progress_dialog.setMinimumDuration(0)
    progress_dialog.setAutoReset(False)
    progress_dialog.setAutoClose(True)

    all_values = []

    file_path = os.path.join(folder)

    for index, sheet in enumerate(sheet_name):
        progress_dialog.setValue(index)
        QApplication.processEvents()

        try:
            excel_data = pd.read_excel(file_path, sheet_name=sheet)
            if column_name in excel_data.columns:
                values = excel_data[column_name].dropna().tolist()
                all_values.append(values)
                print(f"{file_path} {sheet} başarılı")
        except Exception as e:
            print(f"Hata: {folder} dosyasında okuma sırasında bir hata oluştu multi_layer- {e}")

    if motion:
        try:
            excel_data = pd.read_excel(file_path, sheet_name="Input Motion", header=1)
            if "PSA (g)" in excel_data.columns:
                values = excel_data["PSA (g)"].dropna().tolist()
                all_values.append(values)
                print("başarılı input motion")
        except Exception as e:
            print(f"Hata: {folder} dosyasında okuma sırasında bir hata oluştu - {e}")

    progress_dialog.close()
    return all_values

def draw_deprem(y_values,graphname: str, title: str,labels):
    label = labels
    x_values = getperiod()
    plt.figure(figsize=(15, 9))
    plt.xscale('log')  # Y eksenini logaritmik ölçekte gösterme
    plt.tight_layout()
    #plt.ylim(bottom=0,top=1.75)
    for idx, y_data in enumerate(y_values,start=0):
        plt.plot(x_values, y_data,label=labels[idx])
        
    plt.xlabel("Periyot (X)")
    plt.ylabel("PSA (g) (y)")
    plt.title(str(title))
    plt.tight_layout()
    plt.legend()
    plt.savefig(f"sonuclar/{graphname}")
    plt.close()   
    
def draw_layer(y_values,graphname: str, title: str,labels, motion = False):
    label = labels
    x_values = getperiod()
    plt.figure(figsize=(15, 9))
    plt.xscale('log')  # Y eksenini logaritmik ölçekte gösterme
    plt.tight_layout()
    #plt.ylim(bottom=0,top=1.75)
    for idx, y_data in enumerate(y_values,start=0):
        plt.plot(x_values, y_data,label=label[idx])
        
    plt.xlabel("Periyot (X)")
    plt.ylabel("PSA (g) (y)")
    plt.title(str(title))
    plt.tight_layout()
    plt.legend()
    plt.savefig(f"sonuclar/{graphname}")
    plt.close()   


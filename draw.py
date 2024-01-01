import pandas as pd, os
import matplotlib.pyplot as plt

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


def multi_analysis(folders, deprem, layer, column):
    all_y_values = []
    input_motion = []
    for folder in folders:
        file_path = os.path.join(folder, deprem)

        if os.path.exists(folder):
                try:
                    excel_data = pd.read_excel(file_path, sheet_name=layer)
                    y_values = excel_data[column].dropna().tolist()
                    all_y_values.append(y_values)
                    input_data = pd.read_excel(file_path, sheet_name="Input Motion",header=1)
                    input = input_data["PSA (g)"].dropna().tolist()
                    input_motion.append(input)
                
                except Exception as e:
                    print(f"MULTİ ANALYSİS {deprem} dosyasını okuma sırasında hata oluştu: {e}")
                    
        else:
            print(f"{deprem} dosyası bulunamadı: {folder}")
       
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



def multi_deprem(folder_path, file_names, sheet_name, column_name):
    all_values = []

    # Verilen klasördeki her dosya için işlem yap
    for file_name in file_names:
        file_path = os.path.join(folder_path, file_name)

        # Excel dosyasını oku
        try:
            excel_data = pd.read_excel(file_path, sheet_name=sheet_name)
            if column_name in excel_data.columns:
                values = excel_data[column_name].dropna().tolist()
                all_values.append(values)
                print (f"{file_name} {column_name} başarılı")
        except Exception as e:
            print(f"Hata: {file_name} dosyasında okuma sırasında bir hata oluştu multi_deprem- {e}")
    
    return all_values

def multi_layer(folder, sheet_name, column_name, motion = False):
    all_values = []

    # Verilen klasördeki her dosya için işlem yap
    
    file_path = os.path.join(folder)

    # Excel dosyasını oku
    for sheet in sheet_name:
        try:
            excel_data = pd.read_excel(file_path, sheet_name=sheet)
            if column_name in excel_data.columns:
                values = excel_data[column_name].dropna().tolist()
                all_values.append(values)
                print (f"{file_path} {sheet} başarılı")
        except Exception as e:
            print(f"Hata: {folder} dosyasında okuma sırasında bir hata oluştu multi_layer- {e}")
        
    if motion == True:
        try:
            excel_data = pd.read_excel(file_path,sheet_name="Input Motion",header=1)
            if "PSA (g)" in excel_data.columns:
                values = excel_data["PSA (g)"].dropna().tolist()
                all_values.append(values)
                print ("başarılı input motion")    

        except Exception as e:
            print(f"Hata: {folder} dosyasında okuma sırasında bir hata oluştu - {e}")
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


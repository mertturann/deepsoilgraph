import sys,os
from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QListWidget, QRadioButton, QFileDialog, QMessageBox, QCheckBox, QProgressBar
from PyQt6.QtGui import QIcon
from functions import list_folders, list_files, get_layers, getcurrentime, check_files, excel_sheet_check
from PyQt6.uic import loadUi
from mainwindow import Ui_MainWindow
from draw import multi_analysis_with_progress_bar, draw_test, multi_deprem_with_progress_bar, draw_deprem, multi_layer_with_progress_bar, get_sae, check_sonuclar_folder

class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        loadUi("main.ui", self)

        #self.ui = Ui_MainWindow()
        #self.ui.setupUi(self)
        #icon_path = os.path.join(sys._MEIPASS, 'icon.ico')
        self.setWindowIcon(QIcon("icon.ico"))
        self.setWindowTitle("DeepSoil Analysis")
        self.initUiElements()
        self.interactions()

    def initUiElements(self):
        self.analiz = self.findChild(QListWidget, "listWidget")
        self.deprem = self.findChild(QListWidget, "listWidget_2")
        self.layer = self.findChild(QListWidget, "listWidget_3")
        self.browse = self.findChild(QPushButton, "pushButton")
        self.draw = self.findChild(QPushButton, "pushButton_2")
        self.temizle = self.findChild(QPushButton, "pushButton_3")
        self.radio_analiz = self.findChild(QRadioButton, "radioButton_2")
        self.radio_deprem = self.findChild(QRadioButton, "radioButton")
        self.radio_layer = self.findChild(QRadioButton, "radioButton_3")
        self.sae_gez = self.findChild(QCheckBox,"checkBox")
        self.sae_ilica = self.findChild(QCheckBox,"checkBox_2")
        self.sae_yarimca = self.findChild(QCheckBox,"checkBox_3")
        self.checkbox_input_motion = self.findChild(QCheckBox,"checkBox_4")
        self.progress = self.findChild(QProgressBar,"progressBar")

    def interactions(self):
        self.trigger_radio()
        self.browse.clicked.connect(self.browsedir)
        self.analiz.currentItemChanged.connect(self.init_analiz)
        self.deprem.currentItemChanged.connect(self.init_layer)
        self.radio_analiz.clicked.connect(self.trigger_radio)
        self.radio_deprem.clicked.connect(self.trigger_radio)
        self.radio_layer.clicked.connect(self.trigger_radio)
        self.draw.clicked.connect(self.run)
        self.temizle.clicked.connect(self.clear_selections)

        
        
    def if_checkbox(self):
        values = []
        labels = []
        if self.sae_gez.isChecked():
            value = get_sae("sae-gez")
            values.append(value)
            labels.append("SAE-Gez")
        if self.sae_ilica.isChecked():
            value = get_sae("sae-ilica")
            values.append(value)
            labels.append("SAE-Ilıca")
        if self.sae_yarimca.isChecked():
            value = get_sae("sae-yarimca")
            values.append(value)        
            labels.append("SAE-Yarımca")
        return values, labels    
    
    
    def clear_selections(self):
        self.analiz.clearSelection()
        self.deprem.clearSelection()
        self.layer.clearSelection()
        
    def trigger_radio(self):
        if self.radio_analiz.isChecked():
            QMessageBox.information(self, "Information", "You are in analysis mode.")
            self.checkbox_input_motion.setDisabled(False)
            self.deprem.clear()
            self.layer.clear()
            self.analiz.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
            self.deprem.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
            self.layer.setSelectionMode(QListWidget.SelectionMode.SingleSelection)  
        elif self.radio_deprem.isChecked():
            QMessageBox.information(self, "Information", "You are in Input Mode.")
            self.checkbox_input_motion.setDisabled(True)
            self.deprem.clear()
            self.layer.clear()
            self.analiz.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
            self.deprem.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
            self.layer.setSelectionMode(QListWidget.SelectionMode.SingleSelection)  
        elif self.radio_layer.isChecked():
            QMessageBox.information(self, "Information", "You are in Layer Mode.")
            self.checkbox_input_motion.setDisabled(False)
            self.deprem.clear()
            self.layer.clear()
            self.analiz.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
            self.deprem.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
            self.layer.setSelectionMode(QListWidget.SelectionMode.MultiSelection)  
        else:
            QMessageBox.critical(self,"Error","Error in Mode Selection") 

    def browsedir(self):
        check_sonuclar_folder()
        self.analiz.clear()
        HOME_PATH = os.path.expanduser('~\\Documents')
        self.browse = QFileDialog.getExistingDirectory(self, directory=HOME_PATH)
        if self.browse:
            folders = list_folders(self.browse)
            self.analiz.addItems(folders)
        else:
            return    

    def run(self):
        if self.radio_analiz.isChecked():
            self.draw_multi_analysis()
        elif self.radio_deprem.isChecked():
            self.draw_multi_deprem()
        elif self.radio_layer.isChecked():
            self.draw_multi_layer()
        else:
            QMessageBox(self,)
    def draw_multi_analysis(self):
        dirs = []
        labels = []
        layer = self.get_selected_item_text(self.layer)
        deprem = self.get_selected_item_text(self.deprem)
        selected_items = self.analiz.selectedItems()
        basedir = self.browse
        for item in selected_items:
            if check_files(folder=f"{basedir}/{item.text()}", file=deprem):
                if excel_sheet_check(f"{basedir}/{item.text()}/{deprem}",layer):
                    dirs.append(f"{basedir}/{item.text()}")     
                else:
                    QMessageBox.critical(self, "Error", f"Selected Layer Not Found: Dir: {basedir}/{item.text()} Input: {deprem}.\n Analysis Will Not Start.")
                    return

            else:
                 QMessageBox.critical(self, "Hata", f"Selected Input Not Found: Dir: {basedir}/{item.text()} Input: {deprem}.\n Analysis Will Not Start.")
                 return
    
                       
        for label in selected_items:
            labels.append(label.text())
        sae_values, sae_labels = self.if_checkbox()
        
        y, input_motion = multi_analysis_with_progress_bar(folders=dirs, deprem=deprem, layer=layer, column="PSA (g)")
        y.extend(sae_values)
        labels.extend(sae_labels)
        if self.checkbox_input_motion.isChecked():
            motion = True
        else:
            motion = False    
        draw_test(y_values=y, graphname=getcurrentime(), title=deprem, input_motion=input_motion, labels=labels, motion=motion) 

    def draw_multi_deprem(self):
        path = self.browse
        layer = self.get_selected_item_text(self.layer)
        dizin = self.get_selected_item_text(self.analiz)
        folder = f"{path}/{dizin}"
        depremler = self.deprem.selectedItems()
        secili_depremler = [item.text() for item in depremler]
        labels = [((item.text()).replace("Results_profile_0_motion_A-", "")).replace(".xlsx", "") for item in depremler]
        sea_values, sea_labels = self.if_checkbox()                         
        sonuclar = multi_deprem_with_progress_bar(folder_path=folder, file_names=secili_depremler, sheet_name=layer, column_name="PSA (g)")
        sonuclar.extend(sea_values)
        labels.extend(sea_labels)
        draw_deprem(y_values=sonuclar, graphname=getcurrentime(), title=f"Analiz: {dizin}, Layer: {layer}", labels=labels)
        
    def draw_multi_layer(self):
        path = self.browse
        deprem = self.get_selected_item_text(self.deprem)
        dizin = self.get_selected_item_text(self.analiz)
        folder = f"{path}/{dizin}/{deprem}"
        layerlar = self.layer.selectedItems()
        secili_layerlar = [item.text() for item in layerlar]
        labels = [item.text() for item in layerlar]
        if self.checkbox_input_motion.isChecked():                            
            sonuclar = multi_layer_with_progress_bar(folder=folder, sheet_name=secili_layerlar, column_name="PSA (g)",motion=True)
            labels.append("İnput Motion")
        else:
            sonuclar = multi_layer_with_progress_bar(folder=folder, sheet_name=secili_layerlar, column_name="PSA (g)")                
        sea_values, sea_labels = self.if_checkbox()
        sonuclar.extend(sea_values)
        labels.extend(sea_labels)
        draw_deprem(y_values=sonuclar, graphname=getcurrentime(), title=f"Analysis: {dizin}, Input: {deprem}", labels=labels)
       

    def init_analiz(self):
        self.deprem.clear()
        text = self.analiz.currentItem()
        filespath = f"{self.browse}/{text.text()}"
        files = list_files(filespath)
        self.deprem.addItems(files)

    def init_layer(self):
        self.layer.clear()
        analiz = self.analiz.currentItem()
        deprem = self.deprem.currentItem()
        if analiz is not None and deprem is not None:
            path = f"{self.browse}/{analiz.text()}/{deprem.text()}"
            layers = get_layers(path)
            self.layer.addItems(layers)
        else:
            return

    def get_selected_item_text(self, list_widget):
        if list_widget.selectedItems():
            return list_widget.selectedItems()[0].text()
        else:
            QMessageBox.warning(self, "Warning", f"{list_widget.objectName()} not selected.")
            return ""

def main():
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()

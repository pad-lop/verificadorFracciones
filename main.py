from PyQt5 import QtWidgets, uic, QtCore
from PyQt5.QtGui import QFont, QFontDatabase
import sys

import json
from dbfread import DBF
import pandas as pd

import re
import os
import datetime

def clean_value(value):
    if isinstance(value, str):
        # Remove all control characters and non-printable characters
        cleaned_value = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]+', '', value)
        return cleaned_value.strip()  # Remove leading/trailing spaces
    return value

class CatalogoSucursales(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        uic.loadUi("CatalogoSucursales.ui", self)

        self.cargarSucursales()
        self.show()
        self.setWindowTitle("Verificador Fracciones")
        
        self.pushButton_carpeta.clicked.connect(self.selectFolder)
        self.pushButton_agregarSucursal.clicked.connect(self.agregarSucursal)
        
        self.tableWidget.itemChanged.connect(self.guardarCambios)

    def cargarSucursales(self):
        try:
            with open("sucursales.json", "r") as json_file:
                self.local_storage = json.load(json_file)
        except FileNotFoundError:
            with open("sucursales.json", "w") as json_file:
                empty_data = []
                json.dump(empty_data, json_file)
                self.local_storage = empty_data

        self.tableWidget.setRowCount(len(self.local_storage))
        
        for row, value in enumerate(self.local_storage):
            item = QtWidgets.QTableWidgetItem(value["sucursal"])
            self.tableWidget.setItem(row, 0, item)

            item = QtWidgets.QTableWidgetItem(value["carpeta"])
            self.tableWidget.setItem(row, 1, item)
            
            delete_button = QtWidgets.QPushButton("Eliminar")
            delete_button.clicked.connect(self.eliminarFila)
            self.tableWidget.setCellWidget(row, 2, delete_button)

        self.tableWidget.resizeColumnsToContents()

    def selectFolder(self):
        options = QtWidgets.QFileDialog.Options()
        folder_name = QtWidgets.QFileDialog.getExistingDirectory(self, "Select Folder", "", options=options)
        
        if folder_name:
            self.lineEdit_carpeta.setText(folder_name)

    def agregarSucursal(self):
        nueva_sucursal = {
            "sucursal": self.comboBox_sucursal.currentText(),
            "carpeta": self.lineEdit_carpeta.text()
        }

        if nueva_sucursal["sucursal"] and nueva_sucursal["carpeta"]:
            self.local_storage.append(nueva_sucursal)
            
            with open("sucursales.json", "w") as json_file:
                json.dump(self.local_storage, json_file)

            self.cargarSucursales()

    def guardarCambios(self, item):
        row = item.row()
        column = item.column()

        self.tableWidget.resizeColumnsToContents()
        
        if column == 0:
            self.local_storage[row]['sucursal'] = item.text()
        elif column == 1:
            self.local_storage[row]['carpeta'] = item.text()
        
        with open("sucursales.json", "w") as json_file:
            json.dump(self.local_storage, json_file)

    def eliminarFila(self):
        sender_button = self.sender()
        if sender_button:
            index = self.tableWidget.indexAt(sender_button.pos())
            if index.isValid():
                row = index.row()
                
                self.local_storage.pop(row)
                
                with open("sucursales.json", "w") as json_file:
                    json.dump(self.local_storage, json_file)
                
                self.cargarSucursales()




class VerificadorFracciones(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        uic.loadUi("main.ui", self)

        self.show()

        self.pushButton_Fraccionados.clicked.connect(self.reporteFraccionados)
        self.pushButton_NoFraccionados.clicked.connect(self.reporteNoFraccionados)
        self.pushButton_CatalogoSucursales.clicked.connect(self.abrirCatalogoSucursales)

        self.sacarLista()

    def sacarLista(self):
        try:
            with open("sucursales.json", "r") as json_file:
                self.sucursales = json.load(json_file)

        except FileNotFoundError:
            with open("sucursales.json", "w") as json_file:
                empty_data = []
                json.dump(empty_data, json_file)
                self.sucursales = empty_data

        layout = QtWidgets.QVBoxLayout()

        self.button_group = QtWidgets.QButtonGroup(self)  # Create a button group instance
        
        if len(self.sucursales) == 0:
            label = QtWidgets.QLabel()
            label.setText("No hay sucursales registradas")
            layout.addWidget(label)
            

        else:
            for index, item in enumerate(self.sucursales): 
                button = QtWidgets.QRadioButton(f"{item['sucursal']}")
                layout.addWidget(button)
                self.button_group.addButton(button, index)  
                button.clicked.connect(self.seleccionarSucursal)  

        self.groupBox_sucursales.setLayout(layout)
        self.show()

    def seleccionarSucursal(self):
        try:
            with open("sucursales.json", "r") as json_file:
                self.local_storage = json.load(json_file)
        except FileNotFoundError:
            with open("sucursales.json", "w") as json_file:
                empty_data = []
                json.dump(empty_data, json_file)
                self.local_storage = empty_data

        sucursalSeleccionada = self.button_group.checkedButton().text()

        for sucursal in self.local_storage:
            if sucursal["sucursal"] == sucursalSeleccionada: 
                self.sucursalEncontrada = sucursal


    def reporteFraccionados(self):

        if hasattr(self, 'sucursalEncontrada'):

            docum_path = self.sucursalEncontrada['carpeta'] + "/Unidades.DBF"
            table = DBF(docum_path, ignore_missing_memofile=True)
            df_fraccionados = pd.DataFrame(iter(table))
            df_fraccionados['NUMART'] = df_fraccionados['NUMART'].str.lstrip()

            os.makedirs('Reportes Excel', exist_ok=True)
            
            current_date = datetime.datetime.now().strftime("%d-%m-%Y %H-%M")
            fileName = f"{current_date} {self.sucursalEncontrada['sucursal']} Fraccionados.xlsx"

            folder_name = os.path.join(os.getcwd(), 'Reportes Excel')

            # Define the path to save the Excel file
            excel_file_path = os.path.join(folder_name, fileName)

            # Save the DataFrame to the Excel file
            df_fraccionados.to_excel(excel_file_path, index=False)
            
            QtWidgets.QMessageBox.information(None, "Reporte Terminado", f"Fraccionados.xlsx guardado.", QtWidgets.QMessageBox.Ok)


            # Open the Excel file using a platform-specific method (e.g., on Windows)
            if os.name == 'nt':
                os.system(f'start excel "{excel_file_path}"')
            else:
                QtWidgets.QMessageBox.critical(None, "Error", "No se puede abir excel", QtWidgets.QMessageBox.Ok)



        else:
            QtWidgets.QMessageBox.critical(None, "Error", "Seleccione una sucursal", QtWidgets.QMessageBox.Ok)


            

  

    def reporteNoFraccionados(self):

        if hasattr(self, 'sucursalEncontrada'):

            docum_path = self.sucursalEncontrada['carpeta'] + "/Arts.DBF"
            table = DBF(docum_path, ignore_missing_memofile=True)
            df_articulos = pd.DataFrame(iter(table))
            df_articulos = df_articulos.applymap(clean_value)

            docum_path = self.sucursalEncontrada['carpeta'] + "/Unidades.DBF"
            table = DBF(docum_path, ignore_missing_memofile=True)
            df_fraccionados = pd.DataFrame(iter(table))
            df_fraccionados['NUMART'] = df_fraccionados['NUMART'].str.lstrip()

            df_fraccionados = df_fraccionados.drop_duplicates(subset='NUMART')
            df_noFraccionados = df_articulos[~df_articulos['NUMART'].isin(df_fraccionados['NUMART'])]

            os.makedirs('Reportes Excel', exist_ok=True)
            
            current_date = datetime.datetime.now().strftime("%d-%m-%Y %H-%M")
            fileName = f"{current_date} {self.sucursalEncontrada['sucursal']} NoFraccionados.xlsx"

            folder_name = os.path.join(os.getcwd(), 'Reportes Excel')

            # Define the path to save the Excel file
            excel_file_path = os.path.join(folder_name, fileName)

            # Save the DataFrame to the Excel file
            df_noFraccionados.to_excel(excel_file_path, index=False)
            QtWidgets.QMessageBox.information(None, "Reporte Terminado", f"NoFraccionados.xlsx guardado", QtWidgets.QMessageBox.Ok)

            # Open the Excel file using a platform-specific method (e.g., on Windows)
            if os.name == 'nt':
                os.system(f'start excel "{excel_file_path}"')
            else:
                QtWidgets.QMessageBox.critical(None, "Error", "No se puede abir excel", QtWidgets.QMessageBox.Ok)

        else:
            QtWidgets.QMessageBox.critical(None, "Error", "Seleccione una sucursal", QtWidgets.QMessageBox.Ok)



    def abrirCatalogoSucursales(self):
        self.ventanaSucursales = CatalogoSucursales()
        self.ventanaSucursales.show()



if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = VerificadorFracciones()

    ubuntu_font_id = QFontDatabase.addApplicationFont("Ubuntu-L.ttf")
    if ubuntu_font_id != -1:
        font_family = QFontDatabase.applicationFontFamilies(ubuntu_font_id)[0]
        app.setFont(QFont(font_family))

    sys.exit(app.exec())

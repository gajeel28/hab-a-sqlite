from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QMessageBox
import sys
import os
import sqlite3
import pandas as pd
import Ui_Menu

class MainUiClass (QtWidgets.QMainWindow, Ui_Menu.Ui_MainWindow):
    def __init__(self, parent=None):
        super (MainUiClass, self).__init__(parent)
        self.setupUi(self)

class Aplicacion (object):
    def __init__(self, ui):
        self.ui = ui

    def btn_hab_click(self):
        try:
            filename = QtWidgets.QFileDialog.getOpenFileName(filter = "HAB (*.hab)", directory=r'')
        except:
            msgBox_err_filename = QMessageBox(icon = QMessageBox.Critical, text="Error al intentar abrir el directorio.")
            msgBox_err_filename.setWindowTitle("Error")
            msgBox_err_filename.exec_()
            return

        if filename[0]:
            self.archivo = filename[0][::-1]
            busca_barra = self.archivo.find('/')
            self.archivo = self.archivo[:busca_barra]
            self.archivo = self.archivo[::-1]

            try:
                conn = sqlite3.connect('pagos_benef_hab.db')
                c = conn.cursor()
                c.execute('SELECT N_ARCHIVO FROM T_ARCHIVOS WHERE N_ARCHIVO = ' + "\'" + self.archivo + "\'")
                x = c.fetchone()
            except:
                msgBox_err_conn = QMessageBox(icon=QMessageBox.Critical,text="Error al abrir la Base De Datos.")
                msgBox_err_conn.setWindowTitle("Error")
                msgBox_err_conn.exec_()
                conn.close()
                return

            if x:
                msg = QMessageBox()
                msg.setText("Archivo ya existente el la Base de Datos")
                msg.setIcon(QMessageBox.Warning)
                msg.setWindowTitle("¡Atención!")
                msg.exec_()
                return

            self.ui.lbl_hab.setText(filename[0])

            try:
                self.df_hab = pd.read_fwf(filename[0], widths = [3,5,2,1,9,18,8,5,6,22,2,22,43], header=None)
            except:
                msgBox_err_hab = QMessageBox(icon=QMessageBox.Critical,text="Error al abrir el archivo.")
                msgBox_err_hab.setWindowTitle("Error")
                msgBox_err_hab.exec_()
                return

            self.df_hab = self.df_hab.rename(columns={0:'TIPO_CONVENIO', 1:'SUCURSAL', 2:'MONEDA', 3:'SISTEMA', 4:'NRO_CTA', 5:'IMPORTE', 6:'FECHA', 7:'CONVENIO_EMPRESA', 8:'NRO_COMPROBANTE', 9:'CBU', 10:'CUOTA', 11:'CUIT', 12:'ESTADO'})
            self.ui.lbl_cant_registros.setText(str(self.df_hab.shape[0]))
            self.ui.btn_base.setEnabled(True)
            self.ui.txt_programa.setEnabled(True)

    def btn_base_click(self):
        if self.ui.txt_programa.text() == '':
            msgBox_programa_vacio = QMessageBox()
            msgBox_programa_vacio.setText('Error el programa no puede estar Vacío.')
            msgBox_programa_vacio.setIcon(QMessageBox.Critical)
            msgBox_programa_vacio.setWindowTitle("Error")
            msgBox_programa_vacio.exec_()
            return

        try:
            conn = sqlite3.connect('pagos_benef_hab.db')
            c = conn.cursor()
            c.execute('INSERT INTO T_ARCHIVOS VALUES (NULL, ' + "\'" + self.archivo + "\')")
            conn.commit()
            c.execute('SELECT ID_ARCHIVO FROM T_ARCHIVOS WHERE N_ARCHIVO = ' + "\'" + self.archivo + "\'")
            id_archivo = c.fetchone()
        

            self.df_hab['ID_ARCHIVO'] = id_archivo[0]
            self.df_hab['PROGRAMA'] = self.ui.txt_programa.text()
            self.df_hab['CUIT'] = self.df_hab['CUIT'].str.slice(start = - 11)
            self.df_hab.to_sql('T_RESPUESTA_BANCO', conn, index=False, if_exists='append')
            self.df_hab.to_excel(self.archivo + '.xlsx', index=False)
        except:
            msgBox_err_conn = QMessageBox(icon=QMessageBox.Critical,text="Error en la Base De Datos.")
            msgBox_err_conn.setWindowTitle("Error")
            msgBox_err_conn.exec_()
            conn.close()
            return

        self.ui.btn_base.setEnabled(False)
        self.ui.lbl_cant_registros.setText("")
        self.ui.lbl_hab.setText("Ubicación del archivo")
        self.ui.txt_programa.setText("")

        msgBox_finalizado = QMessageBox()
        msgBox_finalizado.setText("Proceso finalizado con éxito.")
        msgBox_finalizado.setIcon(QMessageBox.Information)
        msgBox_finalizado.setWindowTitle("¡Atención!")
        msgBox_finalizado.exec_()
    
    def btn_abrir_directorio_click(self):
        os.system(r"start.")

def main():
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
        ui = MainUiClass()
        apli = Aplicacion(ui)
        ui.show()
        ui.btn_hab.clicked.connect(apli.btn_hab_click)
        ui.btn_base.clicked.connect(apli.btn_base_click)
        ui.btn_base.setEnabled(False)
        ui.txt_programa.setEnabled(False)
        ui.btn_abrir_directorio.clicked.connect(apli.btn_abrir_directorio_click)
        app.exec_()


if __name__ == "__main__":
    main()
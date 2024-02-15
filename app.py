import sys
import os
import re
import win32com.client as win32
import openpyxl
from PyQt5 import uic
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog, QMessageBox
from PyQt5.QtGui import QIcon

class AutomatizacionThread(QThread):
    progressUpdated = pyqtSignal(int)
    finished = pyqtSignal()
    error = pyqtSignal()

    def __init__(self, rutaExcel, rutaArchivos, asunto, mensaje, emisor, cc):
        super().__init__()
        self.rutaExcel = rutaExcel
        self.rutaArchivos = rutaArchivos
        self.asunto = asunto
        self.mensaje = mensaje
        self.emisor = emisor
        self.cc = cc

    def run(self):
        self.progressBarValue = 0
        contador = 0

        excel = openpyxl.load_workbook(self.rutaExcel)
        hoja = excel['Hoja1']

        for fila in range(1, len(hoja["A"]) + 1):
            destinatario = str(hoja["D" + str(fila)].value)
            asunto = self.asunto
            cuerpo = self.mensaje.toHtml()

            archivos_adjuntos = str(hoja["E" + str(fila)].value)

            patron = r"--GRADO--"
            cuerpo = re.sub(patron, str(hoja["B" + str(fila)].value), cuerpo)

            patron = r"--NOMBRE--"
            cuerpo = re.sub(patron, str(hoja["C" + str(fila)].value), cuerpo)

            patron = r"--GENERAL--"
            cuerpo = re.sub(patron, str(hoja["F" + str(fila)].value), cuerpo)

            try:
                self.enviar_correo(destinatario, asunto, cuerpo, archivos_adjuntos)
            except Exception as e:
                self.error.emit()
                return

            contador += 1
            self.progressBarValue = round((contador / len(hoja["A"])) * 100)
            self.progressUpdated.emit(self.progressBarValue)

        self.finished.emit()

    def enviar_correo(self, destinatario, asunto, cuerpo, archivos_adjuntos):
        outlook = win32.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace('MAPI')
        carpeta_enviados = namespace.GetDefaultFolder(5)  # Carpeta de elementos enviados

        oacctouse = None
        for oacc in outlook.Session.Accounts:
            if oacc.SmtpAddress == self.emisor:
                oacctouse = oacc
                break

        correo = outlook.CreateItem(0)  # 0: correo normal, 1: correo HTML, 2: correo sin formato

        if oacctouse:
            correo._oleobj_.Invoke(*(64209, 0, 8, 0, oacctouse))  # Msg.SendUsingAccount = oacctouse

        correo.Subject = asunto
        correo.HTMLBody = cuerpo
        correo.To = destinatario
        if self.cc != None:
            correo.CC = self.cc

        directorio_actual = self.rutaArchivos

        archivos_adjuntos = archivos_adjuntos.split(",")
        archivos_limpios = [archivo.strip() for archivo in archivos_adjuntos]
        
        for archivo in archivos_limpios:
            ruta_absoluta = os.path.join(directorio_actual, archivo)
            # Adjuntar archivo
            adjunto = correo.Attachments.Add(ruta_absoluta)

        # Enviar el correo
        correo.Send()

        print('Correo enviado correctamente.')


class Masificador(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi("main.ui", self)
        icon = QIcon("correo.ico")
        self.setWindowIcon(icon)

        self.rutaExcel = None
        self.rutaArchivos = None
        self.asunto = None
        self.mensaje = None
        self.correos = []
        self.correoSeleccionado = None

        self.comboBoxCorreos.addItem("-")
        self.obtener_correos()

        self.pushButtonExcel.clicked.connect(self.buscarExcel)
        self.pushButtonRuta.clicked.connect(self.buscarRuta)
        self.pushButtonEnviar.clicked.connect(self.enviar)

        self.comboBoxCorreos.currentIndexChanged.connect(self.combo_box_changed)

    def combo_box_changed(self):
        selected_option = self.comboBoxCorreos.currentText()
        print("Opción seleccionada:", selected_option)
        self.correoSeleccionado = selected_option

    def obtener_correos(self):
        outlookMAIN = win32.Dispatch('Outlook.Application')
        oacctouse = None
        for oacc in outlookMAIN.Session.Accounts:
            self.correos.append(oacc.SmtpAddress)
            self.comboBoxCorreos.addItem(oacc.SmtpAddress)

    def buscarExcel(self):
        fileName = QFileDialog.getOpenFileName(self, "Abrir archivo", "C:", "Archivo de Excel(*.xlsx)")
        self.rutaExcel = fileName[0]
        self.lineEditExcel.setText(fileName[0])

    def buscarRuta(self):
        fileName = QFileDialog.getExistingDirectory(self, "Seleccionar ruta", "C:")
        self.rutaArchivos = fileName
        self.lineEditRuta.setText(fileName)

    def enviar(self):
        if self.rutaExcel and self.rutaArchivos and self.correoSeleccionado and self.correoSeleccionado != "-" and self.lineEditAsunto.text() and self.textEditMensaje.toPlainText():
            print("Enviando...")

            self.pushButtonExcel.setEnabled(False)
            self.pushButtonRuta.setEnabled(False)
            self.pushButtonEnviar.setEnabled(False)

            self.comboBoxCorreos.setEnabled(False)
            self.lineEditCC.setReadOnly(True)

            self.lineEditAsunto.setReadOnly(True)
            self.textEditMensaje.setReadOnly(True)

            self.automatizacionThread = AutomatizacionThread(
                self.rutaExcel,
                self.rutaArchivos,
                self.lineEditAsunto.text(),
                self.textEditMensaje,
                self.correoSeleccionado,
                self.lineEditCC.text()
            )
            self.automatizacionThread.progressUpdated.connect(self.updateProgressBar)
            self.automatizacionThread.finished.connect(self.processFinished)
            self.automatizacionThread.error.connect(self.processError)
            self.automatizacionThread.start()
        else:
            QMessageBox.warning(self, "Aviso", "Todos los campos son obligatorios, favor de llenar todos.")

    def processError(self):
        QMessageBox.critical(self, "ERROR", "Ha ocurrido un error, asegurese de tener Outlook abierto y reinicie el programa, si lo tiene abierto y aún asi sigue viendo este mensaje, por favor comuniquese con el desarrollador.")

    def updateProgressBar(self, value):
        self.progressBar.setValue(value)

    def processFinished(self):
        self.pushButtonExcel.setEnabled(True)
        self.pushButtonRuta.setEnabled(True)
        self.pushButtonEnviar.setEnabled(True)

        self.comboBoxCorreos.setEnabled(True)
        self.lineEditCC.setReadOnly(False)

        self.lineEditAsunto.setReadOnly(False)
        self.textEditMensaje.setReadOnly(False)

        self.lineEditExcel.clear()
        self.lineEditRuta.clear()
        self.lineEditAsunto.clear()
        self.textEditMensaje.clear()
        self.lineEditCC.clear()

        print("Envío de correos completado.")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    gui = Masificador()
    gui.show()
    sys.exit(app.exec_())
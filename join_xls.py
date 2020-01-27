from PyQt5 import QtWidgets, uic
from PyQt5.QtWidgets import QMessageBox
import sys
import pandas as pd


class Ui(QtWidgets.QMainWindow):
    def __init__(self):
        super(Ui, self).__init__()

        uic.loadUi('mainWindow.ui', self)

        self.btnPickSource = self.findChild(QtWidgets.QPushButton, 'btnPickSource')
        self.btnPickTarget = self.findChild(QtWidgets.QPushButton, 'btnPickTarget')
        self.btnStart = self.findChild(QtWidgets.QPushButton, 'btnStart')
        self.lblSourcePath = self.findChild(QtWidgets.QLabel, 'lblSourcePath')
        self.lblTargetPath = self.findChild(QtWidgets.QLabel, 'lblTargetPath')

        self.btnPickSource.clicked.connect(self.pickFiles)
        self.btnPickTarget.clicked.connect(self.setTargetFile)
        self.btnStart.clicked.connect(self.joinFiles)

        self.files = []
        self.target = ''

        self.show()

    def pickFiles(self):
        options = QtWidgets.QFileDialog.Options()
        files, _ = QtWidgets.QFileDialog.getOpenFileNames(self, "Wybierz pliki do scalenia...", "",
                                                          "Excel Files (*.xls)", options=options)
        if files:
            self.lblSourcePath.setText(f'{files[0]} + {len(files)-1} plików')
            self.files = files

    def setTargetFile(self):
        options = QtWidgets.QFileDialog.Options()
        fileName, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Plik docelowy...", "",
                                                            "Excel Files (*.xls)", options=options)
        if fileName:
            if not fileName.endswith('.xls'):
                fileName += '.xls'
            self.lblTargetPath.setText(fileName)
            self.target = fileName

    def joinFiles(self):
        done = False
        try:
            if len(self.files) > 0 and len(self.target) > 0:
                if len(self.files)>=2:
                    done = executeJoin(self.files, self.target)
                else:
                    self.showPopup('Wybierz więcej, niż jeden plik', 'Błąd!', QMessageBox.Warning)
            else:
                self.showPopup('Wybierz pliki do scalenia oraz podaj lokalizację i nazwę nowego pliku!', 'BŁĄD!!!', QMessageBox.Warning)
            if done:
                self.showPopup(f'Scalono {len(self.files)+1} plików do pliku:\n{self.target}', 'Sukces!')
        except Exception as e:
            self.showPopup(f'Skontaktuj się z twórcą programu. Prześlij mu poniższy tekst:\n{e}', 'BŁĄD KRYTYCZNY!!!',
                           QMessageBox.Critical)
            print(e)

    def showPopup(self, text, msgHeader='Info', icon=QMessageBox.Information):
        msg = QMessageBox()
        msg.setWindowTitle(msgHeader)
        msg.setText(text)
        msg.setIcon(icon)

        x = msg.exec_()


def executeJoin(filelist, targetFile):
    try:
        result = pd.read_excel(filelist.pop(), header=1)
    except Exception as e:
        print('Błąd: ', e)

    for file in filelist:
        try:
            file_content = pd.read_excel(file, header=1)
            result = pd.concat([result, file_content])
        except Exception as e:
            print('Błąd: ', e)

    try:
        result.to_excel(targetFile, index=False)
    except Exception as e:
        print('Błąd: ', e)

    return True


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = Ui()
    app.exec_()

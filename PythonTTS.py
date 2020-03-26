import sys
from time import sleep

import win32com.client as wincl
from PyQt5.QtCore import QThread
from PyQt5.QtWidgets import QMainWindow, QApplication

import MainWindow


class readThread(QThread):
    def __init__(self, text):
        super().__init__()
        self.text = text

    def run(self) -> None:
        import pythoncom
        pythoncom.CoInitialize()
        speak = wincl.Dispatch("SAPI.SpVoice")
        speak.Speak(self.text)


class MainWin(QMainWindow):
    __flag = True
    __beginIndex = 0
    __strList = []
    __listLen = 0

    def __init__(self):
        super().__init__()
        self.ui = MainWindow.Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.textEdit.textChanged.connect(self.changedText)

    def changedText(self):
        if self.__beginIndex > len(self.ui.textEdit.toPlainText()):
            self.__beginIndex = len(self.ui.textEdit.toPlainText())
        else:
            if self.__flag:
                temp = self.ui.textEdit.toPlainText()
                self.__flag = False
            else:
                temp = self.ui.textEdit.toPlainText()[self.__beginIndex:]
            self.__beginIndex += len(temp)
            print(temp)
            myReadThread = readThread(temp)
            myReadThread.start()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    main = MainWin()
    main.show()
    sys.exit(app.exec_())

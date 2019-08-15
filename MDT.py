#coding=utf-8
import sys
import os
#from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtWidgets import *
from MDT_CI import *
from MDT_FieldInfo_CBTC import CI_MDT_LIST_C
from MDT_FieldInfo_UTO import CI_MDT_LIST_U

class MyMainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MyMainWindow, self).__init__(parent)
        self.setupUi(self)

    def xuanzewenjian(self, event):
        tmpDir = os.path.dirname(os.path.realpath(__file__))
        fileName1, filetype = QFileDialog.getOpenFileName(self, "选取文件", tmpDir.format(),
                                                          "Excel Files (*.xlsx;*.xls);;Text Files (*.txt)")  # 设置文件扩展名过滤,注意用双分号间隔
        tmp = os.path.basename(fileName1)
        self.lineEdit.setText(tmp)

    def shengchengCFG(self, event):
        tmp = self.lineEdit.text()
        if tmp == "":
            QMessageBox.information(self, "错误", "未选择文件", QMessageBox.Ok)
        else :
            if self.radioButton.isChecked():
                CI_MDT_LIST_C(tmp)
                QMessageBox.information(self, "完成", "生成MDT配置成功！", QMessageBox.Ok)
            elif self.radioButton_2.isChecked():
                CI_MDT_LIST_U(tmp)
                QMessageBox.information(self, "完成", "生成MDT配置成功！", QMessageBox.Ok)
            else :
                QMessageBox.information(self, "错误", "未选择版本", QMessageBox.Ok)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    myWin = MyMainWindow()
    myWin.show()
    sys.exit(app.exec_())


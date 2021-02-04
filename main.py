from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog
import os
import sys
import pandas as pd


class Ui_MainWindow(object):


    columns_changed=""
    excel_file_changed=""

    columns_compare=""
    excel_file_compare=""

    def browse_Changed(self):

        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(None, "Open a File", "C:\\Users\\MSI\\Desktop\\Excel_Chan\\excel\\",
                                                  "All Files (*);;Excel 1997-2003 Files (*.xls);;Excel Files (*.xlsx)",
                                                  options=options)


        if fileName != "":

            self.lineEdit_DegistirilecekVeri.setText(fileName)
            self.excel_file_changed = pd.read_excel(fileName, dtype='str')
            self.columns_changed = self.excel_file_changed.columns.ravel()

            for column in self.columns_changed:
                self.comboBox_DegistirelecekVeri.addItem(column)

            data_changed = self.comboBox_DegistirelecekVeri.currentText()



    def browse_Compare(self):

        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(None, "Open a File", "C:\\Users\\MSI\\Desktop\\Excel_Chan\\excel\\",
                                                  "All Files (*);;Excel 1997-2003 Files (*.xls);;Excel Files (*.xlsx)",
                                                  options=options)

        if fileName != "":
            self.lineEdit_KarsilastirilacakVeri.setText(fileName)
            self.excel_file_compare = pd.read_excel(fileName, dtype='str')
            self.columns_compare = self.excel_file_compare.columns.ravel()


            for column in self.columns_compare:
                self.comboBox_KarsilastirilacakVeri.addItem(column)

            data_compare = self.comboBox_KarsilastirilacakVeri.currentText()





    def Excel_Change(self):

        data_changed = self.comboBox_DegistirelecekVeri.currentText()
        data_compare = self.comboBox_KarsilastirilacakVeri.currentText()

        data_changed_array_id = self.excel_file_changed[f'{self.columns_changed[0]}'].tolist()

        data_changed_array = self.excel_file_changed[f'{data_changed}'].tolist()  #Seçili olan Itemların verisi alındı.
        data_compare_array = self.excel_file_compare[f'{data_compare}'].tolist()

        print("Data Changed",data_changed_array)
        print("Data Compare",data_compare_array)

        id_and_data_list = []

        for id_col,selected_item in zip(data_changed_array_id,data_changed_array):
            id_and_data_list.append([id_col,str(selected_item).lower()])


        print("ID_And_Data_List",id_and_data_list)

        modified_data_list = []

        for i in range(0,len(id_and_data_list)):
            for j in range(0,len(data_compare_array)):

                if str(data_compare_array[j]).lower() != "nan":
                    if str(data_compare_array[j]).lower().find(id_and_data_list[i][1]) != -1:

                        print(str(id_and_data_list[i][1]).lower())
                        print(str(data_compare_array[j]).lower())


                        k = i+1
                        a = j+1
                        modified_data_list.append([k,a])


        print("Modified Data List",modified_data_list)


        export_DataFrame = pd.DataFrame(modified_data_list, columns=["ID","İsim"])
        export_sort = export_DataFrame.sort_values(by=["İsim","ID"])
        print("Export DataFrame",export_DataFrame)
        print("-----------------------")
        print("Export Sort",export_sort)
        print("------------------")
        export_columns = export_sort.columns.ravel()
        print("Export Columns",export_columns)
        export_data = export_sort[f"{export_columns[0]}"].tolist()
        print("Export Data",export_data)

        export_DataFrame_Last = pd.DataFrame(export_data, columns=[f"{data_compare}"])
        print("Export DataFrame Last",export_DataFrame_Last)
        export_DataFrame_Last_data = export_DataFrame_Last[f'{data_compare}'].tolist()

        data_compare_array_new = data_compare_array

        dic_replace = {}
        j = 0

        for i in range(0,len(data_compare_array_new)):
            if str(data_compare_array_new[i]).lower() != "nan":

                dic_replace.update({str(data_compare_array_new[i]): str(export_DataFrame_Last_data[j])})

                if j < len(export_DataFrame_Last_data):
                    j = j + 1

                else:
                    break

        data_compare_array_new = [dic_replace.get(x, x) for x in data_compare_array_new]

        export_DataFrame_Last_Last = pd.DataFrame(data_compare_array_new, columns=[f"{data_compare}"])

        print(data_compare_array_new)

        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getSaveFileName(None, "Save A File", "C:\\Users\\",
                                                  "Excel Files (*.xlsx);;Excel 1997-2003 Files (*.xls)",
                                                  options=options)


        filename_export = str(fileName).replace("/","\\\\")
        print(filename_export)
        # export_DataFrame_Last.to_excel(filename_export)
        export_DataFrame_Last_Last.to_excel(filename_export)



    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1065, 417)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.groupBox_DegistirelecekVeri = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_DegistirelecekVeri.setGeometry(QtCore.QRect(10, 20, 501, 311))
        self.groupBox_DegistirelecekVeri.setObjectName("groupBox_DegistirelecekVeri")
        self.groupBox_2 = QtWidgets.QGroupBox(self.groupBox_DegistirelecekVeri)
        self.groupBox_2.setGeometry(QtCore.QRect(500, 710, 501, 711))
        self.groupBox_2.setObjectName("groupBox_2")
        self.comboBox_DegistirelecekVeri = QtWidgets.QComboBox(self.groupBox_DegistirelecekVeri)
        self.comboBox_DegistirelecekVeri.setGeometry(QtCore.QRect(80, 70, 331, 21))
        self.comboBox_DegistirelecekVeri.setObjectName("comboBox_DegistirelecekVeri")

        self.comboBox_DegistirelecekVeri.addItem(" ")
        index = self.comboBox_DegistirelecekVeri.findText(" ", QtCore.Qt.MatchFixedString)
        self.comboBox_DegistirelecekVeri.setCurrentIndex(index)

        self.lineEdit_DegistirilecekVeri = QtWidgets.QLineEdit(self.groupBox_DegistirelecekVeri)
        self.lineEdit_DegistirilecekVeri.setGeometry(QtCore.QRect(10, 30, 401, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.lineEdit_DegistirilecekVeri.setFont(font)
        self.lineEdit_DegistirilecekVeri.setInputMask("")
        self.lineEdit_DegistirilecekVeri.setReadOnly(True)
        self.lineEdit_DegistirilecekVeri.setObjectName("lineEdit_DegistirilecekVeri")
        self.pushButton_Browse_Degistirlecek = QtWidgets.QPushButton(self.groupBox_DegistirelecekVeri)
        self.pushButton_Browse_Degistirlecek.setGeometry(QtCore.QRect(420, 30, 71, 21))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.pushButton_Browse_Degistirlecek.setFont(font)
        self.pushButton_Browse_Degistirlecek.setObjectName("pushButton_Browse_Degistirlecek")
        self.groupBox_3 = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_3.setGeometry(QtCore.QRect(550, 20, 501, 311))
        self.groupBox_3.setObjectName("groupBox_3")
        self.groupBox_4 = QtWidgets.QGroupBox(self.groupBox_3)
        self.groupBox_4.setGeometry(QtCore.QRect(500, 710, 501, 711))
        self.groupBox_4.setObjectName("groupBox_4")
        self.pushButton_Browse_Karsilastirilacak = QtWidgets.QPushButton(self.groupBox_3)
        self.pushButton_Browse_Karsilastirilacak.setGeometry(QtCore.QRect(430, 30, 71, 21))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.pushButton_Browse_Karsilastirilacak.setFont(font)
        self.pushButton_Browse_Karsilastirilacak.setObjectName("pushButton_Browse_Karsilastirilacak")
        self.lineEdit_KarsilastirilacakVeri = QtWidgets.QLineEdit(self.groupBox_3)
        self.lineEdit_KarsilastirilacakVeri.setGeometry(QtCore.QRect(20, 30, 401, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.lineEdit_KarsilastirilacakVeri.setFont(font)
        self.lineEdit_KarsilastirilacakVeri.setReadOnly(True)
        self.lineEdit_KarsilastirilacakVeri.setObjectName("lineEdit_KarsilastirilacakVeri")
        self.comboBox_KarsilastirilacakVeri = QtWidgets.QComboBox(self.groupBox_3)
        self.comboBox_KarsilastirilacakVeri.setGeometry(QtCore.QRect(90, 70, 331, 21))
        self.comboBox_KarsilastirilacakVeri.setObjectName("comboBox_KarsilastirilacakVeri")

        self.comboBox_KarsilastirilacakVeri.addItem(" ")
        index = self.comboBox_KarsilastirilacakVeri.findText(" ", QtCore.Qt.MatchFixedString)
        self.comboBox_KarsilastirilacakVeri.setCurrentIndex(index)

        self.pushButton_OK = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_OK.setGeometry(QtCore.QRect(470, 360, 121, 31))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.pushButton_OK.setFont(font)
        self.pushButton_OK.setObjectName("pushButton_OK")

        self.pushButton_Browse_Degistirlecek.clicked.connect(self.browse_Changed)
        self.pushButton_Browse_Karsilastirilacak.clicked.connect(self.browse_Compare)
        self.pushButton_OK.clicked.connect(self.Excel_Change)

        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Excel Changer"))
        self.groupBox_DegistirelecekVeri.setTitle(_translate("MainWindow", "Değiştirilecek Veri"))
        self.groupBox_2.setTitle(_translate("MainWindow", "GroupBox"))
        self.pushButton_Browse_Degistirlecek.setText(_translate("MainWindow", "..."))
        self.groupBox_3.setTitle(_translate("MainWindow", "Karşılaştırılacak Veri"))
        self.groupBox_4.setTitle(_translate("MainWindow", "GroupBox"))
        self.pushButton_Browse_Karsilastirilacak.setText(_translate("MainWindow", "..."))
        self.pushButton_OK.setText(_translate("MainWindow", "OK"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

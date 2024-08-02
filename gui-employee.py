import pandas as pd
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog
import os
import mysql.connector



class Ui_MainWindow(object):
	pathFile = ''
	nameFile = ''
	fileSize = ''
	countTotalRows = ''
	columnstotal = ''
	mydb = ''
	values = ''
	data = ''
	list_excel_headers = ''
	# config = configparser.ConfigParser()
	# config.read('config.ini', 'utf8')
	appName = 'Insert data from excel to mysql'






	def setupUi(self, MainWindow):
		MainWindow.setObjectName(self.appName)
		MainWindow.setEnabled(True)
		# MainWindow.resize(501, 467)
		MainWindow.setLayoutDirection(QtCore.Qt.RightToLeft)
		MainWindow.setStyleSheet("")
		MainWindow.setFixedSize(501, 467)
		MainWindow.setWindowIcon(QtGui.QIcon('ico.png'))
		self.centralwidget = QtWidgets.QWidget(MainWindow)
		self.centralwidget.setObjectName("centralwidget")
		self.plainTextEdit = QtWidgets.QTextEdit(self.centralwidget)
		self.plainTextEdit.setEnabled(True)
		self.plainTextEdit.setGeometry(QtCore.QRect(40, 270, 441, 171))
		self.plainTextEdit.setContextMenuPolicy(QtCore.Qt.DefaultContextMenu)
		self.plainTextEdit.setLayoutDirection(QtCore.Qt.RightToLeft)
		self.plainTextEdit.setStyleSheet("direction:rtl;\n"
		                                 "text-align: right;\n"
		                                 "position: absolute;\n"
		                                 "right: 0;")
		self.plainTextEdit.setLocale(QtCore.QLocale(QtCore.QLocale.Persian, QtCore.QLocale.Iran))
		self.plainTextEdit.setFrameShadow(QtWidgets.QFrame.Plain)
		self.plainTextEdit.setMidLineWidth(0)
		self.plainTextEdit.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
		self.plainTextEdit.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
		self.plainTextEdit.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
		self.plainTextEdit.setReadOnly(True)
		self.plainTextEdit.setObjectName("plainTextEdit")
		self.groupBox = QtWidgets.QGroupBox(self.centralwidget)
		self.groupBox.setGeometry(QtCore.QRect(10, 50, 241, 171))
		self.groupBox.setObjectName("groupBox")
		self.lineEdit = QtWidgets.QLineEdit(self.groupBox)
		self.lineEdit.setGeometry(QtCore.QRect(70, 20, 111, 21))
		self.lineEdit.setObjectName("lineEdit")
		self.label = QtWidgets.QLabel(self.groupBox)
		self.label.setGeometry(QtCore.QRect(10, 20, 47, 21))
		self.label.setObjectName("label")
		self.lineEdit_2 = QtWidgets.QLineEdit(self.groupBox)
		self.lineEdit_2.setGeometry(QtCore.QRect(70, 50, 111, 21))
		self.lineEdit_2.setObjectName("lineEdit_2")
		self.label_5 = QtWidgets.QLabel(self.groupBox)
		self.label_5.setGeometry(QtCore.QRect(10, 50, 47, 21))
		self.label_5.setObjectName("label_5")
		self.label_6 = QtWidgets.QLabel(self.groupBox)
		self.label_6.setGeometry(QtCore.QRect(10, 80, 47, 21))
		self.label_6.setObjectName("label_6")
		self.lineEdit_3 = QtWidgets.QLineEdit(self.groupBox)
		self.lineEdit_3.setGeometry(QtCore.QRect(70, 80, 111, 21))
		self.lineEdit_3.setObjectName("lineEdit_3")
		self.label_7 = QtWidgets.QLabel(self.groupBox)
		self.label_7.setGeometry(QtCore.QRect(10, 110, 47, 21))
		self.label_7.setObjectName("label_7")
		self.lineEdit_4 = QtWidgets.QLineEdit(self.groupBox)
		self.lineEdit_4.setGeometry(QtCore.QRect(70, 110, 111, 21))
		self.lineEdit_4.setObjectName("lineEdit_4")
		self.pushButton = QtWidgets.QPushButton(self.groupBox)
		self.pushButton.setGeometry(QtCore.QRect(140, 140, 91, 23))
		self.pushButton.setObjectName("pushButton")
		self.label_8 = QtWidgets.QLabel(self.groupBox)
		self.label_8.setGeometry(QtCore.QRect(20, 140, 91, 21))
		self.label_8.setObjectName("label_8")
		self.groupBox_2 = QtWidgets.QGroupBox(self.centralwidget)
		self.groupBox_2.setGeometry(QtCore.QRect(270, 50, 211, 61))
		self.groupBox_2.setStyleSheet("")
		self.groupBox_2.setObjectName("groupBox_2")
		self.label_9 = QtWidgets.QLabel(self.groupBox_2)
		self.label_9.setGeometry(QtCore.QRect(80, 28, 121, 21))
		self.label_9.setObjectName("label_9")
		self.btnBrowse = QtWidgets.QToolButton(self.groupBox_2)
		self.btnBrowse.setGeometry(QtCore.QRect(40, 28, 41, 21))
		self.btnBrowse.setObjectName("btnBrowse")
		self.groupBox_3 = QtWidgets.QGroupBox(self.centralwidget)
		self.groupBox_3.setGeometry(QtCore.QRect(270, 110, 211, 111))
		self.groupBox_3.setObjectName("groupBox_3")
		self.label_12 = QtWidgets.QLabel(self.groupBox_3)
		self.label_12.setGeometry(QtCore.QRect(130, 20, 71, 21))
		self.label_12.setObjectName("label_12")
		self.label_13 = QtWidgets.QLabel(self.groupBox_3)
		self.label_13.setGeometry(QtCore.QRect(130, 40, 71, 21))
		self.label_13.setObjectName("label_13")
		self.label_14 = QtWidgets.QLabel(self.groupBox_3)
		self.label_14.setGeometry(QtCore.QRect(130, 60, 71, 21))
		self.label_14.setObjectName("label_14")
		self.label_15 = QtWidgets.QLabel(self.groupBox_3)
		self.label_15.setGeometry(QtCore.QRect(130, 80, 71, 21))
		self.label_15.setObjectName("label_15")
		self.label_16 = QtWidgets.QLabel(self.groupBox_3)
		self.label_16.setGeometry(QtCore.QRect(50, 60, 71, 21))
		self.label_16.setLayoutDirection(QtCore.Qt.RightToLeft)
		self.label_16.setStyleSheet("direction: rtl;")
		self.label_16.setScaledContents(False)
		self.label_16.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignTrailing | QtCore.Qt.AlignVCenter)
		self.label_16.setObjectName("label_16")
		self.label_17 = QtWidgets.QLabel(self.groupBox_3)
		self.label_17.setGeometry(QtCore.QRect(50, 80, 71, 21))
		self.label_17.setLayoutDirection(QtCore.Qt.RightToLeft)
		self.label_17.setStyleSheet("direction: rtl;")
		self.label_17.setScaledContents(False)
		self.label_17.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignTrailing | QtCore.Qt.AlignVCenter)
		self.label_17.setObjectName("label_17")
		self.label_18 = QtWidgets.QLabel(self.groupBox_3)
		self.label_18.setGeometry(QtCore.QRect(50, 20, 71, 21))
		self.label_18.setLayoutDirection(QtCore.Qt.RightToLeft)
		self.label_18.setStyleSheet("direction: rtl;")
		self.label_18.setScaledContents(False)
		self.label_18.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignTrailing | QtCore.Qt.AlignVCenter)
		self.label_18.setObjectName("label_18")
		self.label_19 = QtWidgets.QLabel(self.groupBox_3)
		self.label_19.setGeometry(QtCore.QRect(50, 40, 71, 21))
		self.label_19.setLayoutDirection(QtCore.Qt.RightToLeft)
		self.label_19.setStyleSheet("direction: rtl;")
		self.label_19.setScaledContents(False)
		self.label_19.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignTrailing | QtCore.Qt.AlignVCenter)
		self.label_19.setObjectName("label_19")
		self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
		self.pushButton_2.setGeometry(QtCore.QRect(340, 230, 141, 31))
		self.pushButton_2.setObjectName("pushButton_2")
		self.label_10 = QtWidgets.QLabel(self.centralwidget)
		self.label_10.setGeometry(QtCore.QRect(0, 10, 461, 21))
		font = QtGui.QFont()
		font.setFamily("IRANSans light")
		font.setPointSize(12)
		font.setBold(False)
		font.setWeight(50)


		font_elements = QtGui.QFont()
		font_elements.setFamily("tahoma")
		font_elements.setPointSize(8)
		font_elements.setBold(False)
		font_elements.setWeight(1)

		self.label_10.setFont(font)
		self.label_10.setObjectName("label_10")
		MainWindow.setCentralWidget(self.centralwidget)
		self.menubar = QtWidgets.QMenuBar(MainWindow)
		self.menubar.setGeometry(QtCore.QRect(0, 0, 501, 22))
		self.menubar.setObjectName("menubar")

		self.btnBrowse.clicked.connect(self.getfiles)
		self.pushButton.clicked.connect(self.test_connection)
		self.pushButton_2.clicked.connect(self.insert_into_mysqal_db)
		MainWindow.setMenuBar(self.menubar)
		# MainWindow.setFont(font_elements)

		self.retranslateUi(MainWindow)
		QtCore.QMetaObject.connectSlotsByName(MainWindow)

	def getfiles(self):
		try:
			fname = QFileDialog.getOpenFileName(None, "Select a file...", "C:\\", filter="All files (*)")
			self.plainTextEdit.setText("")
			explode = fname[0].split('/')
			self.pathFile = fname
			self.nameFile = explode[-1]
			self.fileSize = int(os.stat(fname[0]).st_size / 1024)
			self.data = pd.read_excel(r'{}'.format(fname[0]))
			excel_headers = self.data.columns
			excel_headers_list = []
			for i in excel_headers:
				excel_headers_list.append(i)
			self.list_excel_headers = ','.join(excel_headers_list)

			self.countTotalRows = len(self.data)
			self.columnstotal = len(self.data.axes[1])

			self.label_16.setText(str(self.columnstotal))
			self.label_19.setText(str(self.countTotalRows))
			self.label_18.setText(self.nameFile)
			self.label_17.setText(str(self.fileSize) + ' KB')
			self.plainTextEdit.insertPlainText('فایل با موفقیت باز شد' + '\n')
			self.plainTextEdit.setReadOnly(True)
			# self.plainTextEdit.insertPlainText('\n' + 'تعداد ستون ها: {}'.format(self.columnstotal))
			# self.plainTextEdit.insertPlainText('\n' + 'تعداد سطر ها: {}'.format(self.countTotalRows))
			self.plainTextEdit.insertPlainText('مسیر فایل: {}'.format(fname[0]) + '\n')
			# df_col_rowId = pd.DataFrame(self.data, columns=['id'])
			# df_col_name = pd.DataFrame(self.data, columns=['name'])
			# df_col_time = pd.DataFrame(self.data, columns=['time'])
		except:
			pass





	def insert_into_mysqal_db(self):
		try:
			ip = self.lineEdit.text()
			db = self.lineEdit_2.text()
			username = self.lineEdit_3.text()
			password = self.lineEdit_4.text()
			total_tables = 11
			table_name = ['employee_name','employee_post','national_id','birth_date','company_name','phone_number','date_create_card1','date_expire1','date_return_card','comment']

			self.plainTextEdit.insertPlainText('در حال اضافه کردن به دیتابیس' + '\n')

			temp_list_rows = []
			temp_dic_rows = {}
			error_code = 0
			for index, row in self.data.iterrows():
				for q in range(10):  # Adjusted the range to 10 since we removed employee_id
					temp_list_rows.append(row[table_name[q]])
					temp_dic_rows[table_name[q]] = row[table_name[q]]

				values_dict = temp_dic_rows.values()
				values_dict_list = list(values_dict)
				my_lst_str = ",".join(map(str, values_dict_list))

				try:
					mycursor = self.mydb.cursor()
					sql = "INSERT INTO edata (employee_name,employee_post,national_id,birth_date,company_name,phone_number,date_create_card1,date_expire1,date_return_card,comment) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
					val = values_dict_list
					print(val)
					mycursor.execute(sql, val)
					self.mydb.commit()
					print(mycursor.rowcount, "record inserted.")
				except:
					print('error insert db')
					error_code = 1
			if error_code == 0:
				self.plainTextEdit.insertPlainText(
					"تعداد {} سطر به دیتابیس اضافه شد".format(self.countTotalRows) + '\n')
			else:
				self.plainTextEdit.insertPlainText("خطا در افزدون دیتا" + '\n')
		except:
			print('fail in insert')
			self.plainTextEdit.insertPlainText("خطا: فایلی انتخاب نشده است" + '\n')

	def test_connection(self):
		ip = self.lineEdit.text()
		db = self.lineEdit_2.text()
		username = self.lineEdit_3.text()
		password = self.lineEdit_4.text()
		print('ip : {}'.format(ip))
		print('db : {}'.format(db))
		print('username : {}'.format(username))
		print('password : {}'.format(password))
		try:
			self.mydb = mysql.connector.connect(
				host=ip,
				user=username,
				password=password,
				database=db
			)
			if self.mydb.is_connected() == True:
				print('connected')
				self.label_8.setStyleSheet("color: #388E3C;font-weight: bold;")
				self.label_8.setText('متصل شد')
				self.plainTextEdit.insertPlainText('اتصال موفق به دیتابیس' + '\n')

		except:
			print('cant connect to db')
			self.label_8.setStyleSheet("color: #E53935;font-weight: bold;")
			self.label_8.setText('خطا در اتصال')
			self.plainTextEdit.insertPlainText('عدم اتصال به دیتابیس' + '\n')


	def retranslateUi(self, MainWindow):
		_translate = QtCore.QCoreApplication.translate
		MainWindow.setWindowTitle(_translate("MainWindow", self.appName))
		self.plainTextEdit.setPlainText(_translate("MainWindow", ""))
		self.groupBox.setTitle(_translate("MainWindow", "دیتابیس"))
		self.label.setText(_translate("MainWindow", "IP/host"))
		self.label_5.setText(_translate("MainWindow", "Database"))
		self.label_6.setText(_translate("MainWindow", "User"))
		self.label_7.setText(_translate("MainWindow", "Password"))
		self.pushButton.setText(_translate("MainWindow", "تست اتصال"))
		self.label_8.setText(_translate("MainWindow", ""))
		self.groupBox_2.setTitle(_translate("MainWindow", "انتخاب فایل"))
		self.label_9.setText(_translate("MainWindow", "انتخاب فایل اکسل :"))
		self.btnBrowse.setText(_translate("MainWindow", "..."))
		self.groupBox_3.setTitle(_translate("MainWindow", "اطلاعات فایل"))
		self.label_12.setText(_translate("MainWindow", "نام فایل:"))
		self.label_13.setText(_translate("MainWindow", "تعداد سطر:"))
		self.label_14.setText(_translate("MainWindow", "تعداد ستون:"))
		self.label_15.setText(_translate("MainWindow", "حجم فایل:"))
		self.label_16.setText(_translate("MainWindow", ""))
		self.label_17.setText(_translate("MainWindow", ""))
		self.label_18.setText(_translate("MainWindow", ""))
		self.label_19.setText(_translate("MainWindow", ""))
		self.lineEdit.setText(_translate("MainWindow", "localhost"))
		self.lineEdit_2.setText(_translate("MainWindow", "python"))
		self.lineEdit_3.setText(_translate("MainWindow", "root"))
		self.lineEdit_4.setText(_translate("MainWindow", ""))
		self.pushButton_2.setText(_translate("MainWindow", "اضافه کن"))
		self.label_10.setText(
			_translate("MainWindow", "افزودن اطلاعات از اکسل به دیتابیس"))


if __name__ == "__main__":
	import sys
	app = QtWidgets.QApplication(sys.argv)
	# app.setStyleSheet(qdarkgraystyle.load_stylesheet())
	# apply_stylesheet(app, theme='dark_cyan.xml')
	app.setStyle('Fusion')
	MainWindow = QtWidgets.QMainWindow()
	ui = Ui_MainWindow()
	ui.setupUi(MainWindow)
	MainWindow.show()
	sys.exit(app.exec_())

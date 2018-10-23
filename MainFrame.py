from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import openpyxl, cx_Oracle, pymysql, pymssql, Oracle_Tibero, MySQL, MSSQL, datetime, threading, collections, PIL

class MainFrame(Frame):
	def __init__(self, master):
		Frame.__init__(self, master)
		self.master = master
		self.master.title('Database Tool')
		self.pack(fill=BOTH, expand=True)
		self.information = {}
		self.DBinfo = {}
		self.connCheck = True

	# Select DBMS.
		frame1 = Frame(self)
		frame1.pack(fill=X)

		labelDBMS = Label(frame1, text='DBMS', width=10)
		labelDBMS.pack(side=LEFT, padx=10, pady=10)
		self.comboDBMS = ttk.Combobox(frame1, width=20)
		self.comboDBMS['values'] = ('Oracle / Tibero', 'Altibase', 'MS-SQL', 'MySQL / MariaDB')
		self.comboDBMS.current(0)
		self.comboDBMS.config(state='readonly')
		self.comboDBMS.pack(side=LEFT, pady=10)
		self.buttonConnect = ttk.Button(frame1, text='Connect', command=self.connectFunction)
		self.buttonConnect.pack(side=LEFT, padx=20)

	# Function button.
		frame6 = Frame(self)
		frame6.pack(fill=X, padx=10, pady=5)

		buttonED = ttk.Button(frame6, text='Excel -> DB Scheme', width=40, command=self.clickED)
		buttonES = ttk.Button(frame6, text='Excel -> SQL File', width=40, command=self.clickES)
		buttonDE = ttk.Button(frame6, text='DB Scheme -> Excel', width=40, command=self.clickDE)
		buttonDS = ttk.Button(frame6, text='DB Scheme -> SQL File', width=40, command=self.clickDS)

		buttonED.grid(row=0, column=0, padx=10, pady=10)
		buttonES.grid(row=0, column=1, padx=10, pady=10)
		buttonDE.grid(row=1, column=0, padx=10, pady=10)
		buttonDS.grid(row=1, column=1, padx=10, pady=10)

	# Set Progress bar.
		frame7 = Frame(self)
		frame7.pack(fill=X, padx=10)

		self.DBinfo['Progress'] = ttk.Progressbar(frame7, orient=HORIZONTAL, mode='determinate')
		self.DBinfo['Progress'].pack(fill=BOTH)

	# State of progress.
		frame8 = Frame(self)
		frame8.pack(fill=X, padx=10, pady=10)

		scrollbar = Scrollbar(frame8)
		scrollbar.pack(side=RIGHT, fill=Y)
		self.textB = Text(frame8)
		self.textB.pack(fill=BOTH, expand=1)
		self.textB.config(yscrollcommand=scrollbar.set)
		scrollbar.config(command=self.textB.yview)

# MainFrame excel file open.
	def openFunction(self):
		self.pathStr.set(filedialog.askopenfilename(initialdir="/", title='Select file', filetypes=(('excel files','*.xlsx'), ('sql files', '*.sql'), ('all files', '*.*'))))
		filePath = self.entryPath.get()
		if '.xlsx' in filePath:
			excelFile = openpyxl.load_workbook(filename=filePath)
			items = ['all'] + excelFile.sheetnames
			self.comboSheet['values'] = items
		else:
			self.comboSheet['values'] = []
		self.pathWindow.lift()

# DB Connection function.
	def connectFunction(self):
		self.textB.delete(1.0, END)
		if self.connCheck:
			self.connectionWindow = Toplevel()
			self.connectionWindow.title('DB Connection')
			self.connectionWindow.geometry('650x160+200+200')
			self.connectionWindow.resizable(False, False)

		# Connection info.
			frame_C1 = Frame(self.connectionWindow)
			frame_C1.pack(fill=X, padx=10, pady=10)

			labelAddr = Label(frame_C1, text='IP', width=5)
			labelAddr.pack(side=LEFT, padx=5, pady=10)
			self.entryAddr = ttk.Entry(frame_C1)
			self.entryAddr.pack(side=LEFT, expand=False)

			labelPort = Label(frame_C1, text='Port', width=5)
			labelPort.pack(side=LEFT, padx=5, pady=10)
			self.entryPort = ttk.Entry(frame_C1)
			self.entryPort.pack(side=LEFT, expand=False)

			if self.comboDBMS.get() == 'Oracle / Tibero':
				labelSid = Label(frame_C1, text='sid', width=5)
				labelSid.pack(side=LEFT, padx=5, pady=10)
			elif self.comboDBMS.get() == 'MySQL / MariaDB' or self.comboDBMS.get() == 'MS-SQL':
				labelDB = Label(frame_C1, text='DB', width=5)
				labelDB.pack(side=LEFT, padx=5, pady=10)
			self.entrySid = ttk.Entry(frame_C1)
			self.entrySid.pack(side=LEFT, expand=False)
		# User info.
			frameC2 = Frame(self.connectionWindow)
			frameC2.pack(fill=X, padx=10, pady=5)

			labelID = Label(frameC2, text='ID', width=5)
			labelID.pack(side=LEFT, padx=5, pady=10)
			self.entryID = ttk.Entry(frameC2)
			self.entryID.pack(side=LEFT, expand=False)

			labelPW = Label(frameC2, text='PW', width=5)
			labelPW.pack(side=LEFT, padx=5, pady=10)
			self.entryPW = ttk.Entry(frameC2)
			self.entryPW.pack(side=LEFT, expand=False)
			buttonTest = ttk.Button(frameC2, text='Connection Test', width=25, command=self.connectionTestFunction)
			buttonTest.pack(side=RIGHT, padx = 40)

		# Connect button.
			frameC3 = Frame(self.connectionWindow)
			frameC3.pack(fill=X, padx=10, pady=5)

			self.buttonConnectSave = ttk.Button(frameC3, text='Connect', width=21, command=self.connectionFunction)
			self.buttonConnectSave.pack(side=RIGHT, padx = 5)
			buttonAliasD = ttk.Button(frameC3, text='Alias Delete', width=20, command=self.aliasDeleteFunction)
			buttonAliasD.pack(side=RIGHT, padx = 5)
			buttonAliasR = ttk.Button(frameC3, text='Alias Registration', width=20, command=self.aliasFunction)
			buttonAliasR.pack(side=RIGHT, padx = 5)
			labelAlias = Label(frameC3, text='Alias', width=5)
			labelAlias.pack(side=LEFT, padx=5, pady=10)
			self.comboAlias = ttk.Combobox(frameC3, width=10)
			self.aliasRead()
			self.comboAlias['values'] = (['None'] + list(self.comboAliasValues.keys()))
			self.comboAlias.current(0)
			self.comboAlias.config(state='readonly')
			self.comboAlias.bind("<<ComboboxSelected>>", self.comboSelection)
			self.comboAlias.pack(side=LEFT, pady=10)

		else:
			self.db.close()
			self.connCheck = True
			f = open('C:\\Users\\Secuve\\Desktop\\Database Tool\\log\\log.txt', 'a')
			f.write(str('%s-%s-%s %s:%s:%s' %(datetime.datetime.now().year, datetime.datetime.now().month, datetime.datetime.now().day, datetime.datetime.now().hour, datetime.datetime.now().minute, datetime.datetime.now().second)) + '\tDisconnected.\t\t\t\t\t\t[ ' + self.information['IP'] + ', '+ str(self.information['Port']) + ', ' + self.information['sid'] + ', ' + self.information['ID'] + ', ' + self.information['PW'] + ' ]\n')
			f.close()
			for key in self.information.keys():
				self.information[key] = ''
			self.DBinfo['Progress'].stop()
			self.textB.delete(1.0, END)
			self.textB.insert(1.0, 'Database Disconnected!!')
			self.buttonConnect.configure(text='Connect')
			messagebox.showinfo('Info', 'Disconnected.')

# Alias read function.
	def aliasRead(self):
		f = open('C:\\Users\\Secuve\\Desktop\\Database Tool\\alias\\alias.txt', 'r')
		self.lines = f.readlines()
		f.close()
		self.comboAliasValues = collections.OrderedDict()
		for aliasList in self.lines:
			self.comboAliasValues[aliasList.split('*')[0]] = aliasList.split('*')[1].split('^')
		self.comboAlias['values'] = (['None'] + list(self.comboAliasValues.keys()))

# Combobox selection event.
	def comboSelection(self, event=None):
		self.entryAddr.delete(0, END)
		self.entryPort.delete(0, END)
		self.entrySid.delete(0, END)
		self.entryID.delete(0, END)
		self.entryPW.delete(0, END)
		if self.comboAlias.get() != 'None':
			self.entryAddr.insert(0, self.comboAliasValues[self.comboAlias.get()][0])
			self.entryPort.insert(0, self.comboAliasValues[self.comboAlias.get()][1])
			self.entrySid.insert(0, self.comboAliasValues[self.comboAlias.get()][2])
			self.entryID.insert(0, self.comboAliasValues[self.comboAlias.get()][3])
			self.entryPW.insert(0, self.comboAliasValues[self.comboAlias.get()][4].replace('\n', ''))

# Click events.
	def clickED(self):
		if self.connCheck:
			messagebox.showwarning('Warning', 'Please connect with DBMS.')
		else:
		# File path.
			self.pathWindow = Toplevel()
			self.pathWindow.title('Excel -> DB Scheme')
			self.pathWindow.geometry('650x100+200+200')
			self.pathWindow.resizable(False, False)

			frame5pathED = Frame(self.pathWindow)
			frame5pathED.pack(fill=X, pady=10)

			self.pathStr = StringVar()
			labelPathED = Label(frame5pathED, text='Path', width=5)
			labelPathED.pack(side=LEFT, padx=5)
			self.entryPath = ttk.Entry(frame5pathED, textvariable=self.pathStr)
			self.entryPath.pack(side=LEFT, fill=X, padx=5, expand=True)
			buttonPathED = ttk.Button(frame5pathED, text='open', command=self.openFunction)
			buttonPathED.pack(side=LEFT, padx=20)

		# Select sheet.
			frame5sheetED = Frame(self.pathWindow)
			frame5sheetED.pack(fill=X, padx=10)

			self.chk = IntVar()
			checkDrop = ttk.Checkbutton(frame5sheetED, text='DropTable', variable=self.chk)
			checkDrop.pack(side=LEFT, padx=10)
			buttonSheetED = ttk.Button(frame5sheetED, text='OK', command=self.clickED_S)
			buttonSheetED.pack(side=RIGHT, padx=10)
			self.comboSheet = ttk.Combobox(frame5sheetED, width=20)
			self.comboSheet['values'] = ('all')
			self.comboSheet.current(0)
			self.comboSheet.config(state='readonly')
			self.comboSheet.pack(side=RIGHT, padx=5, pady=10)
			labelSheetED = Label(frame5sheetED, text='Sheet', wid=5)
			labelSheetED.pack(side=RIGHT, padx=5)

	def clickED_S(self):
		self.textB.delete(1.0, END)
		self.DBinfo['Sheet'] = self.comboSheet.get()
		self.DBinfo['Path'] =  self.entryPath.get()
		self.DBinfo['Type'] = 'ED'
		self.DBinfo['Drop'] = self.chk.get()
		self.DBinfo['sid'] = self.information['sid']
		self.pathWindow.destroy()
		if self.chk.get() == 0:
			try:
				self.callThread()
			except IOError:
				self.DBinfo['Progress'].stop()
				messagebox.showwarning('Warning','Please select a Excel file.')
				self.pathWindow.lift()
			except cx_Oracle.DatabaseError as e:
				self.DBinfo['Progress'].stop()
				error, = e.args
				if error.code == 942:
					messagebox.showwarning('Warning','There is no Table to drop.')
					self.pathWindow.lift()
				elif error.code == 955:
					messagebox.showwarning('Warning', 'Please check the DB.\nThe table name is already used.')
					self.pathWindow.lift()
				elif error.code == 12569:
					messagebox.showwarning('Warning', 'Please check the DBMS.')
					self.pathWindow.lift()
			except pymysql.InternalError as e:
				self.DBinfo['Progress'].stop()
				code, message = e.args
				print (code, message)
				if code == 1050:
					messagebox.showwarning('Warning', 'Please check the DB.\nThe table name is already used.')
					self.pathWindow.lift()
		else:
			try:
				self.callThread()
			except cx_Oracle.DatabaseError as e:
				self.DBinfo['Progress'].stop()
				error, = e.args
				if error.code == 942:
					messagebox.showwarning('Warning','There is no Table to drop.')
					self.pathWindow.lift()

	def clickES(self):
	# File path.
		self.pathWindow = Toplevel()
		self.pathWindow.title('Excel -> SQL File')
		self.pathWindow.geometry('650x100+200+200')
		self.pathWindow.resizable(False, False)

		frame5pathES = Frame(self.pathWindow)
		frame5pathES.pack(fill=X, pady=10)

		self.pathStr = StringVar()
		labelPathES = Label(frame5pathES, text='Path', width=5)
		labelPathES.pack(side=LEFT, padx=5)
		self.entryPath = ttk.Entry(frame5pathES, textvariable=self.pathStr)
		self.entryPath.pack(side=LEFT, fill=X, padx=5, expand=True)
		buttonPathES = ttk.Button(frame5pathES, text='open', command=self.openFunction)
		buttonPathES.pack(side=LEFT, padx=20)

	# Select sheet.
		frame5sheetES = Frame(self.pathWindow)
		frame5sheetES.pack(fill=X, padx=10)

		buttonSheetES = ttk.Button(frame5sheetES, text='OK', command=self.clickES_S)
		buttonSheetES.pack(side=RIGHT, padx=10)
		self.comboSheet = ttk.Combobox(frame5sheetES, width=20)
		self.comboSheet['values'] = ('all')
		self.comboSheet.current(0)
		self.comboSheet.config(state='readonly')
		self.comboSheet.pack(side=RIGHT, padx=5, pady=10)
		labelSheetES = Label(frame5sheetES, text='Sheet', wid=5)
		labelSheetES.pack(side=RIGHT, padx=5)

	def clickES_S(self):
		try:
			self.textB.delete(1.0, END)
			self.DBinfo['Sheet'] = self.comboSheet.get()
			self.DBinfo['Path'] =  self.entryPath.get()
			self.DBinfo['Type'] = 'ES'
			self.saveWindow = Toplevel()
			self.saveWindow.title('SQL File')
			self.saveWindow.geometry('600x100+200+200')
			self.saveWindow.resizable(False, False)
		# File save path.
			frame_save = Frame(self.saveWindow)
			frame_save.pack(fill=X, padx=10, pady=10)

			lablePath_save = Label(frame_save, text='Path', width=5)
			lablePath_save.pack(side=LEFT, padx=5)
			self.entryPath_save = ttk.Entry(frame_save)
			self.entryPath_save.pack(side=LEFT, fill=X, padx=5, expand=True)
			buttonPath_save = ttk.Button(frame_save, text='path', command=self.pathESFunction)
			buttonPath_save.pack(side=LEFT, padx=5)
			buttonSave_save = ttk.Button(frame_save, text='save', command=self.callThread)
			buttonSave_save.pack(side=RIGHT, padx=5, pady=5)
			self.saveWindow.mainloop()

		except IOError:
			self.DBinfo['Progress'].stop()
			messagebox.showwarning('Warning','Please select a Excel file.')
			self.pathWindow.lift()

	def clickDE(self):
		if self.connCheck:
			messagebox.showwarning('Warning', 'Please connect with DBMS.')
		else:
			self.textB.delete(1.0, END)
			self.DBinfo['Type'] = 'DE'
			self.DBinfo['sid'] = self.information['sid']
			try:
				if self.comboDBMS.get() == 'Oracle / Tibero':
					self.DBinfo['Cursor'].execute('SELECT TABLE_NAME FROM tabs')
					tableList = self.DBinfo['Cursor'].fetchall()
					tableList.reverse()
				elif self.comboDBMS.get() == 'MySQL / MariaDB':
					self.DBinfo['Cursor'].execute('SHOW tables')
					tableList = list(self.DBinfo['Cursor'].fetchall())
				elif self.comboDBMS.get() == 'MS-SQL':
					self.DBinfo['Cursor'].execute('SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES')
					tableList = list(self.DBinfo['Cursor'].fetchall())
				self.DEWindow = Toplevel()
				self.DEWindow.title('DB Scheme -> Excel')
				self.DEWindow.geometry('600x200+200+200')
				self.DEWindow.resizable(False, False)

				frame_DE = Frame(self.DEWindow)
				frame_DE.pack(fill=X, padx=10, pady=10)
				DElablePath = Label(frame_DE, text='Path', width=5)
				DElablePath.pack(side=LEFT, padx=5)
				self.DEentryPath = ttk.Entry(frame_DE)
				self.DEentryPath.pack(side=LEFT, fill=X, padx=5, expand=True)
				DEbuttonPath = ttk.Button(frame_DE, text='path', command=self.pathDEFunction)
				DEbuttonPath.pack(side=LEFT, padx=5)

				frame_DE2 = Frame(self.DEWindow)
				frame_DE2.pack(fill=X, padx=10, pady=5)
				DEbuttonSave_save = ttk.Button(frame_DE2, text='save', command=self.callThread)
				DEbuttonSave_save.pack(side=RIGHT, padx=5, pady=5)
				self.listboxDE = Listbox(frame_DE2, width=30, selectmode=EXTENDED)
				self.listboxDE.pack(side=RIGHT, padx=5)
				self.listboxDE.delete(0, END)
				for item in tableList:
					self.listboxDE.insert(END, item)
				labelSheet = Label(frame_DE2, text='Sheet', width=5)
				labelSheet.pack(side=LEFT, padx=5, pady=5)
				self.entrySheet = ttk.Entry(frame_DE2, width=20)
				self.entrySheet.pack(side=LEFT, padx=5, pady=5)
				self.DEWindow.mainloop()
			except cx_Oracle.DatabaseError:
				self.DBinfo['Progress'].stop()
				messagebox.showwarning('Warning','Please check the DBConnection informations.')

	def clickDS(self):
		if self.connCheck:
			messagebox.showwarning('Warning', 'Please connect with DBMS.')
		else:
			self.textB.delete(1.0, END)
			self.DBinfo['Type'] = 'DS'
			self.DBinfo['sid'] = self.information['sid']
			try:
				if self.comboDBMS.get() == 'Oracle / Tibero':
					self.DBinfo['Cursor'].execute('SELECT TABLE_NAME FROM tabs')
					tableList = self.DBinfo['Cursor'].fetchall()
					tableList.reverse()
				elif self.comboDBMS.get() == 'MySQL / MariaDB':
					self.DBinfo['Cursor'].execute('SHOW tables')
					tableList = list(self.DBinfo['Cursor'].fetchall())
				elif self.comboDBMS.get() == 'MS-SQL':
					self.DBinfo['Cursor'].execute('SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES')
					tableList = list(self.DBinfo['Cursor'].fetchall())
				self.DSWindow = Toplevel()
				self.DSWindow.title('DB Scheme -> SQL File')
				self.DSWindow.geometry('600x200+200+200')
				self.DSWindow.resizable(False, False)

				frame_DS = Frame(self.DSWindow)
				frame_DS.pack(fill=X, padx=10, pady=10)
				DSlablePath = Label(frame_DS, text='Path', width=5)
				DSlablePath.pack(side=LEFT, padx=5)
				self.DSentryPath = ttk.Entry(frame_DS)
				self.DSentryPath.pack(side=LEFT, fill=X, padx=5, expand=True)
				DSbuttonPath = ttk.Button(frame_DS, text='path', command=self.pathDSFunction)
				DSbuttonPath.pack(side=LEFT, padx=5)

				frame_DS2 = Frame(self.DSWindow)
				frame_DS2.pack(fill=X, padx=10, pady=5)
				DSbuttonSave_save = ttk.Button(frame_DS2, text='save', command=self.callThread)
				DSbuttonSave_save.pack(side=RIGHT, padx=5, pady=5)
				self.listboxDS = Listbox(frame_DS2, width=50, selectmode=EXTENDED)
				self.listboxDS.pack(side=RIGHT, padx=5)
				self.listboxDS.delete(0, END)
				for item in tableList:
					self.listboxDS.insert(END, item)
				self.DSWindow.mainloop()
			except cx_Oracle.DatabaseError:
				self.DBinfo['Progress'].stop()
				messagebox.showwarning('Warning','Please check the DBConnection informations.')

	def callThread(self):
		if self.DBinfo['Type'] == 'ES':
			self.DBinfo['ESSave'] = self.entryPath_save.get()
			self.saveWindow.destroy()
			self.pathWindow.destroy()
		elif self.DBinfo['Type'] == 'DE':
			self.DBinfo['DEPath'] = self.DEentryPath.get()
			self.DBinfo['DESheet'] = self.entrySheet.get()
			self.DBinfo['DEListBox'] = self.listboxDE
			self.DBinfo['Window'] = self.DEWindow
			self.DEWindow.lower()
		elif self.DBinfo['Type'] == 'DS':
			self.DBinfo['DSPath'] = self.DSentryPath.get()
			self.DBinfo['DSListBox'] = self.listboxDS
			self.DBinfo['Window'] = self.DSWindow
			self.DSWindow.lower()
		th = threading.Thread(target=self.functionThread)
		th.start()

	def functionThread(self):
		try:
			if self.comboDBMS.get() == 'Oracle / Tibero':
				Oracle_Tibero.Oracle_Tibero(self.DBinfo, self.textB)
			elif self.comboDBMS.get() == 'MySQL / MariaDB':
				MySQL.MySQL(self.DBinfo, self.textB)
			elif self.comboDBMS.get() == 'MS-SQL':
				MSSQL.MSSQL(self.DBinfo, self.textB)
		except pymssql.OperationalError as e:
			self.DBinfo['Progress'].stop()
			code, message = e.args
			if code == 2714:
				messagebox.showwarning('Warning', 'Please check the DB.\nThe table name is already used.')

	def connectionTestFunction(self):
		try:
			testThread = threading.Thread(target=self.connectionTestThread)
			testThread.start()
		except cx_Oracle.DatabaseError as e:
			self.DBinfo['Progress'].stop()
			f = open('C:\\Users\\Secuve\\Desktop\\Database Tool\\log\\log.txt', 'a')
			f.write(str('%s-%s-%s %s:%s:%s' %(datetime.datetime.now().year, datetime.datetime.now().month, datetime.datetime.now().day, datetime.datetime.now().hour, datetime.datetime.now().minute, datetime.datetime.now().second)) + '\tConnection Test Failed.\t\t\t\t[ ' + self.entryAddr.get() + ', '+ self.entryPort.get() + ', ' + self.entrySid.get() + ', ' + self.entryID.get() + ', ' + self.entryPW.get() + ' ]\n')
			f.close()
			messagebox.showwarning('Warning', e)
			self.connectionWindow.lift()

	def connectionTestThread(self):
		if self.comboDBMS.get() == 'Oracle / Tibero':
			dsnTest = cx_Oracle.makedsn(self.entryAddr.get(), self.entryPort.get(), self.entrySid.get())
			test = cx_Oracle.connect(self.entryID.get(), self.entryPW.get(), dsnTest)
		elif self.comboDBMS.get() == 'MySQL / MariaDB':
			test = pymysql.connect(host=self.entryAddr.get(), port=int(self.entryPort.get()), user=self.entryID.get(), password=self.entryPW.get(), db=self.entrySid.get(), charset='utf8')
		elif self.comboDBMS.get() == 'MS-SQL':
			test = pymssql.connect(host=self.entryAddr.get(), port=int(self.entryPort.get()), user=self.entryID.get(), password=self.entryPW.get(), database=self.entrySid.get())
		test.close()
		messagebox.showinfo('info', 'Connection complete.')
		self.connectionWindow.lift()
		f = open('C:\\Users\\Secuve\\Desktop\\Database Tool\\log\\log.txt', 'a')
		f.write(str('%s-%s-%s %s:%s:%s' %(datetime.datetime.now().year, datetime.datetime.now().month, datetime.datetime.now().day, datetime.datetime.now().hour, datetime.datetime.now().minute, datetime.datetime.now().second)) + '\tConnection Tested.\t\t\t\t\t[ ' + self.entryAddr.get() + ', '+ self.entryPort.get() + ', ' + self.entrySid.get() + ', ' + self.entryID.get() + ', ' + self.entryPW.get() + ' ]\n')
		f.close()

	def connectionFunction(self):
			try:
				self.DBinfo['Progress'].start()
				if self.comboDBMS.get() == 'Oracle / Tibero':
					self.information = {'IP' : self.entryAddr.get(), 'Port' : self.entryPort.get(), 'sid' : self.entrySid.get(), 'ID' : self.entryID.get(), 'PW' : self.entryPW.get()}
					self.dsn = cx_Oracle.makedsn(self.information['IP'], self.information['Port'], self.information['sid'])
				elif self.comboDBMS.get() == 'MySQL / MariaDB':
					self.information = {'IP' : self.entryAddr.get(), 'Port' : int(self.entryPort.get()), 'sid' : self.entrySid.get(), 'ID' : self.entryID.get(), 'PW' : self.entryPW.get()}
				elif self.comboDBMS.get() == 'MS-SQL':
					self.information = {'IP' : self.entryAddr.get(), 'Port' : int(self.entryPort.get()), 'sid' : self.entrySid.get(), 'ID' : self.entryID.get(), 'PW' : self.entryPW.get()}
				self.connectionWindow.destroy()
				connThread = threading.Thread(target=self.ConnectThread)
				connThread.start()
			except IOError:
				self.DBinfo['Progress'].stop()
				messagebox.showwarning('Warning', 'Please fill out the all information.')
				self.connectionWindow.lift()
			except cx_Oracle.DatabaseError as e:
				self.DBinfo['Progress'].stop()
				f = open('C:\\Users\\Secuve\\Desktop\\Database Tool\\log\\log.txt', 'a')
				f.write(str('%s-%s-%s %s:%s:%s' %(datetime.datetime.now().year, datetime.datetime.now().month, datetime.datetime.now().day, datetime.datetime.now().hour, datetime.datetime.now().minute, datetime.datetime.now().second)) + '\tConnection Failed.\t\t\t\t\t[ ' + self.information['IP'] + ', '+ self.information['Port'] + ', ' + self.information['sid'] + ', ' + self.information['ID'] + ', ' + self.information['PW'] + ' ]\n')
				f.close()
				messagebox.showwarning('Warning', e)
				self.connectionWindow.lift()

	def ConnectThread(self):
		try:
			if self.comboDBMS.get() == 'Oracle / Tibero':
				self.db = cx_Oracle.connect(self.information['ID'], self.information['PW'], self.dsn)
			elif self.comboDBMS.get() == 'MySQL / MariaDB':
				self.db = pymysql.connect(host=self.information['IP'], port=self.information['Port'], user=self.information['ID'], password=self.information['PW'], db=self.information['sid'], charset='utf8')
			elif self.comboDBMS.get() == 'MS-SQL':
				self.db = pymssql.connect(host=self.information['IP'], port=self.information['Port'], user=self.information['ID'], password=self.information['PW'], database=self.information['sid'], charset='utf8')
			self.cursor = self.db.cursor()
			self.DBinfo['DB'] = self.db
			self.DBinfo['Cursor'] = self.cursor
			if self.connCheck:
				self.DBinfo['Progress'].stop()
				self.textB.delete(1.0, END)
				self.textB.insert(1.0, 'Database Connect!!')
				self.buttonConnect.configure(text='Disconnect')
				self.connCheck = False
				f = open('C:\\Users\\Secuve\\Desktop\\Database Tool\\log\\log.txt', 'a')
				if self.comboDBMS.get() == 'Oracle / Tibero':
					f.write(str('%s-%s-%s %s:%s:%s' %(datetime.datetime.now().year, datetime.datetime.now().month, datetime.datetime.now().day, datetime.datetime.now().hour, datetime.datetime.now().minute, datetime.datetime.now().second)) + '\tConnected.\t\t\t\t\t\t\t[ ' + self.information['IP'] + ', '+ self.information['Port'] + ', ' + self.information['sid'] + ', ' + self.information['ID'] + ', ' + self.information['PW'] + ' ]\n')
				elif self.comboDBMS.get() == 'MySQL / MariaDB':
					f.write(str('%s-%s-%s %s:%s:%s' %(datetime.datetime.now().year, datetime.datetime.now().month, datetime.datetime.now().day, datetime.datetime.now().hour, datetime.datetime.now().minute, datetime.datetime.now().second)) + '\tConnected.\t\t\t\t\t\t\t[ ' + self.information['IP'] + ', '+ str(self.information['Port']) + ', ' + self.information['sid'] + ', ' + self.information['ID'] + ', ' + self.information['PW'] + ' ]\n')
				f.close()
		except cx_Oracle.DatabaseError as e:
			self.DBinfo['Progress'].stop()
			error, = e.args
			if error.code == 12569:
				messagebox.showwarning('Warning', 'Please check the DBMS.')
		except pymssql.InterfaceError as e:
			self.DBinfo['Progress'].stop()
			messagebox.showwarning('Warning', e)

	def aliasFunction(self):
		self.aliasWindow = Toplevel()
		self.aliasWindow.title('Alias registration')
		self.aliasWindow.geometry('350x60+300+300')
		self.aliasWindow.resizable(False, False)

		frameC4 = Frame(self.aliasWindow)
		frameC4.pack(fill=X, padx=10, pady=5)

		labelAlias = Label(frameC4, text='Alias', width=5)
		labelAlias.pack(side=LEFT, padx=5, pady=10)
		self.entryAlias = ttk.Entry(frameC4)
		self.entryAlias.pack(side=LEFT, expand=False)
		buttonAliasR = ttk.Button(frameC4, text='Registration', width=21, command=self.registrationFunction)
		buttonAliasR.pack(side=LEFT, padx=20)

	def registrationFunction(self):
		try:
			self.information = {'IP' : self.entryAddr.get(), 'Port' : self.entryPort.get(), 'sid' : self.entrySid.get(), 'ID' : self.entryID.get(), 'PW' : self.entryPW.get()}
			for key, value in self.information.items():
				if value == '':
					raise IOError
			if self.entryAlias.get() == '' or self.entryAlias.get() == 'None':
				messagebox.showwarning('Warning', 'Please fill out the alias name.')
			else:
				if self.entryAlias.get() in self.comboAliasValues.keys():
					messagebox.showwarning('Warning', 'The alias name is already in use.')
				else:
					content = self.entryAlias.get() + '*' + self.information['IP'] + '^' + self.information['Port'] + '^' + self.information['sid'] + '^' + self.information['ID'] + '^' + self.information['PW'] + '*' + str(datetime.datetime.now()) + '\n'
					self.comboAlias['values'] = (['None'] + list(self.comboAliasValues.keys()))
					self.comboAlias.set(self.entryAlias.get())
					f = open('C:\\Users\\Secuve\\Desktop\\Database Tool\\alias\\alias.txt', 'a')
					f.write(content)
					f.close()
					self.aliasRead()
					f = open('C:\\Users\\Secuve\\Desktop\\Database Tool\\log\\log.txt', 'a')
					f.write(str('%s-%s-%s %s:%s:%s' %(datetime.datetime.now().year, datetime.datetime.now().month, datetime.datetime.now().day, datetime.datetime.now().hour, datetime.datetime.now().minute, datetime.datetime.now().second)) + '\tAlias Registered.\t\t\t\t\t[ ' + self.entryAlias.get() + ' ]\n')
					f.close()
			self.connectionWindow.lift()
			self.aliasWindow.destroy()
		except IOError:
			self.DBinfo['Progress'].stop()
			self.aliasWindow.destroy()
			messagebox.showwarning('Warning', 'Please fill out the all information.')
			self.connectionWindow.lift()

	def aliasDeleteFunction(self):
		if self.comboAlias.get() != 'None':
			content = ''
			if self.comboAlias.get() in self.comboAliasValues:
				del self.comboAliasValues[self.comboAlias.get()]
			KeyValue = self.comboAliasValues.items()
			for cnt in KeyValue:
				content += cnt[0] +  '*' + cnt[1][0] + '^' + cnt[1][1] + '^' + cnt[1][2] + '^' + cnt[1][3] + '^' + cnt[1][4] + '*' + str(datetime.datetime.now()) + '\n'
			f = open('C:\\Users\\Secuve\\Desktop\\Database Tool\\alias\\alias.txt', 'w')
			f.write(content)
			f.close()
			self.aliasRead()
			f = open('C:\\Users\\Secuve\\Desktop\\Database Tool\\log\\log.txt', 'a')
			f.write(str('%s-%s-%s %s:%s:%s' %(datetime.datetime.now().year, datetime.datetime.now().month, datetime.datetime.now().day, datetime.datetime.now().hour, datetime.datetime.now().minute, datetime.datetime.now().second)) + '\tAlias Deleted.\t\t\t\t\t\t[ ' + self.comboAlias.get() + ' ]\n')
			f.close()
			self.comboAlias['values'] = (['None'] + list(self.comboAliasValues.keys()))
			self.comboAlias.current(0)
			self.entryAddr.delete(0, END)
			self.entryPort.delete(0, END)
			self.entrySid.delete(0, END)
			self.entryID.delete(0, END)
			self.entryPW.delete(0, END)
		else:
			messagebox.showwarning('Warning', 'Please select a alias to delete.')
			self.connectionWindow.lift()

	def pathESFunction(self):
		self.entryPath_save.delete(0, END)
		self.entryPath_save.insert(0, filedialog.asksaveasfilename(defaultextension='.sql', initialdir='/',title='Select file', filetypes=(('sql files', '*.sql'), ('all files', '*.*'))))
		self.saveWindow.lift()

	def pathDEFunction(self):
		self.DEentryPath.delete(0, END)
		self.DEentryPath.insert(0, filedialog.asksaveasfilename(defaultextension='.xlsx', initialdir='/',title='Select file', filetypes=(('excel files','*.xlsx'), ('all files', '*.*'))))
		self.DEWindow.lift()

	def pathDSFunction(self):
		self.DSentryPath.delete(0, END)
		self.DSentryPath.insert(0, filedialog.asksaveasfilename(defaultextension='.sql', initialdir='/',title='Select file', filetypes=(('sql files', '*.sql'), ('all files', '*.*'))))
		self.DSWindow.lift()

def main():
# Create window.
	window = Tk()
	window.geometry('640x250+100+100')
	window.resizable(False, False)
	MainFrame(window)
	window.mainloop()

if __name__ == '__main__':
	main()
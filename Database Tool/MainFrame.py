from Tkinter import *
import tkFileDialog, ttk, openpyxl, Oracle_Tibero

class MainFrame(Frame) :
	def __init__(self, master) :
		Frame.__init__(self, master)
		self.master = master
		self.master.title('Database Tool')
		self.pack(fill = BOTH, expand = True)

	# Select DBMS.
		frame1 = Frame(self)
		frame1.pack(fill = X)

		labelDBMS = Label(frame1, text = 'DBMS', width = 10)
		labelDBMS.pack(side = LEFT, padx = 10, pady = 10)
		self.comboDBMS = ttk.Combobox(frame1, width = 20)
		self.comboDBMS['values'] = ('Oracle / Tibero', 'Altibase', 'MS-SQL', 'MySQL / MariaDB')
		self.comboDBMS.current(0)
		self.comboDBMS.pack(side = LEFT, pady = 10)

	# Connection info.
		frame2 = Frame(self)
		frame2.pack(fill = X)

		labelAddr = Label(frame2, text = 'IP', width = 5)
		labelAddr.pack(side = LEFT, padx = 10, pady = 10)
		self.entryAddr = ttk.Entry(frame2)
		self.entryAddr.pack(side = LEFT, expand = False)

		labelPort = Label(frame2, text = 'Port', width = 5)
		labelPort.pack(side = LEFT, padx = 10, pady = 10)
		self.entryPort = ttk.Entry(frame2)
		self.entryPort.pack(side = LEFT, expand = False)

		labelSid = Label(frame2, text = 'sid', width = 5)
		labelSid.pack(side = LEFT, padx = 10, pady = 10)
		self.entrySid = ttk.Entry(frame2)
		self.entrySid.pack(side = LEFT, expand = False)

	# User info.
		frame3 = Frame(self)
		frame3.pack(fill = X)

		labelID = Label(frame3, text = 'ID', width = 5)
		labelID.pack(side = LEFT, padx = 10, pady = 10)
		self.entryID = ttk.Entry(frame3)
		self.entryID.pack(side = LEFT, expand = False)

		labelPW = Label(frame3, text = 'PW', width = 5)
		labelPW.pack(side = LEFT, padx = 10, pady = 10)
		self.entryPW = ttk.Entry(frame3)
		self.entryPW.pack(side = LEFT, expand = False)

	# File path.
		frame4 = Frame(self)
		frame4.pack(fill = X, pady = 10)

		self.pathStr = StringVar()

		labelPath = Label(frame4, text = 'Path', width = 5)
		labelPath.pack(side = LEFT, padx = 5)
		self.entryPath = ttk.Entry(frame4, textvariable = self.pathStr)
		self.entryPath.pack(side = LEFT, fill = X, padx = 5, expand = True)
		buttonPath = ttk.Button(frame4, text = 'open', command = self.openFunction)
		buttonPath.pack(side = LEFT, padx = 20)
	# Select sheet.
		frame5 = Frame(self)
		frame5.pack(fill = X, padx = 60)

		self.comboSheet = ttk.Combobox(frame5, width = 20)
		self.comboSheet['values'] = ('all')
		self.comboSheet.current(0)
		self.comboSheet.pack(side = RIGHT, pady = 10)
		labelSheet = Label(frame5, text = 'Sheet', wid = 5)
		labelSheet.pack(side = RIGHT, padx = 5)

	# Function button.
		frame6 = Frame(self)
		frame6.pack(fill = X, padx = 10, pady = 10)

		buttonED = ttk.Button(frame6, text = 'Excel -> DB Scheme', width = 40, command = self.clickED)
		buttonES = ttk.Button(frame6, text = 'Excel -> SQL File', width = 40, command = self.clickES)
		buttonDE = ttk.Button(frame6, text = 'DB Scheme -> Excel', width = 40, command = self.clickDE)
		buttonDS = ttk.Button(frame6, text = 'DB Scheme -> SQL File', width = 40, command = self.clickDS)

		buttonED.grid(row = 0, column = 0, padx = 10, pady = 10)
		buttonES.grid(row = 0, column = 1, padx = 10, pady = 10)
		buttonDE.grid(row = 1, column = 0, padx = 10, pady = 10)
		buttonDS.grid(row = 1, column = 1, padx = 10, pady = 10)

	# State of progress.
		frame7 = Frame(self)
		frame7.pack(fill = X, padx = 10, pady = 10)

		scrollbar = Scrollbar(frame7)
		scrollbar.pack(side = RIGHT, fill = Y)
		self.textB = Text(frame7)
		self.textB.pack(fill = BOTH, expand = 1)
		self.textB.config(yscrollcommand = scrollbar.set)
		scrollbar.config(command = self.textB.yview)

# MainFrame excel file open.
	def openFunction(self) :
		self.pathStr.set(tkFileDialog.askopenfilename(initialdir = "/", title = "Select file", filetypes = (("excel files","*.xlsx"), ("sql files", "*.sql"), ("all files", "*.*"))))
		filePath = self.entryPath.get()
		if '.xlsx' in filePath :
			excelFile = openpyxl.load_workbook(filename = filePath)
			items = ['all'] + excelFile.sheetnames
			self.comboSheet['values'] = items
		else :
			self.comboSheet['values'] = []

# Click events.
	def clickED(self) :
		parameter = {'IP' : self.entryAddr.get(), 'Port' : self.entryPort.get(), 'sid' : self.entrySid.get(), 'ID' : self.entryID.get(), 'PW' : self.entryPW.get(), 'Path' : self.entryPath.get(), 'Sheet' : self.comboSheet.get()}
		if self.comboDBMS.get() == 'Oracle / Tibero' :
			parameter['Type'] = 'ED'
			Oracle_Tibero.Oracle_Tibero(parameter, self.textB)

	def clickES(self) :
		parameter = {'IP' : self.entryAddr.get(), 'Port' : self.entryPort.get(), 'sid' : self.entrySid.get(), 'ID' : self.entryID.get(), 'PW' : self.entryPW.get(), 'Path' : self.entryPath.get(), 'Sheet' : self.comboSheet.get()}
		if self.comboDBMS.get() == 'Oracle / Tibero' :
			parameter['Type'] = 'ES'
			Oracle_Tibero.Oracle_Tibero(parameter, self.textB)

	def clickDE(self) :
		parameter = {'IP' : self.entryAddr.get(), 'Port' : self.entryPort.get(), 'sid' : self.entrySid.get(), 'ID' : self.entryID.get(), 'PW' : self.entryPW.get(), 'Path' : self.entryPath.get(), 'Sheet' : self.comboSheet.get()}
		if self.comboDBMS.get() == 'Oracle / Tibero' :
			parameter['Type'] = 'DE'
			Oracle_Tibero.Oracle_Tibero(parameter, self.textB)

	def clickDS(self) :
		parameter = {'IP' : self.entryAddr.get(), 'Port' : self.entryPort.get(), 'sid' : self.entrySid.get(), 'ID' : self.entryID.get(), 'PW' : self.entryPW.get(), 'Path' : self.entryPath.get(), 'Sheet' : self.comboSheet.get()}
		if self.comboDBMS.get() == 'Oracle / Tibero' :
			parameter['Type'] = 'DS'
			Oracle_Tibero.Oracle_Tibero(parameter, self.textB)
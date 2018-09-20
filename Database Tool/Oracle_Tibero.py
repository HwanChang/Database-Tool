import tkFileDialog, ttk, openpyxl, cx_Oracle, tkMessageBox
from Tkinter import *
from openpyxl.styles import Font

class Oracle_Tibero :
	def __init__(self, info, textB) :
		self.info = info
		self.textB = textB

	# Functions by button type.
		try :
			if self.info['Type'] == 'ED' or self.info['Type'] == 'ES' :
				self.sendSQL = {}
				self.excel_document = openpyxl.load_workbook(self.info['Path'])
				self.sheetList = self.excel_document.sheetnames
				if self.info['Sheet'] == 'all' :
					for sheetnameList in self.sheetList :
						self.sendSQL[str(sheetnameList)] = ''
						self.ED_ESFunction(str(sheetnameList))
				else :
					self.ED_ESFunction(self.info['Sheet'])
			elif self.info['Type'] == 'DE' :
				self.DEmkFunction()
			elif self.info['Type'] == 'DS' :
				self.DSmkFunction()
		except IOError :
				tkMessageBox.showwarning('Warning',"Please select a Excel file.")
		except cx_Oracle.DatabaseError as e:
		    error, = e.args
		    if error.code == 12533:
		        tkMessageBox.showwarning('Warning',"Please fill out the DBConnection informations.")
		        # sys.exit()?
		    elif error.code == 955:
		        tkMessageBox.showwarning('Warning',"Please check the DB.\nThe table name is already used.")

# Main function about excel -> DB Scheme, Excel -> SQL File.
	def ED_ESFunction(self, name) :
		self.name = name
		sheet = self.excel_document[name]

		cl = []
		ty = []
		le = []
		nu = []
		ky = []
		SQL = ''
		PKEY = ''
		k = False

		all_rows = sheet.rows
		for row in all_rows :
			if row[1].value is not None :
				if isinstance(row[2].value, float) :
					le.append(str(int(row[2].value)))
				else :
					le.append(str(row[2].value))
				cl.append(row[0].value)
				ty.append(row[1].value)
				nu.append(row[3].value)
				ky.append(row[4].value)

		tableName = sheet['A1'].value
		for i in range(1, len(ty)) :
			if i == len(ty)-1 :
				if nu[i] == 'N' :
					SQL += cl[i] + ' ' + ty[i] + '(' + le[i] + ')' + ' NOT NULL'
				else :
					SQL += cl[i] + ' ' + ty[i] + '(' + le[i] + ')'

			else :
				if nu[i] == 'N' :
					SQL += cl[i] + ' ' + ty[i] + '(' + le[i] + ')' + ' NOT NULL, '
				else :
					SQL += cl[i] + ' ' + ty[i] + '(' + le[i] + ')' + ', '

			if ky[i] == 'PK' :
				k = True
				PKEY += ', CONSTRAINT ' + str(tableName) + 'pk PRIMARY KEY (' + cl[i] + ')'
		if k == True :
			self.sendSQL[name] = 'CREATE TABLE ' + str(tableName) + ' ( ' + SQL + PKEY + ' )'
		else :
			self.sendSQL[name] = 'CREATE TABLE ' + str(tableName) + ' ( ' + SQL + ' )'

	# Excel -> DB Scheme function
		if self.info['Type'] == 'ED' :
			self.textB.delete(1.0, END)
			dsnED = cx_Oracle.makedsn(self.info['IP'], self.info['Port'], self.info['sid'])
			dbED = cx_Oracle.connect(self.info['ID'], self.info['PW'], dsnED)
			cursorED = dbED.cursor()
			cursorED.execute(self.sendSQL[name])
			cursorED.close()
			dbED.close()
			self.textB.insert(1.0, 'Excel -> DB Scheme Complete!')

		elif self.info['Type'] == 'ES' :
			if self.info['Sheet'] == 'all' :
				if name == self.sheetList[len(self.sheetList)-1] :
					self.ESmkFunction()
			else :
				self.ESmkFunction()

# Excel -> SQL File function
	def ESmkFunction(self) :
		self.saveWindow = Toplevel()
		self.saveWindow.title('SQL File')
		self.saveWindow.geometry('600x100+200+200')
		self.saveWindow.resizable(False, False)

	# File save path.
		frame_save = Frame(self.saveWindow)
		frame_save.pack(fill = X, padx = 10, pady = 10)

		lablePath_save = Label(frame_save, text = 'Path', width = 5)
		lablePath_save.pack(side = LEFT, padx = 5)
		self.entryPath_save = ttk.Entry(frame_save)
		self.entryPath_save.pack(side = LEFT, fill = X, padx = 5, expand = True)
		buttonPath_save = ttk.Button(frame_save, text = 'path', command = self.pathESFunction)
		buttonPath_save.pack(side = LEFT, padx = 5)
		buttonSave_save = ttk.Button(frame_save, text = 'save', command = self.Excel_SQLFunction)
		buttonSave_save.pack(side = RIGHT, padx = 5, pady = 5)
		self.saveWindow.mainloop()

	def DEmkFunction(self) :
		try :
			self.dsnDE = cx_Oracle.makedsn(self.info['IP'], self.info['Port'], self.info['sid'])
			self.dbDE = cx_Oracle.connect(self.info['ID'], self.info['PW'], self.dsnDE)
			self.cursorDE = self.dbDE.cursor()
			self.cursorDE.execute('SELECT TABLE_NAME FROM tabs')
			tableList = self.cursorDE.fetchall()

			self.DEWindow = Toplevel()
			self.DEWindow.title('Excel File')
			self.DEWindow.geometry('600x100+200+200')
			self.DEWindow.resizable(False, False)

			frame_DE = Frame(self.DEWindow)
			frame_DE.pack(fill = X, padx = 10, pady = 10)
			DElablePath = Label(frame_DE, text = 'Path', width = 5)
			DElablePath.pack(side = LEFT, padx = 5)
			self.DEentryPath = ttk.Entry(frame_DE)
			self.DEentryPath.pack(side = LEFT, fill = X, padx = 5, expand = True)
			DEbuttonPath = ttk.Button(frame_DE, text = 'path', command = self.pathDEFunction)
			DEbuttonPath.pack(side = LEFT, padx = 5)

			frame_DE2 = Frame(self.DEWindow)
			frame_DE2.pack(fill = X, padx = 10, pady = 5)
			DEbuttonSave_save = ttk.Button(frame_DE2, text = 'save', command = self.DB_ExcelFunction)
			DEbuttonSave_save.pack(side = RIGHT, padx = 5, pady = 5)
			tbl = []
			for table in tableList :
				tbl = tbl + [table[0]]
			self.comboTbl = ttk.Combobox(frame_DE2, width = 20)
			self.comboTbl['values'] = tbl
			self.comboTbl.current(0)
			self.comboTbl.pack(side = RIGHT, padx = 5, pady = 5)
			labelSheet = Label(frame_DE2, text = 'Sheet', width = 5)
			labelSheet.pack(side = LEFT, padx = 5, pady = 5)
			self.entrySheet = ttk.Entry(frame_DE2, width = 20)
			self.entrySheet.pack(side = LEFT, padx = 5, pady = 5)
			self.DEWindow.mainloop()
		except cx_Oracle.DatabaseError :
					tkMessageBox.showwarning('Warning',"Please fill out the DBConnection informations.")

	def DSmkFunction(self) :
		try :
			self.dsnDS = cx_Oracle.makedsn(self.info['IP'], self.info['Port'], self.info['sid'])
			self.dbDS = cx_Oracle.connect(self.info['ID'], self.info['PW'], self.dsnDS)
			self.cursorDS = self.dbDS.cursor()
			self.cursorDS.execute('SELECT TABLE_NAME FROM tabs')
			tableList = self.cursorDS.fetchall()

			self.DSWindow = Toplevel()
			self.DSWindow.title('SQL File')
			self.DSWindow.geometry('600x100+200+200')
			self.DSWindow.resizable(False, False)

			frame_DS = Frame(self.DSWindow)
			frame_DS.pack(fill = X, padx = 10, pady = 10)
			DSlablePath = Label(frame_DS, text = 'Path', width = 5)
			DSlablePath.pack(side = LEFT, padx = 5)
			self.DSentryPath = ttk.Entry(frame_DS)
			self.DSentryPath.pack(side = LEFT, fill = X, padx = 5, expand = True)
			DSbuttonPath = ttk.Button(frame_DS, text = 'path', command = self.pathDSFunction)
			DSbuttonPath.pack(side = LEFT, padx = 5)

			frame_DS2 = Frame(self.DSWindow)
			frame_DS2.pack(fill = X, padx = 10, pady = 5)
			DSbuttonSave_save = ttk.Button(frame_DS2, text = 'save', command = self.DB_SQLFunction)
			DSbuttonSave_save.pack(side = RIGHT, padx = 5, pady = 5)
			tbl = []
			for table in tableList :
				tbl = tbl + [table[0]]
			self.comboTblDS = ttk.Combobox(frame_DS2, width = 20)
			self.comboTblDS['values'] = tbl
			self.comboTblDS.current(0)
			self.comboTblDS.pack(side = RIGHT, padx = 5, pady = 5)

			self.DSWindow.mainloop()
		except cx_Oracle.DatabaseError :
			tkMessageBox.showwarning('Warning',"Please fill out the DBConnection informations.")

	def pathESFunction(self) :
		self.entryPath_save.delete(0, END)
		self.entryPath_save.insert(0, tkFileDialog.asksaveasfilename(defaultextension='.sql', initialdir = "/",title = "Select file", filetypes = (("sql files", "*.sql"), ("all files", "*.*"))))
		self.saveWindow.lift()

	def pathDEFunction(self) :
		self.DEentryPath.delete(0, END)
		self.DEentryPath.insert(0, tkFileDialog.asksaveasfilename(defaultextension='.xlsx', initialdir = "/",title = "Select file", filetypes = (("excel files","*.xlsx"), ("all files", "*.*"))))
		self.DEWindow.lift()

	def pathDSFunction(self) :
		self.DSentryPath.delete(0, END)
		self.DSentryPath.insert(0, tkFileDialog.asksaveasfilename(defaultextension='.sql', initialdir = "/",title = "Select file", filetypes = (("sql files", "*.sql"), ("all files", "*.*"))))
		self.DSWindow.lift()

	def Excel_SQLFunction(self) :
		try :
			self.textB.delete(1.0, END)
			f = open(self.entryPath_save.get(), 'w')
			if self.info['Sheet'] == 'all' :
				for sheetnamesList in self.sheetList :
					if sheetnamesList == self.sheetList[len(self.sheetList)-1] :
						f.write(self.sendSQL[str(sheetnamesList)])
						break
					f.write(self.sendSQL[str(sheetnamesList)] + '\n\n')
			else :
				f.write(self.sendSQL[self.name])
			f.close()
			self.textB.insert(1.0, 'Excel -> SQL File Complete!\n\n')
			self.textB.insert(END, self.entryPath_save.get())
			self.saveWindow.destroy()
		except IOError :
			tkMessageBox.showwarning('Warning',"Please select a SQL file to save.")
			self.saveWindow.lift()

	def DB_ExcelFunction(self) :
		try :
			self.textB.delete(1.0, END)
			self.cursorDE.execute("SELECT COLUMN_NAME, DATA_TYPE, DATA_LENGTH FROM USER_TAB_COLUMNS WHERE TABLE_NAME = '" + self.comboTbl.get() +"'")
			saveList = self.cursorDE.fetchall()
			self.cursorDE.execute("SELECT COLUMN_NAME FROM USER_CONS_COLUMNS WHERE CONSTRAINT_NAME = '" + self.comboTbl.get() + "PK'")
			pkList = self.cursorDE.fetchall()
			self.cursorDE.execute("SELECT SEARCH_CONDITION FROM ALL_CONSTRAINTS WHERE TABLE_NAME = '" + self.comboTbl.get() + "'")
			nullList = self.cursorDE.fetchall()

			wb = openpyxl.Workbook()
			sheetmkNew = wb.active
			sheetmkNew.title = self.entrySheet.get()

			sheetmkNew.column_dimensions['A'].width = 25
			sheetmkNew.column_dimensions['B'].width = 15
			sheetmkNew.column_dimensions['C'].width = 15
			sheetmkNew.column_dimensions['D'].width = 10
			sheetmkNew.column_dimensions['E'].width = 10

			sheetmkNew.freeze_panes = 'A4'

			fontObj = Font(size = 20, bold = True)
			fontBold = Font(bold = True)
			sheetmkNew['A1'].font = fontObj
			sheetmkNew['A3'].font = fontBold
			sheetmkNew['B3'].font = fontBold
			sheetmkNew['C3'].font = fontBold
			sheetmkNew['D3'].font = fontBold
			sheetmkNew['E3'].font = fontBold

			cnt = 4
			sheetmkNew.cell(row = 1, column = 1).value = self.comboTbl.get()
			sheetmkNew.cell(row = 3, column = 1).value = 'COLUMN_NAME'
			sheetmkNew.cell(row = 3, column = 2).value = 'DATA_TYPE'
			sheetmkNew.cell(row = 3, column = 3).value = 'DATA_LENGTH'
			sheetmkNew.cell(row = 3, column = 4).value = 'NULL'
			sheetmkNew.cell(row = 3, column = 5).value = 'KEY'

			for into in saveList :
				sheetmkNew.cell(row = cnt, column = 1).value = into[0]
				sheetmkNew.cell(row = cnt, column = 2).value = into[1]
				sheetmkNew.cell(row = cnt, column = 3).value = into[2]
				cnt += 1

			nullName = []
			for nulllist in nullList :
				for i in range(1, len(str(nulllist[0]))) :
					if str(nulllist[0])[i] == '"' :
						nullName.append(str(nulllist[0])[1:i])

			cnt = 4
			for count in range(0, len(saveList)) :
				for n in range(0, len(nullName)) :
					if nullName[n] in saveList[count] :
						sheetmkNew.cell(row = cnt, column = 4).value = 'N'
				cnt += 1

			cnt = 4
			for count in range(0, len(saveList)) :
				for k in range(0, len(pkList)) :
					if pkList[k][0] in saveList[count] :
						sheetmkNew.cell(row = cnt, column = 5).value = 'PK'
				cnt += 1

			wb.save(self.DEentryPath.get())
			self.textB.insert(1.0, 'DB Scheme -> Excel Complete!\n\n')
			self.textB.insert(END, self.DEentryPath.get())
			self.cursorDE.close()
			self.dbDE.close()
			self.DEWindow.destroy()
		except IOError :
			tkMessageBox.showwarning('Warning',"Please select a Excel file to save.")
			self.DEWindow.lift()
		except ValueError :
			tkMessageBox.showwarning('Warning',"Please fill out the Sheet name.")
			self.DEWindow.lift()

	def DB_SQLFunction(self) :
		try :
			self.textB.delete(1.0, END)
			self.cursorDS.execute("SELECT COLUMN_NAME, DATA_TYPE, DATA_LENGTH FROM USER_TAB_COLUMNS WHERE TABLE_NAME = '" + self.comboTblDS.get() +"'")
			sqlList = self.cursorDS.fetchall()
			self.cursorDS.execute("SELECT COLUMN_NAME FROM USER_CONS_COLUMNS WHERE CONSTRAINT_NAME = '" + self.comboTblDS.get() + "PK'")
			pkList = self.cursorDS.fetchall()
			self.cursorDS.execute("SELECT SEARCH_CONDITION FROM ALL_CONSTRAINTS WHERE TABLE_NAME = '" + self.comboTblDS.get() + "'")
			nullList = self.cursorDS.fetchall()

			SQLsentence = 'CREATE TABLE ' + self.comboTblDS.get() + ' ( '
			SQLsentenceList = []
			nllist = []
			for sqllist in sqlList :
				SQLsentenceList = SQLsentenceList + [sqllist[0] + ' ' + sqllist[1] + '(' + str(int(sqllist[2])) + ')']
			for nulllist in nullList :
				for i in range(1, len(str(nulllist[0]))) :
					if str(nulllist[0])[i] == '"' :
						nllist.append(str(nulllist[0])[1:i])
			for count in range(0, len(SQLsentenceList)) :
				for j in range(0, len(nllist)) :
					if nllist[j] in SQLsentenceList[count] :
						SQLsentenceList[count] = SQLsentenceList[count] + ' ' + 'NOT NULL'
				if count == len(SQLsentenceList)-1 :
					break
				SQLsentenceList[count] = SQLsentenceList[count] + ','

				for k in range(0, len(pkList)) :
					if pkList[k][0] in SQLsentenceList[count] :
						SQLsentenceList.append('CONSTRAINT ' + self.comboTblDS.get() + 'pk PRIMARY KEY (' + pkList[k][0] + ')')

			for makeSentence in range(0, len(SQLsentenceList)) :
				SQLsentence = SQLsentence + SQLsentenceList[makeSentence] + ' '
			SQLsentence = SQLsentence + ')'

			f = open(self.DSentryPath.get(), 'w')
			f.write(SQLsentence)
			f.close()
			self.textB.insert(1.0, 'DB Scheme -> SQL File Complete!\n\n')
			self.textB.insert(END, self.DSentryPath.get())
			self.cursorDS.close()
			self.dbDS.close()
			self.DSWindow.destroy()
		except IOError :
			tkMessageBox.showwarning('Warning',"Please select a SQL file to save.")
			self.DSWindow.lift()
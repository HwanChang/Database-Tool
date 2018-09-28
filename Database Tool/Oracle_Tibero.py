import tkFileDialog, ttk, openpyxl, cx_Oracle, tkMessageBox, datetime
from Tkinter import *
from openpyxl.styles import Font

class Oracle_Tibero:
	def __init__(self, info, textB):
		self.info = info
		self.textB = textB

	# Functions by button type.
		if self.info['Type'] == 'ED' or self.info['Type'] == 'ES':
			self.sendSQL = {}
			self.excel_document = openpyxl.load_workbook(self.info['Path'])
			self.sheetList = self.excel_document.sheetnames
			if self.info['Sheet'] == 'all':
				for sheetnameList in self.sheetList:
					self.sendSQL[str(sheetnameList)] = ''
					self.ED_ESFunction(str(sheetnameList))
			else:
				self.ED_ESFunction(self.info['Sheet'])
		elif self.info['Type'] == 'DE':
			self.DEmkFunction()
		elif self.info['Type'] == 'DS':
			self.DSmkFunction()

# Main function about excel -> DB Scheme, Excel -> SQL File.
	def ED_ESFunction(self, name):
		self.name = name
		self.sheet = self.excel_document[name]
		check_tableName = []
		self.sendSQL[name] = []
		for check_rows in self.sheet.rows:
			if check_rows[0].value is not None:
				cnt = list(self.sheet.rows).index(check_rows)
				cl = []
				ty = []
				le = []
				nu = []
				ky = []
				mkList = []
				SQL = ''
				PKEY = ''
				k = False

				all_rows = self.sheet.rows
				for row in list(all_rows)[cnt+1:]:
					if row[2].value is not None:
						if isinstance(row[3].value, float):
							le.append(str(int(row[3].value)))
						else:
							le.append(str(row[3].value))
						cl.append(row[1].value)
						ty.append(row[2].value)
						nu.append(row[4].value)
						ky.append(row[5].value)
					else:
						break

				tableName = self.sheet['A' + str(cnt+1)].value
				for i in range(1, len(cl)):
					if i == len(cl)-1:
						if nu[i] == 'N':
							SQL += cl[i] + ' ' + ty[i] + '(' + le[i] + ')' + ' NOT NULL'
						else:
							SQL += cl[i] + ' ' + ty[i] + '(' + le[i] + ')'

					else:
						if nu[i] == 'N':
							SQL += cl[i] + ' ' + ty[i] + '(' + le[i] + ')' + ' NOT NULL, '
						else:
							SQL += cl[i] + ' ' + ty[i] + '(' + le[i] + ')' + ', '

					if ky[i] == 'PK':
						k = True
						PKEY += ', CONSTRAINT ' + str(tableName) + 'pk PRIMARY KEY (' + cl[i] + ')'

				if k == True:
					self.sendSQL[name].append('CREATE TABLE ' + str(tableName) + ' ( ' + SQL + PKEY + ' )')
				else:
					self.sendSQL[name].append('CREATE TABLE ' + str(tableName) + ' ( ' + SQL + PKEY + ' )')
			# Excel -> DB Scheme function
		if self.info['Type'] == 'ED':
			self.textB.delete(1.0, END)
			for send in self.sendSQL[name]:
				self.info['Cursor'].execute(send)
			self.textB.insert(1.0, 'Excel -> DB Scheme Complete!')
			self.info['Window'].destroy()
			f = open('C:\\Users\\Secuve\\Desktop\\Database Tool\\log\\log.txt', 'a')
			f.write(str('%s-%s-%s %s:%s:%s' %(datetime.datetime.now().year, datetime.datetime.now().month, datetime.datetime.now().day, datetime.datetime.now().hour, datetime.datetime.now().minute, datetime.datetime.now().second)) + '\tExcel -> DB Scheme Function.\t\t[ ' + str(self.info['Path'].encode('euc-kr')) + ' ]' + '\n')
			f.close()

		elif self.info['Type'] == 'ES':
			if self.info['Sheet'] == 'all':
				if name == self.sheetList[len(self.sheetList)-1]:
					self.ESmkFunction()
			else:
				self.ESmkFunction()

# Excel -> SQL File function
	def ESmkFunction(self):
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
		buttonSave_save = ttk.Button(frame_save, text='save', command=self.Excel_SQLFunction)
		buttonSave_save.pack(side=RIGHT, padx=5, pady=5)
		self.saveWindow.mainloop()

	def DEmkFunction(self):
		try:
			self.info['Cursor'].execute('SELECT TABLE_NAME FROM tabs')
			tableList = self.info['Cursor'].fetchall()

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
			DEbuttonSave_save = ttk.Button(frame_DE2, text='save', command=self.DB_ExcelFunction)
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
					tkMessageBox.showwarning('Warning','Please check the DBConnection informations.')

	def DSmkFunction(self):
		try:
			self.info['Cursor'].execute('SELECT TABLE_NAME FROM tabs')
			tableList = self.info['Cursor'].fetchall()

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
			DSbuttonSave_save = ttk.Button(frame_DS2, text='save', command=self.DB_SQLFunction)
			DSbuttonSave_save.pack(side=RIGHT, padx=5, pady=5)
			self.listboxDS = Listbox(frame_DS2, width=50, selectmode=EXTENDED)
			self.listboxDS.pack(side=RIGHT, padx=5)
			self.listboxDS.delete(0, END)
			for item in tableList:
				self.listboxDS.insert(END, item)

			self.DSWindow.mainloop()
		except cx_Oracle.DatabaseError:
			tkMessageBox.showwarning('Warning','Please check the DBConnection informations.')

	def pathESFunction(self):
		self.entryPath_save.delete(0, END)
		self.entryPath_save.insert(0, tkFileDialog.asksaveasfilename(defaultextension='.sql', initialdir='/',title='Select file', filetypes=(('sql files', '*.sql'), ('all files', '*.*'))))
		self.saveWindow.lift()

	def pathDEFunction(self):
		self.DEentryPath.delete(0, END)
		self.DEentryPath.insert(0, tkFileDialog.asksaveasfilename(defaultextension='.xlsx', initialdir='/',title='Select file', filetypes=(('excel files','*.xlsx'), ('all files', '*.*'))))
		self.DEWindow.lift()

	def pathDSFunction(self):
		self.DSentryPath.delete(0, END)
		self.DSentryPath.insert(0, tkFileDialog.asksaveasfilename(defaultextension='.sql', initialdir='/',title='Select file', filetypes=(('sql files', '*.sql'), ('all files', '*.*'))))
		self.DSWindow.lift()

	def Excel_SQLFunction(self):
		try:
			self.textB.delete(1.0, END)
			f = open(self.entryPath_save.get(), 'w')
			if self.info['Sheet'] == 'all':
				for sheetnamesList in self.sheetList:
					for i in range(0, len(self.sendSQL[sheetnamesList])):
						f.write(self.sendSQL[str(sheetnamesList)][i] + ';\n\n')
			else:
				for i in range(0, len(self.sendSQL[self.name])):
					f.write(self.sendSQL[self.name][i] + ';\n\n')
			f.close()
			self.textB.insert(1.0, 'Excel -> SQL File Complete!\n\n')
			self.textB.insert(END, self.entryPath_save.get())
			self.info['Window'].destroy()
			f = open('C:\\Users\\Secuve\\Desktop\\Database Tool\\log\\log.txt', 'a')
			f.write(str('%s-%s-%s %s:%s:%s' %(datetime.datetime.now().year, datetime.datetime.now().month, datetime.datetime.now().day, datetime.datetime.now().hour, datetime.datetime.now().minute, datetime.datetime.now().second)) + '\tExcel -> SQL File Function.\t\t\t[ ' + str(self.info['Path'].encode('euc-kr')) + ' -> ' + self.entryPath_save.get() + ' ]' + '\n')
			f.close()
			self.saveWindow.destroy()
		except IOError:
			tkMessageBox.showwarning('Warning','Please select a SQL file to save.')
			self.saveWindow.lift()

	def DB_ExcelFunction(self):
		try:
			wb = openpyxl.Workbook()
			sheetmkNew = wb.active
			sheetmkNew.title = self.entrySheet.get()

			sheetmkNew.column_dimensions['A'].width = 1.5
			sheetmkNew.column_dimensions['B'].width = 25
			sheetmkNew.column_dimensions['C'].width = 15
			sheetmkNew.column_dimensions['D'].width = 15
			sheetmkNew.column_dimensions['E'].width = 10
			sheetmkNew.column_dimensions['F'].width = 10
			fontObj = Font(size=20, bold=True)
			fontBold = Font(bold = True)

			self.textB.delete(1.0, END)

			lineNum = 1
			lineCnt = 0
			for index in self.listboxDE.curselection():
				self.info['Cursor'].execute("SELECT COLUMN_NAME, DATA_TYPE, DATA_LENGTH FROM USER_TAB_COLUMNS WHERE TABLE_NAME = '" + str(self.listboxDE.get(index)[0]) +"'")
				saveList = self.info['Cursor'].fetchall()
				self.info['Cursor'].execute("SELECT COLUMN_NAME FROM USER_CONS_COLUMNS WHERE CONSTRAINT_NAME = '" + str(self.listboxDE.get(index)[0]) + "PK'")
				pkList = self.info['Cursor'].fetchall()
				self.info['Cursor'].execute("SELECT SEARCH_CONDITION FROM ALL_CONSTRAINTS WHERE TABLE_NAME = '" + str(self.listboxDE.get(index)[0]) + "'")
				nullList = self.info['Cursor'].fetchall()

				sheetmkNew['A' + str(lineNum)].font = fontObj
				sheetmkNew['B' + str(lineNum+1)].font = fontBold
				sheetmkNew['C' + str(lineNum+1)].font = fontBold
				sheetmkNew['D' + str(lineNum+1)].font = fontBold
				sheetmkNew['E' + str(lineNum+1)].font = fontBold
				sheetmkNew['F' + str(lineNum+1)].font = fontBold

				cnt = lineNum + 2
				sheetmkNew.cell(row=lineNum, column=1).value = str(self.listboxDE.get(index)[0])
				sheetmkNew.cell(row=lineNum+1, column=2).value = 'COLUMN_NAME'
				sheetmkNew.cell(row=lineNum+1, column=3).value = 'DATA_TYPE'
				sheetmkNew.cell(row=lineNum+1, column=4).value = 'DATA_LENGTH'
				sheetmkNew.cell(row=lineNum+1, column=5).value = 'NULL'
				sheetmkNew.cell(row=lineNum+1, column=6).value = 'KEY'

				for into in saveList:
					sheetmkNew.cell(row=cnt, column=2).value = into[0]
					sheetmkNew.cell(row=cnt, column=3).value = into[1]
					sheetmkNew.cell(row=cnt, column=4).value = into[2]
					cnt += 1

				nullName = []
				for nulllist in nullList:
					for i in range(1, len(str(nulllist[0]))):
						if str(nulllist[0])[i] == '"':
							nullName.append(str(nulllist[0])[1:i])

				cnt = lineNum + 2
				for count in range(0, len(saveList)):
					for n in range(0, len(nullName)):
						if nullName[n] in saveList[count]:
							sheetmkNew.cell(row=cnt, column=5).value = 'N'
					cnt += 1

				cnt = lineNum + 2
				for count in range(0, len(saveList)):
					for k in range(0, len(pkList)):
						if pkList[k][0] in saveList[count]:
							sheetmkNew.cell(row=cnt, column=6).value = 'PK'
					cnt += 1
				lineNum += len(saveList) + 4
				lineCnt += 1
			wb.save(self.DEentryPath.get())
			self.textB.insert(1.0, 'DB Scheme -> Excel Complete!\n\n')
			self.textB.insert(END, self.DEentryPath.get())
			f = open('C:\\Users\\Secuve\\Desktop\\Database Tool\\log\\log.txt', 'a')
			f.write(str('%s-%s-%s %s:%s:%s' %(datetime.datetime.now().year, datetime.datetime.now().month, datetime.datetime.now().day, datetime.datetime.now().hour, datetime.datetime.now().minute, datetime.datetime.now().second)) + '\tDB Scheme -> Excel File Function.\t[ ' + self.DEentryPath.get() + ' ]' + '\n')
			f.close()
			self.DEWindow.destroy()
		except IOError:
			tkMessageBox.showwarning('Warning','Please select a Excel file to save.')
			self.DEWindow.lift()
		except ValueError:
			tkMessageBox.showwarning('Warning','Please fill out the Sheet name.')
			self.DEWindow.lift()

	def DB_SQLFunction(self):
		try:
			self.textB.delete(1.0, END)
			SQLsentenceLast = ''
			for index in self.listboxDS.curselection():
				self.info['Cursor'].execute("SELECT COLUMN_NAME, DATA_TYPE, DATA_LENGTH FROM USER_TAB_COLUMNS WHERE TABLE_NAME = '" + str(self.listboxDS.get(index)[0]) + "'")
				sqlList = self.info['Cursor'].fetchall()
				self.info['Cursor'].execute("SELECT COLUMN_NAME FROM USER_CONS_COLUMNS WHERE CONSTRAINT_NAME = '" + str(self.listboxDS.get(index)[0]) + "PK'")
				pkList = self.info['Cursor'].fetchall()
				self.info['Cursor'].execute("SELECT SEARCH_CONDITION FROM ALL_CONSTRAINTS WHERE TABLE_NAME = '" + str(self.listboxDS.get(index)[0]) + "'")
				nullList = self.info['Cursor'].fetchall()

				SQLsentence = 'CREATE TABLE ' + str(self.listboxDS.get(index)[0]) + ' ( '
				SQLsentenceList = []
				nllist = []
				for sqllist in sqlList:
					SQLsentenceList = SQLsentenceList + [sqllist[0] + ' ' + sqllist[1] + '(' + str(int(sqllist[2])) + ')']
				for nulllist in nullList:
					for i in range(1, len(str(nulllist[0]))):
						if str(nulllist[0])[i] == '"':
							nllist.append(str(nulllist[0])[1:i])
				for count in range(0, len(SQLsentenceList)):
					for j in range(0, len(nllist)):
						if nllist[j] in SQLsentenceList[count]:
							SQLsentenceList[count] = SQLsentenceList[count] + ' ' + 'NOT NULL'
					if count == len(SQLsentenceList)-1:
						break
					SQLsentenceList[count] = SQLsentenceList[count] + ','

					for k in range(0, len(pkList)):
						if pkList[k][0] in SQLsentenceList[count]:
							SQLsentenceList.append('CONSTRAINT ' + str(self.listboxDS.get(index)[0]) + 'pk PRIMARY KEY (' + pkList[k][0] + ')')

				for makeSentence in range(0, len(SQLsentenceList)):
					SQLsentence = SQLsentence + SQLsentenceList[makeSentence] + ' '
				SQLsentence = SQLsentence + ');\n\n'
				SQLsentenceLast += SQLsentence
			f = open(self.DSentryPath.get(), 'w')
			f.write(SQLsentenceLast)
			f.close()
			self.textB.insert(1.0, 'DB Scheme -> SQL File Complete!\n\n')
			self.textB.insert(END, self.DSentryPath.get())
			f = open('C:\\Users\\Secuve\\Desktop\\Database Tool\\log\\log.txt', 'a')
			f.write(str('%s-%s-%s %s:%s:%s' %(datetime.datetime.now().year, datetime.datetime.now().month, datetime.datetime.now().day, datetime.datetime.now().hour, datetime.datetime.now().minute, datetime.datetime.now().second)) + '\tDB Scheme -> SQL File Function.\t\t[ ' + self.DSentryPath.get() + ' ]' + '\n')
			f.close()
			self.DSWindow.destroy()
		except IOError:
			tkMessageBox.showwarning('Warning','Please select a SQL file to save.')
			self.DSWindow.lift()

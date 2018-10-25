from time import sleep
from tkinter import *
class Status:
	def __init__(self, textB):
		self.textB = textB
		self.textB.delete(1.0, END)
		self.statusCheck = True
	def statusFunction(self, string):
		length = float(str('1.' + str(len(string) + 5)))
		self.textB.delete(1.0, END)
		self.textB.insert(1.0, string + ' ...../')
		while(self.statusCheck):
			self.textB.delete(length, END)
			self.textB.insert(END, '-')
			sleep(0.1)
			self.textB.delete(length, END)
			self.textB.insert(END, '\\')
			sleep(0.1)
			self.textB.delete(length, END)
			self.textB.insert(END, '|')
			sleep(0.1)
			self.textB.delete(length, END)
			self.textB.insert(END, '#')
			sleep(0.1)
	def stopFunction(self, check):
		self.statusCheck = check
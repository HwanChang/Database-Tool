import MainFrame
from Tkinter import *

def main() :
# Create window.
	window = Tk()
	window.geometry('640x450+100+100')
	window.resizable(False, False)
	MainFrame.MainFrame(window)
	window.mainloop()
if __name__ == '__main__' :
	main()
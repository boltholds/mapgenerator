#utf-8
#!/home/bh/Projects/mapgenerator/envs/bin/python

from tkinter import *
from tkinter import filedialog
from maper import parse_xls,outjob,wide_map

class GUI:

	def __init__(self):
		self.map_dict = dict()
		self.load_file =''
		self.out_file =''
		self.file_legacy = None
	def Quit(self,ev):
		global root
		#del(map_dict)
		root.destroy()

	def LoadFile(self,ev): 
		self.out_file = filedialog.Open(root, filetypes = [('*.xlsx files', '.xlsx'),('*.xls files', '.xls')]).show()
		if self.out_file == '':
		    return
		if self.out_file.endswith(".xls"):
			self.file_legacy = 1
		#print(type(self.out_file))
		self.map_dict =  parse_xls(self.out_file,self.file_legacy)

	def SaveFile(self,ev):
		self.load_file = filedialog.SaveAs(root, filetypes = [('*.py files', '.py')]).show()
		if self.load_file == '':
		    return
		if not self.load_file.endswith(".py"):
		    fn+=".py"
		temp = wide_map(self.map_dict)
		outjob(temp,self.load_file)

			
if __name__ == '__main__':

	root = Tk()
	mygui = GUI()
	panelFrame = Frame(root, height = 60, bg = 'gray')
	textFrame = Frame(root, height = 340, width = 600)

	panelFrame.pack(side = 'top', fill = 'x')
	textFrame.pack(side = 'bottom', fill = 'both', expand = 1)

	textbox = Text(textFrame, font='Arial 14', wrap='word')
	scrollbar = Scrollbar(textFrame)

	scrollbar['command'] = textbox.yview
	textbox['yscrollcommand'] = scrollbar.set

	textbox.pack(side = 'left', fill = 'both', expand = 1)
	scrollbar.pack(side = 'right', fill = 'y')

	loadBtn = Button(panelFrame, text = 'Load')
	saveBtn = Button(panelFrame, text = 'Save')
	quitBtn = Button(panelFrame, text = 'Quit')

	loadBtn.bind("<Button-1>", mygui.LoadFile)
	saveBtn.bind("<Button-1>", mygui.SaveFile)
	quitBtn.bind("<Button-1>", mygui.Quit)

	loadBtn.place(x = 10, y = 10, width = 40, height = 40)
	saveBtn.place(x = 60, y = 10, width = 40, height = 40)
	quitBtn.place(x = 110, y = 10, width = 40, height = 40)
	
	root.title("Менеджер по созданию map.py")
	root.mainloop()

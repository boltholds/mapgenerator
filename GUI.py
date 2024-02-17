#utf-8

import tkinter as tk
from tkinter import filedialog 
from tkinter.scrolledtext import ScrolledText
from maper import parse_xls,outjob,wide_map
import logging
from pathlib import Path
#logger = logging.getLogger(__name__)



class WidgetLogger(logging.Handler):
    def __init__(self, widget):
        logging.Handler.__init__(self)
        self.setLevel(logging.DEBUG)
        self.widget = widget
        self.widget.config(state='disabled')
        self.widget.tag_config("INFO", foreground="black")
        self.widget.tag_config("DEBUG", foreground="grey")
        self.widget.tag_config("WARNING", foreground="orange")
        self.widget.tag_config("ERROR", foreground="red")
        self.widget.tag_config("CRITICAL", foreground="red", underline=1)

        self.red = self.widget.tag_configure("red", foreground="red")
        
    def emit(self, record):
        self.widget.config(state='normal')
        # Append message (record) to the widget
        self.widget.insert(tk.END, self.format(record) + '\n', record.levelname)
        self.widget.see(tk.END)  # Scroll to the bottom
        self.widget.config(state='disabled') 
        self.widget.update() # Refresh the widget


class GUI:

	def __init__(self):
		self.map_dict = dict()
		self.load_file =''
		self.out_file =''

	def Quit(self,ev):
		global root
		#del(map_dict)
		root.destroy()

	def LoadFile(self,ev): 
		self.out_file = filedialog.Open(root, filetypes = [('*.xls files', '.xls'),('*.xlsx files', '.xlsx')]).show()
		if self.out_file == '':
			logger.error('Not open file! Try again!')
			return
		logger.info(f"Open file {self.out_file}")
		self.map_dict = parse_xls(self.out_file)
		logger.info('Complete! Save result file')

	def SaveFile(self, ev):
		self.load_file = filedialog.SaveAs(root, filetypes = [('*.py files', '.py')]).show()
		if self.load_file == '':
		    return
		if not self.load_file.endswith(".py"):
		    fn+=".py"
		temp = wide_map(self.map_dict)
		logger.info(f"Write file {self.load_file}")
		outjob(self.load_file,temp)
		logger.info('Complete!')
	


			
if __name__ == '__main__':

	root = tk.Tk()
	mygui = GUI()
	panelFrame = tk.Frame(root, height = 60, bg = 'gray')
	textFrame = tk.Frame(root, height = 340, width = 600)

	panelFrame.pack(side = 'top', fill = 'x')
	textFrame.pack(side = 'bottom', fill = 'both', expand = 1)

	
	
	
	st = ScrolledText(textFrame, state='disabled')
	st.configure(font='TkFixedFont')
	st.pack()
	# Create textLogger
	text_handler = WidgetLogger(st)

	# Add the handler to logger
	logger = logging.getLogger()
	logger.addHandler(text_handler)
	logger.setLevel(logging.INFO)
	# Log some messages

	logger.info('Load table adress file')


	
	loadBtn = tk.Button(panelFrame, text = 'Load')
	saveBtn = tk.Button(panelFrame, text = 'Save')
	quitBtn = tk.Button(panelFrame, text = 'Quit')

	loadBtn.bind("<Button-1>", mygui.LoadFile)
	saveBtn.bind("<Button-1>", mygui.SaveFile)
	quitBtn.bind("<Button-1>", mygui.Quit)

	loadBtn.place(x = 10, y = 10, width = 40, height = 40)
	saveBtn.place(x = 60, y = 10, width = 40, height = 40)
	quitBtn.place(x = 110, y = 10, width = 40, height = 40)

	root.title("Менеджер по созданию map.py")
	file_ico = set(Path.cwd().rglob(f"{__name__}.png"))
	if file_ico:
		root.iconphoto(False, tk.PhotoImage(file=file_ico.pop()))
	root.mainloop()

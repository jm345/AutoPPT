from tkinter import messagebox,Button,Label,Text,Tk,Entry,Frame,PhotoImage,filedialog,StringVar,Menu,Toplevel,Checkbutton,IntVar,LabelFrame,ttk

class MainWindow:

	def __init__(self,master):

		self.master = master

		# Variables
		self.is_checked = IntVar()
		self.user_title_string = 'Furniture Range'

		# Widgets
		self.master.title('AutoPPT v1.0')
		self.credential_frame = Frame(self.master)
		self.label_info = Label(self.credential_frame,text = 'Customise your presentation using the fields below.')
		self.label_user = Label(self.credential_frame,text = 'Title Slide: ')
		#self.label_password = Label(self.credential_frame, text = 'Password: ')
		self.entry_user = Entry(self.credential_frame,width = 40)
		#self.entry_password = Entry(self.credential_frame, width = 40,show = '*')
		self.button_verify = Button(self.master,text = 'Create Presentation',command = self.set_user_title_string)
		#self.check_password = Checkbutton(self.credential_frame, variable = self.is_checked)
		#self.label_check_password = Label(self.credential_frame,text = 'Use Current Date')
		#self.pb = ttk.Progressbar(self.master,orient ="horizontal",length = 200, mode ="determinate")
		#self.pb.start(50)
		
		# Layout
		self.credential_frame.grid(row = 0,column = 0,padx = 5,pady = 5)
		self.label_info.grid(columnspan = 2,padx = 10,pady = 10)
		self.label_user.grid(row = 1,column = 0,padx = 5,pady = 5,sticky = 'W')
		#self.label_password.grid(row=2, column=0, padx=5, pady=5,sticky = 'W')
		self.entry_user.grid(row=1, column=1, padx=5, pady=5)
		#self.entry_password.grid(row=2, column=1, padx=5, pady=5)
		self.button_verify.grid(padx = 10,pady = 10)
		#self.label_check_password.grid(padx = 5,pady = 5, sticky = 'W')
		#self.check_password.grid(pady = 5,column = 1,row = 3, sticky = 'W')
		#self.pb.grid(padx = 10,pady = 10)
		
	def set_user_title_string(self):
		self.user_title_string = self.entry_user.get()
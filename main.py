import clr
import Orca
import os
import threading
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Windows.Forms import *
from System.Drawing import Size,Point,Color
from time import sleep



def printer():
		print("sleeping 5 seconds")
		sleep(5)
		print("sleeping 5 seconds")
		sleep(5)
		print("sleeping 5 seconds")

def dprint(string):
	global tbox
	tbox.AppendText(string + "\r\n")
	#tbox.ScrollToEnd()

def cprint(string):
	if('tbox' in locals() or 'tbox' in globals()):
		global tbox
		first = tbox.SelectionColor
		tbox.SelectionColor = Color.Red
		tbox.AppendText(string + "\r\n")
		#tbox.ScrollToEnd()
		tbox.SelectionColor = first
	else:
		print(string)
	
class ProcessThread():
	Target = None
	Callback = None
	PreRun = None
	Active = False
	def SetPreRun(self,prerun):
		self.PreRun = prerun
	def SetTarget(self,target):
		self.Target = target
	def SetCallback(self,callback):
		self.Callback = callback
	def ThreadMethod(self):
		print("Thread Method Called")
		if(self.Active == False):
			self.Active = True
			self.PreRun()
			self.Target()
			self.Callback()
			
			self.Active = False
			self.Target = None
			self.Callback = None
			self.PreRun = None
		else:
			print("Thread already active")
	def RunThread(self):
		t = threading.Thread(target=self.ThreadMethod)
		t.start()
	
class OrcaForm(Form):
	def __init__(self):
		self.OrcaThread = ProcessThread()
		self.Text = "Orca v1"
		self.Size = Size(870,820)
		
		mainMenu = MainMenu()
		menuItem1 = MenuItem()
		menuItem2 = MenuItem()
		menuItem1.Text = "File"
		menuItem2.Text = "Hello"
		menuItem2.Click += self.SleepTest
		#menuItem1.Click += self.Disabler
		#mainMenu.MenuItems.Add(menuItem1)
		#mainMenu.MenuItems.Add(menuItem2)
		self.Menu = mainMenu
		
		self.tabControl1 = TabControl()
		self.tabPage1 = TabPage()
		self.tabPage1.Text = "Old Citrix"
		self.tabPage2 = TabPage()
		self.tabPage2.Text = "New Citrix"
		self.tabControl1.Size = Size(835,420)
		self.tabControl1.Location = Point(10,15)

		self.tabControl1.Parent = self
		self.tabControl1.Controls.Add(self.tabPage1)
		self.tabControl1.Controls.Add(self.tabPage2)
		
		self.userBox = TextBox()
		self.userBox.Location = Point(80,50)
		self.userBox.Width = 220
		self.userBox.Parent = self.tabPage1
		label = Label()
		label.Text = "Username"
		label.Location = Point(10,50)
		label.Parent = self.tabPage1
		btn_MigrateOld = Button()
		btn_MigrateOld.Text = "Migrate Old!"
		btn_MigrateOld.Location = Point(320,50)
		btn_MigrateOld.Click += self.MigrateOld
		btn_MigrateOld.Parent = self.tabPage1
		
		self.tbox = RichTextBox()
		self.tbox.Parent = self
		self.tbox.Location = Point(10,450)
		self.tbox.Multiline = True
		self.tbox.Width = 835
		self.tbox.Height = 300
		self.tbox.BackColor = Color.Black
		self.tbox.ForeColor = Color.LightGreen
		self.tbox.ScrollBars = ScrollBars.Vertical
		self.tbox.ReadOnly = True
		self.tbox.HideSelection = False
		self.tbox.DetectUrls = False
		
		self.Load += self.Loaded

	def MigrateOld(self,sender,args):
		self.tbox.AppendText("Migrating Old %s\r\n" % self.userBox.Text)
		message = "Are you sure you want to migrate %s?" % self.userBox.Text
		result = MessageBox.Show(message ,"Migration Confirmation", MessageBoxButtons.OKCancel)
		if(result == DialogResult.OK):
			cprint("Migrating %s" % self.userBox.Text)
			MigrateUser = Orca.OldCitrix(self.userBox.Text)
			self.OrcaThread.SetTarget(MigrateUser.Migrate)
			self.OrcaThread.SetCallback(self.EnableInput)
			self.OrcaThread.SetPreRun(self.DisableInput)
			self.OrcaThread.RunThread()
		else:
			cprint("Migration Cancelled")
	def DisableInput(self):
		self.tabPage1.Enabled = False 
		print("Disabling tab1")
	def EnableInput(self):
		self.tabPage1.Enabled = True
		print("Enabling Tab1")
	def SleepTest(self,sender,args):
		t = threading.Thread(target=printer,args=(self.tbox,))
		t.start()
			
	def BtnClick(self,sender,args):
		self.tbox.AppendText("added " + self.lb_old.SelectedItem + "\r\n")
	
	def Loaded(self,sender,args):
		Orca.SetTbox(self.tbox)
		global tbox
		global print
		tbox = self.tbox
		print=dprint
Application.Run(OrcaForm())
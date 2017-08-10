import os, sys
import win32api
import win32con
import win32file
import win32security
import ntsecuritycon
import pywintypes
import clr
clr.AddReference("System.Drawing")
from win32com.shell import shell,shellcon
from time import sleep
import colorama
from System.Drawing import Color
from pathlib import Path
#username=win32api.GetUserName()
global tbox

class ShellFolder:
	

	# self.sourceSubDir = os.path.join(source,sourceSubDir)
	def __init__(self):
		self.source = None
		self.subsource = None
		self.target = None
		self.subtarget = None
		self.user = None
	def MigratePreCheck(self):
		if not os.path.exists(self.target):
			try:
				path = Path(self.target)
				path.mkdir(parents=True,exist_ok=True)
				return True
			except Exception as er:
				return False
		else:
			return True
	def __str__(self):
		return self.target
	def Clean(self):
		Unseal(win32api.GetUserName(),self.target)
		Clean(self.target)
		Reseal(self.user,self.target)
	def Migrate(self):
		current_user = win32api.GetUserName()
		print("ShellFolder.Migrate %s to %s for %s" % (self.source,self.target,current_user))
		Unseal(current_user,self.source)
		Unseal(current_user,self.target)
		if self.MigratePreCheck():
			CopyFolder(self.subsource,self.subtarget)
			Clean(self.target)
			Reseal(self.user,self.target)
			Reseal(self.user,self.source)
		else:
			print("[ShellFolder->Migrate] Skipping %s as either the source or target do not exist" % self.source)
		
	def UnsealSource(self):
		current_user = win32api.GetUserName()
		print("ShellFolder.Unseal %s" % self.source)
		Unseal(current_user,self.source)
		
	def NukeFolder(self):
		NukeFolder(self.source)
		NukeFolder(self.target)
		
class OldCitrix():
	
	def __init__(self,username):
		# self.documents = ShellFolder("\\\\ctx-old\\mydocuments$\\%s" % username,"Documents", "\\\\ctx-new\\ctx_folder_redir$\\%s\\Documents" % username, username)
		self.documents = ShellFolder()
		self.documents.source = "\\\\ctx-old\\mydocuments$\\%s" % username
		self.documents.subsource = "\\\\ctx-old\\mydocuments$\\%s\\Documents\\*" % username
		self.documents.target = "\\\\ctx-new\\ctx_folder_redir$\\%s\\Documents" % username
		self.documents.subtarget = "\\\\ctx-new\\ctx_folder_redir$\\%s\\Documents\\" % username
		self.documents.user = username
		
		# self.desktop = ShellFolder("\\\\ctx-old\\mydesktop$\\%s" % username,"Desktop", "\\\\ctx-new\\ctx_folder_redir$\\%s\\Desktop" % username , username )
		self.desktop = ShellFolder()
		self.desktop.source = "\\\\ctx-old\\mydesktop$\\%s" % username
		self.desktop.subsource = "\\\\ctx-old\\mydesktop$\\%s\\Desktop\\*" % username
		self.desktop.target = "\\\\ctx-new\\ctx_folder_redir$\\%s\\Desktop" % username
		self.desktop.subtarget = "\\\\ctx-new\\ctx_folder_redir$\\%s\\Desktop\\" % username
		self.desktop.user = username
		
		# self.videos = ShellFolder("\\\\ctx-old\\myvideos$\\%s" % username,"Videos","\\\\ctx-new\\ctx_folder_redir$\\%s\\Videos" % username,username)
		self.videos = ShellFolder()
		self.videos.source = "\\\\ctx-old\\myvideos$\\%s" % username
		self.videos.subsource = "\\\\ctx-old\\myvideos$\\%s\\Videos\\*" % username
		self.videos.target = "\\\\ctx-new\\ctx_folder_redir$\\%s\\Videos" % username
		self.videos.subtarget = "\\\\ctx-new\\ctx_folder_redir$\\%s\\Videos\\" % username
		self.videos.user = username
		
		# self.downloads = ShellFolder("\\\\ctx-old\\mydownloads$\\%s" % username, "Downloads", "\\\\ctx-new\\ctx_folder_redir$\\%s\\Documents\\Downloads" % username, username)
		self.downloads = ShellFolder()
		self.downloads.source = "\\\\ctx-old\\mydownloads$\\%s" % username
		self.downloads.subsource = "\\\\ctx-old\\mydownloads$\\%s\\Downloads\\*" % username
		self.downloads.target = "\\\\ctx-new\\ctx_folder_redir$\\%s\\Documents\\Downloads" % username
		self.downloads.subtarget = "\\\\ctx-new\\ctx_folder_redir$\\%s\\Documents\\Downloads\\" % username
		self.downloads.user = username
		
		# self.pictures = ShellFolder("\\\\ctx-old\\mypictures$\\%s" % username, "Pictures", "\\\\ctx-new\\ctx_folder_redir$\\%s\\Documents\\Pictures" % username,username)
		self.pictures = ShellFolder()
		self.pictures.source = "\\\\ctx-old\\mypictures$\\%s" % username
		self.pictures.subsource = "\\\\ctx-old\\mypictures$\\%s\\Pictures\\*" % username
		self.pictures.target = "\\\\ctx-new\\ctx_folder_redir$\\%s\\Documents\\Pictures" % username
		self.pictures.subtarget = "\\\\ctx-new\\ctx_folder_redir$\\%s\\Documents\\Pictures\\" % username
		self.pictures.user = username
		
		# self.music = ShellFolder("\\\\ctx-old\\mymusic$\\%s" % username,"Music", "\\\\ctx-new\\ctx_folder_redir$\\%s\\Documents\\Music" % username,username)
		self.music = ShellFolder()
		self.music.source = "\\\\ctx-old\\mymusic$\\%s" % username
		self.music.subsource = "\\\\ctx-old\\mymusic$\\%s\\Music\\*" % username
		self.music.target = "\\\\ctx-new\\ctx_folder_redir$\\%s\\Documents\\Music" % username
		self.music.subtarget = "\\\\ctx-new\\ctx_folder_redir$\\%s\\Documents\\Music\\" % username
		self.music.user = username
		
		# self.favorites = ShellFolder("\\\\ctx-old\\favorites$\\%s" % username, "Favorites", "\\\\ctx-new\\ctx_folder_redir$\\%s\\Favorites" % username, username)
		self.favorites = ShellFolder()
		self.favorites.source = "\\\\ctx-old\\favorites$\\%s" % username
		self.favorites.subsource = "\\\\ctx-old\\favorites$\\%s\\Favorites\\*" % username
		self.favorites.target = "\\\\ctx-new\\ctx_folder_redir$\\%s\\Favorites" % username
		self.favorites.subtarget = "\\\\ctx-new\\ctx_folder_redir$\\%s\\Favorites\\" % username
		self.favorites.user = username
		
		
	def Check(self):
		print("Documents:\t%s" % self.documents)
		print("Desktop:\t%s" % self.desktop)
		print("Downloads:\t%s" % self.downloads)
		print("Favorites:\t%s" % self.favorites)
		print("Pictures:\t%s" % self.pictures)
		print("Music:\t%s" % self.music)
		print("Videos:\t%s" % self.videos)
		self.Sleep()
		
	def Migrate(self):
		self.documents.Migrate()
		self.desktop.Migrate()
		self.downloads.Migrate()
		self.favorites.Migrate()
		self.pictures.Migrate()
		self.music.Migrate()
		self.videos.Migrate()
	
	def UnsealSource(self):
		self.documents.UnsealSource()
		self.desktop.UnsealSource()
		self.downloads.UnsealSource()
		self.favorites.UnsealSource()
		self.pictures.UnsealSource()
		self.music.UnsealSource()
		self.videos.UnsealSource()
		
	#ReEnable if you Dare... Kind of Dangerous
	def NukeUser(self):
		self.documents.NukeFolder() 
		self.desktop.NukeFolder() 
		self.downloads.NukeFolder()
		self.favorites.NukeFolder()
		self.pictures.NukeFolder()
		self.music.NukeFolder()
		self.videos.NukeFolder()
		
def SetTbox(mytbox):
	global tbox
	tbox = mytbox
	tbox.AppendText("Orca Libraries Initialized\r\n")
	tbox.AppendText("Overriding Print\r\n")
	global print
	print = dprint
	
def dprint(string):
	tbox.AppendText(string + "\r\n")

def cprint(string,color='red'):
	if('tbox' in locals() or 'tbox' in globals()):
		global tbox
		scolor = None
		if color == 'red': scolor = Color.Red
		if color == 'yellow': scolor = Color.Yellow
		first = tbox.SelectionColor
		tbox.SelectionColor = Color.Red
		tbox.AppendText(string + "\r\n")
		tbox.SelectionColor = first
	else:
		scolor = None
		if color == 'red': scolor = colorama.Fore.RED + colorama.Style.BRIGHT 
		if color == 'yellow': scolor = colorama.Fore.YELLOW + colorama.Style.BRIGHT 
		print(scolor + string + colorama.Style.RESET_ALL)

def DeleteFolder(source):
	print("DeleteFolder %s" % source)
	shell.SHFileOperation((0,shellcon.FO_DELETE,source,None,shellcon.FOF_NOCONFIRMATION,None,None))
	
def CopyFolder(source,target):
	cprint("Copying Folder Source %s to: %s " % (source,target),"yellow")
	try:
		shell.SHFileOperation((0,shellcon.FO_COPY,source,target,shellcon.FOF_SILENT | shellcon.FOF_RENAMEONCOLLISION | shellcon.FOF_NO_UI |shellcon.FOF_NOCONFIRMMKDIR | shellcon.FOF_NOCOPYSECURITYATTRIBS | shellcon.FOF_NOCONFIRMATION,None,None))
	except Exception as er:
		print("Move Folder Exception: %s" % type(er).__name__)
		
def Clean(folder):
	print("Cleaning Folder %s" % folder)
	for root,dirs,files in os.walk(folder):
		for name in dirs:
			if(name == "$RECYCLE.BIN"):
				cprint("Cleanup %s" % os.path.join(root,name),"yellow")
				Unseal(win32api.GetUserName(),os.path.join(root,name))
				DeleteFolder(os.path.join(root,name))

def TakeOwnership(username,folder):
	try:
		account_sid = (win32security.LookupAccountName(None,username))[0]
		print("Setting " + username + " to owner on " + folder)
		sd = win32security.SECURITY_DESCRIPTOR()
		sd.SetSecurityDescriptorOwner(account_sid,0)
		win32security.SetFileSecurity(folder,win32security.OWNER_SECURITY_INFORMATION,sd)
	except Exception as er:
		print("TakeOwnership Exception: %s" % type(er).__name__)
		
def AddACL(username,folder):
	try:
		account_sid = (win32security.LookupAccountName(None,username))[0]
		everyone = win32security.LookupAccountName ("", "Everyone")[0]
		system = win32security.LookupAccountName ("", "System")[0]
		creator = win32security.LookupAccountName ("", "CREATOR OWNER")[0]
		domain_admins = win32security.LookupAccountName ("", "Domain Admins")[0]
		wolfbyte = win32security.LookupAccountName(None,"wolfbyte")[0]
		devon = win32security.LookupAccountName(None,"devon.dieffenbach")[0]
		dacl = win32security.ACL ()
		dacl.AddAccessAllowedAceEx (win32security.ACL_REVISION,3, ntsecuritycon.FILE_ALL_ACCESS, account_sid)
		dacl.AddAccessAllowedAceEx (win32security.ACL_REVISION,3, 2032127, system)
		dacl.AddAccessAllowedAceEx (win32security.ACL_REVISION,19, 2032127, domain_admins)
		dacl.AddAccessAllowedAceEx (win32security.ACL_REVISION,3, 2032127, wolfbyte)
		dacl.AddAccessAllowedAceEx (win32security.ACL_REVISION,19, 2032127, devon)
		dacl.AddAccessAllowedAceEx(win32security.ACL_REVISION,3, 2031616, creator)
		sd = win32security.GetFileSecurity (folder, win32security.DACL_SECURITY_INFORMATION)
		sd.SetSecurityDescriptorDacl (1, dacl, 0)
		win32security.SetFileSecurity (folder, win32security.DACL_SECURITY_INFORMATION, sd)
	except Exception as er:
		cprint("AddACL exception %s on %s" % (type(er).__name__,folder))
		
def ReadACL(file):
	info_dacl = win32security.GetFileSecurity(file,win32security.DACL_SECURITY_INFORMATION)
	acl_dacl = info_dacl.GetSecurityDescriptorDacl()
	info_owner = win32security.GetFileSecurity(file,win32security.OWNER_SECURITY_INFORMATION)
	owner= win32security.LookupAccountSid(None,info_owner.GetSecurityDescriptorOwner())[0]
	print(file + "   Owned by: " + owner)
	for x in range(acl_dacl.GetAceCount()):
		ace = acl_dacl.GetAce(x)
		name=win32security.LookupAccountSid(None,ace[2])[0]
		mask=ace[1]
		print("\t" + str(x) + "\t" + name + "\tMask:" + str(mask))
	return acl_dacl
	
def WalkACL(folder):
	for root,dirs,files in os.walk(folder,topdown=False):
		for name in dirs:
			ReadACL(os.path.join(root,name))

def Unseal(username,folder):
	entries=None
	try:
		print("Enumerating " + folder)
		TakeOwnership(username,folder)
		AddACL(username,folder)
		entries = list(os.scandir(folder))
	except PermissionError as er:
		cprint("\tUnseal Attempting to Take Ownership of " + folder)
		TakeOwnership(username,folder)
		cprint("\tUnseal Attempting to Grant Permissions to " + folder)
		AddACL(username,folder)
		cprint("\tUnseal Re-Iterating Unseal")
		Unseal(username,folder)
	except Exception as er:
		cprint("Unseal Unidentified exception %s on %s" % (type(er).__name__,folder))
	if entries is not None:
		for entry in entries:
			if entry.is_dir():
				print("Unseal Recursively Unsealing: " + entry.path)
				TakeOwnership(username,entry.path)
				AddACL(username,entry.path)
				Unseal(username,entry.path)
			if entry.is_file():
				print("Unseal Unsealing File %s" % entry.path)
				TakeOwnership(username,entry.path)
				AddACL(username,entry.path)
	else:
		cprint("There are no entries to iterate through on %s" % folder)
				
def Reseal(username,folder):
	for root,dirs,files in os.walk(folder,topdown=False):
		for name in dirs:
			print("Reseal - Granting Ownership of " + os.path.join(root,name) + "to " + username)
			AddACL(username,os.path.join(root,name))
			TakeOwnership(username,os.path.join(root,name))
		for name in files:
			print("Reseal Granting Ownership of " + os.path.join(root,name) + "to " + username)
			AddACL(username,os.path.join(root,name))
			TakeOwnership(username,os.path.join(root,name))
	AddACL(username,folder)
	TakeOwnership(username,folder)

def NukeFolder(folder):
	username=win32api.GetUserName()
	print("NukeFolder Unsealing " + folder + " for " +  username)
	Unseal(username,folder)
	print("[NukeFolder] starting os.walk on %s " % folder)
	for root,dirs,files in os.walk(folder,topdown=False):
		for name in dirs:
			print("[D] Dir found: " + os.path.join(root,name))
			try:
				win32api.SetFileAttributes(os.path.join(root,name),win32con.FILE_ATTRIBUTE_NORMAL)
				win32file.RemoveDirectory(os.path.join(root,name))
			except Exception as er:
				cprint("Exception on Removing Directory %s " % os.path.join(root,name))
		for name in files:
			print("[F] Delete: %s " % os.path.join(root,name))
			win32file.DeleteFile(os.path.join(root,name))
	
	try:
		win32api.SetFileAttributes(folder,win32con.FILE_ATTRIBUTE_NORMAL)
		win32file.RemoveDirectory(folder)
	except Exception as er:
		cprint("NukeFolder Exception on Setting File Attributes %s on %s " % (type(er).__name__, folder))

def NukeVDI():
	vdi1 = OldCitrix("vdi1")
	vdi2 = OldCitrix("vdi2")
	vdi3 = OldCitrix("vdi3")
	vdi4 = OldCitrix("vdi4")
	vdi1.NukeUser()
	vdi2.NukeUser()
	vdi3.NukeUser()
	vdi4.NukeUser()
	
	DeleteFolder("\\\\ctx-old\\p$\\UPM-XD-7\\VDI1")
	DeleteFolder("\\\\ctx-old\\p$\\UPM-XD-7\\VDI2")
	DeleteFolder("\\\\ctx-old\\p$\\UPM-XD-7\\VDI3")
	DeleteFolder("\\\\ctx-old\\p$\\UPM-XD-7\\VDI4")
	DeleteFolder("\\\\ctx-new\\upm$\\windows 10\\vdi1")
	DeleteFolder("\\\\ctx-new\\upm$\\windows 10\\vdi2")
	DeleteFolder("\\\\ctx-new\\upm$\\windows 10\\vdi3")
	DeleteFolder("\\\\ctx-new\\upm$\\windows 10\\vdi4")

class Cloner(ShellFolder):
	def Migrate(self):
		current_user = win32api.GetUserName()
		print("ShellFolder.Migrate %s to %s for %s" % (self.subsource,self.subtarget,current_user))
		Unseal(current_user,self.target)
		if self.MigratePreCheck():
			CopyFolder(self.subsource,self.target)
			Clean(self.target)
			Reseal(self.user,self.target)
		else:
			print("[ShellFolder->Migrate] Skipping %s as either the source or target do not exist" % self.source)
	
def CloneVDI(clone):
	username = "vdiClone"
	# documents = Cloner("\\\\ctx-old\\mydocuments$\\%s" % username,"Documents", "\\\\ctx-old\\mydocuments$\\%s" % clone, clone)
	# desktop = Cloner("\\\\ctx-old\\mydesktop$\\%s" % username, "Desktop", "\\\\ctx-old\\mydesktop$\\%s" % clone , clone )
	# videos = Cloner("\\\\ctx-old\\myvideos$\\%s" % username,"Videos","\\\\ctx-old\\myvideos$\\%s" % clone,clone)
	# downloads = Cloner("\\\\ctx-old\\mydownloads$\\%s" % username, "Downloads", "\\\\ctx-old\\mydownloads$\\%s" % clone, clone)
	# pictures = Cloner("\\\\ctx-old\\mypictures$\\%s" % username, "Pictures", "\\\\ctx-old\\mypictures$\\%s" % clone,clone)
	# music = Cloner("\\\\ctx-old\\mymusic$\\%s" % username,"Music", "\\\\ctx-old\\mymusic$\\%s" % clone,clone)
	# favorites = Cloner("\\\\ctx-old\\favorites$\\%s" % username, "Favorites", "\\\\ctx-old\\favorites$\\%s" % clone, clone)
	
	documents = ShellFolder()
	documents.source = "\\\\ctx-old\\mydocuments$\\%s" % username
	documents.subsource = "\\\\ctx-old\\mydocuments$\\%s\\Documents\\*" % username
	documents.target = "\\\\ctx-old\\mydocuments$\\%s" % clone
	documents.subtarget = "\\\\ctx-old\\mydocuments$\\%s\\Documents\\" % clone
	documents.user = clone
	
	desktop = ShellFolder()
	desktop.source = "\\\\ctx-old\\mydesktop$\\%s" % username
	desktop.subsource = "\\\\ctx-old\\mydesktop$\\%s\\Desktop\\*" % username
	desktop.target = "\\\\ctx-old\\mydesktop$\\%s" % clone
	desktop.subtarget = "\\\\ctx-old\\mydesktop$\\%s\\Desktop\\" % clone
	desktop.user = clone
	
	videos = ShellFolder()
	videos.source = "\\\\ctx-old\\myvideos$\\%s" % username
	videos.subsource = "\\\\ctx-old\\myvideos$\\%s\\Videos\\*" % username
	videos.target = "\\\\ctx-old\\myvideos$\\%s" % clone
	videos.subtarget = "\\\\ctx-old\\myvideos$\\%s\\Videos\\" % clone
	videos.user = clone
	
	downloads = ShellFolder()
	downloads.source = "\\\\ctx-old\\mydownloads$\\%s" % username
	downloads.subsource = "\\\\ctx-old\\mydownloads$\\%s\\Downloads\\*" % username
	downloads.target = "\\\\ctx-old\\mydownloads$\\%s" % clone
	downloads.subtarget = "\\\\ctx-old\\mydownloads$\\%s\\Downloads\\" % clone
	downloads.user = clone	

	pictures = ShellFolder()
	pictures.source = "\\\\ctx-old\\mypictures$\\%s" % username
	pictures.subsource = "\\\\ctx-old\\mypictures$\\%s\\Pictures\\*" % username
	pictures.target = "\\\\ctx-old\\mypictures$\\%s" % clone
	pictures.subtarget = "\\\\ctx-old\\mypictures$\\%s\\Pictures\\" % clone
	pictures.user = clone	
	
	music = ShellFolder()
	music.source = "\\\\ctx-old\\mymusic$\\%s" % username
	music.subsource = "\\\\ctx-old\\mymusic$\\%s\\Music\\*" % username
	music.target = "\\\\ctx-old\\mymusic$\\%s"% clone
	music.subtarget = "\\\\ctx-old\\mymusic$\\%s\\Music\\"  % clone
	music.user = clone	
	
	# favorites = ShellFolder()
	# favorites.source = "\\\\ctx-old\\mydesktop$\\%s" % username
	# favorites.subsource = "\\\\ctx-old\\mydesktop$\\%s\\Desktop\\*" % username
	# favorites.target = "\\\\ctx-old\\mydesktop$\\%s" % clone
	# favorites.subtarget = "\\\\ctx-old\\mydesktop$\\%s\\Desktop\\"% clone
	# favorites.user = clone
		
	documents.Migrate()
	desktop.Migrate()
	downloads.Migrate()
	#favorites.Migrate()
	pictures.Migrate()
	music.Migrate()
	videos.Migrate()
	#return documents,desktop,downloads,favorites,pictures,music,videos
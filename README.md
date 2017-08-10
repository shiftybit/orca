# Orca - used to migrate a redirected user shell folder from one server to another. 
## Use Case Scenarios
* Citrix XenApp/XenDesktop Profile Migrations.
* Unsealing (Take Ownership and Add Self to DACL) of a User Shell Redirected folder, for the purpose of seamlessly migrating the folder to a new location.
* Mass Migrations of User Shell Folders - You can use the Orca Library in your program to help mass migrations of protected user shares, and to ensure that the user shares are properly secured with the correct DACLs and Object owners afterwards

## How it works
* Orca.py
 * Makes heavy use of the win32api bindings. Reference Documentation here: http://timgolden.me.uk/pywin32-docs/PyWin32.html
 * Main.py
 * Draws a very simple GUI by binding to the .NET clr, and interfacing with System.Windows.Forms. You might be familiar with this, as it is a common way to build a GUI in visual studio with C# or other languages. Documentation Available here: http://pythonnet.github.io/

## Building from Source
* make.ps1 was created to automate this process. It does this by
 * Running PyInstaller -f main.py
 * Ensuring the build directory is clean. 

## PreRequisites
* make sure pyinstaller, pypiwin32, pythonnet, and colorama are installed
* tested on Python 3.5.2

![orca2](https://raw.githubusercontent.com/wolfbyte/orca/master/orca2.png)
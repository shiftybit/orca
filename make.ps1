$builds = $pwd.path + "\builds"
if(gcm pyinstaller -ErrorAction SilentlyContinue){
	Write-Host "PyInstaller detected"
}else{
	Write-Error "PyInstaller not Detected"
	return
}

pyinstaller -w -F main.py --distpath $builds --name Orca
Remove-Item -Recurse __pycache__
Remove-Item -Recurse build
Remove-Item orca.spec

Write-Host -Foreground Green "$builds\orca.exe successfully compiled."
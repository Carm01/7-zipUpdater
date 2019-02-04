#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Proycontec-Construction-Screen-settings.ico
#AutoIt3Wrapper_Outfile_x64=7-zipInstaller_x1.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=7-zip clean installer
#AutoIt3Wrapper_Res_Fileversion=1.3.0.0
#AutoIt3Wrapper_Res_ProductVersion=1.3.0.0
#AutoIt3Wrapper_Res_LegalCopyright=carm0
#AutoIt3Wrapper_Res_Language=1033
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
If UBound(ProcessList(@ScriptName)) > 2 Then Exit
#include <InetConstants.au3>
#include <File.au3>
#include <array.au3>
#include <Inet.au3>
#include <EventLog.au3>

Local $sVersions, $splita1, $z, $y, $ecode, $CVersion, $scheck = 0

If $CmdLine[0] >= 1 Then
	Call("line")
EndIf

SplashTextOn("Progress", "", 220, 60, -1, -1, 16, "Tahoma", 10)
;ControlSetText("Progress", "", "Static1", "Downloading and installing 7-zip version " & $sVersions, 2)
ControlSetText("Progress", "", "Static1", "Initializing", 2)

getfiles()
If $scheck = 1 Then
	compare()
EndIf
Uninstall()
Install()

Func compare()
	; determine current version
	$sCurrentVersion = RegRead('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\7-Zip', 'DisplayVersion')
	If $sVersions = $sCurrentVersion Then
		SplashOff()
		$CVersion = FileGetVersion("C:\Program Files\7-Zip\7z.exe", $FV_FILEVERSION)
		$ecode = '444'
		EventLog()
		MsgBox(0, "GrassHopper Says:", "You have the most current version of 7-zip", 5)
		Exit
	EndIf
EndFunc   ;==>compare


Func Uninstall()
	If FileExists('C:\Program Files\7-Zip\Uninstall.exe') Then
		$cmd1 = 'taskkill.exe /im explorer.exe /f'
		RunWait('"' & @ComSpec & '" /c ' & $cmd1, @SystemDir, @SW_HIDE)
		Sleep(175)
		ShellExecuteWait('Uninstall.exe', ' /S', 'C:\Program Files\7-Zip\', "", @SW_HIDE)
		Sleep(175)
		DirRemove('C:\Program Files\7-Zip\', 1)
	EndIf

	#cs
		If FileExists('C:\Program Files\WinRAR') Then
		ShellExecuteWait('Uninstall.exe', ' /S', 'C:\Program Files\WinRAR', "", @SW_HIDE)
		DirRemove('C:\Program Files\WinRAR', 1)
		EndIf
	#ce
EndFunc   ;==>Uninstall


Func Install()
	ControlSetText("Progress", "", "Static1", "Downloading and installing 7-zip version " & $sVersions, 2)
	$sSite = 'https://www.7-zip.org/' & $y[9]
	$sDLfile = "c:\windows\temp\" & $z[3]

	$hDownload7 = InetGet($sSite, $sDLfile, $INET_FORCERELOAD, $INET_DOWNLOADBACKGROUND)
	Do
		Sleep(800)
	Until InetGetInfo($hDownload7, $INET_DOWNLOADCOMPLETE)

	Sleep(1000)
	;https://windowsreport.com/open-file-security-warning/
	RegWrite('HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Associations', 'LowRiskFileTypes', 'REG_SZ', '.avi;.bat;.cmd;.exe;.htm;.html;.lnk;.mpg;.mpeg;.mov;.mp3;.mp4;.mkv;.msi;.m3u;.rar;.reg;.txt;.vbs;.wav;.zip;.7z')

	If Not FileExists('C:\Program Files\7-Zip\7z.exe') Then
		ShellExecuteWait($z[3], ' /S', "c:\windows\temp\", "", @SW_HIDE)
	EndIf

	If Not ProcessExists('explorer.exe') Then
		$cmd1 = 'start "Shell Restarter" /d "%systemroot%" /i /normal c:\windows\explorer.exe'
		RunWait('"' & @ComSpec & '" /c ' & $cmd1, @SystemDir, @SW_HIDE)
	EndIf
	$CVersion = FileGetVersion("C:\Program Files\7-Zip\7z.exe", $FV_FILEVERSION)
	FileDelete('c:\windows\temp\' & $z[3])
	$ecode = '411'
	EventLog()
EndFunc   ;==>Install


Func getfiles()
	; get latest download
	Local $sTxt, $sTxt1
	$xjs = "C:\windows\temp\xjs.tmp"
	;$xjs1 = "C:\windows\temp\xjs1.tmp"
	$sSite = "https://www.7-zip.org/"
	;$sNotes = "https://www.mozilla.org/en-US/firefox/notes/"
	$source = _INetGetSource($sSite)
	$sTxt = StringSplit($source, @LF)

	$x = 0
	$i = 0
	For $i = 1 To UBound($sTxt) - 1 ; is like saying read the line number
		;GUIGetMsg();prevent high cpu usage
		If StringInStr($sTxt[$i], 'Download 7-Zip') > 1 Then
			$sActiveX1 = StringSplit($sTxt[$i], 'Download 7-Zip', 1)
			$sActiveX2 = StringSplit($sTxt[$i], ' ')
			$sVersions = StringStripWS($sActiveX2[3], 3)
			$x = 1
		EndIf
		If $x = 1 And StringInStr($sTxt[$i], '-x64') > 1 Then
			$y = StringSplit($sTxt[$i], '="')
			$z = StringSplit($y[9], 'a/')
		EndIf
	Next

	If $i = UBound($sTxt) - 1 Then
		$ecode = '404'
		EventLog()
		Exit
	EndIf
EndFunc   ;==>getfiles



Func EventLog()

	If $ecode = '404' Then
		Local $hEventLog, $aData[4] = [0, 4, 0, 4]
		$hEventLog = _EventLog__Open("", "Application")
		_EventLog__Report($hEventLog, 1, 0, 404, @UserName, @UserName & ' No "exe" found for 7-zip. The webpage and/or download link might have changed. ' & @CRLF, $aData)
		_EventLog__Close($hEventLog)
	EndIf

	If $ecode = '411' Then
		Local $hEventLog, $aData[4] = [0, 4, 1, 1]
		$hEventLog = _EventLog__Open("", "Application")
		_EventLog__Report($hEventLog, 0, 0, 411, @UserName, @UserName & " 7-Zip " & "version " & $CVersion & " successfully installed." & @CRLF, $aData)
		_EventLog__Close($hEventLog)
	EndIf

	If $ecode = '444' Then
		Local $hEventLog, $aData[4] = [0, 4, 4, 4]
		$hEventLog = _EventLog__Open("", "Application")
		_EventLog__Report($hEventLog, 0, 0, 444, @UserName, @UserName & " The current version of 7-ZIP is already installed " & $CVersion & @CRLF, $aData)
		_EventLog__Close($hEventLog)
	EndIf

EndFunc   ;==>EventLog

Func line()

	For $z = 1 To UBound($CmdLine) - 1

		If StringInStr($CmdLine[$z], "-") <> 1 Then
			MsgBox(0, "Grasshopper Says:", 'Wrong switch please use a "-"')
			Exit
		EndIf
		; the -i command cannot be used alone but with one of the following a,n,s o install the selected players
		If StringInStr($CmdLine[$z], "c") = 2 Then
			$scheck = 1
		EndIf

		If StringInStr($CmdLine[$z], "c") <> 2 Then
			MsgBox(0, "Invalad parameter", "Valid parameters are currently:" & @CRLF & " -c (check and only reinstall if out of date)", 5)
			Exit
		EndIf
	Next
EndFunc   ;==>line

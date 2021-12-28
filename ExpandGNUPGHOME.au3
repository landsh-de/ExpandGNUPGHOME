#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=res\ExpandGNUPGHOME.ico
#AutoIt3Wrapper_Outfile=bin\ExpandGNUPGHOME.exe
#AutoIt3Wrapper_Outfile_x64=bin\ExpandGNUPGHOME64.exe
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_Res_Comment=Exporting necessary GNUPG variables into local USER environment with environment update at runtime or login.
#AutoIt3Wrapper_Res_Description=Exporting necessary GNUPG variables into local USER environment with environment update at runtime or login.
#AutoIt3Wrapper_Res_Fileversion=1.0.0.4
#AutoIt3Wrapper_Res_LegalCopyright=Copyright 2021 © Veit Berwig
#AutoIt3Wrapper_Res_Language=1031
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#AutoIt3Wrapper_Res_Field=Author|Veit Berwig
#AutoIt3Wrapper_Res_Field=Made By|Veit Berwig
#AutoIt3Wrapper_Res_Field=Info|This program exports necessary GNUPG variables into every local user environment at runtime or login.
#AutoIt3Wrapper_Run_Tidy=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#Region
#EndRegion

;
; Author: 	Veit Berwig
; Desc.: 	Expands the GNUPGHOME-var with its content into local
;   		USER environment with environment update at runtime or login.
;
; 			GPGHOME will be build from %APPDATA%\GnuPG but there are
; 			Race-Conditions in evaluating %APPDATA%, so the most safe way is
; 			to use the REG-Key-Way over "User Shell Folders".
;
; 			Now retrieve APPDATA from registry key "User Shell Folders" under
; 			HLCU first. This is the most copatible way for all OS'es.
; 			If this does not work use a the default value:
; 			"%USERPROFILE%\AppData\Roaming\GnuPG". Then read the ini-file if
; 			it is existent, otherwise use the above default value.
; 			When the ini-file contains valid values, use the ini file,
; 			otherwise use the default value above.

; Version: 	1.0.0.2
; Important Info:	We have to do a case sensitive    string-diff with:
; 					If Not ("String1" == "String2") Then ...
;					We have to do a case in-sensitive string-diff with:
; 					If Not ("String1" <> "String2") Then ...
; 					here !!
;
; Version: 	1.0.0.3
;
;	Update to support additional GnuPG v2 functions ...
;
;	Added %CAROOT% environment variable for X.509-support of ROOT CA
;	Certs of GnuPG v2 and support of "mkcert"-tool by "Filippo Valsorda"
;	(FiloSottile). With "mkcert" you may create secure self-signed X.509
;	certificates. When %CAROOT% exist, "mkcert" will look at %CAROOT% for
; 	"MKCERT_CA.pem" in order to create a self-signed X.509 certificate.
;
; 	Added deletion of variables from environment, when "-" is defined
; 	as a value for the specific variable.
;
;	All these environment variable are exported into local environment
; 	without hard-coding in user-registry.
;
; Version: 	1.0.0.4
;
;	Renamed %CAROOT% to %GNUPGCAROOT% because of naming-conflic with
;   "mkcert"-tool by "Filippo Valsorda", which is creating a
; 	user-based CA-certificate in a directory, represented by this
; 	variable. Generally the user has no write-access to the directory,
; 	represented by %GNUPGCAROOT%
;	(here: "Common AppData\GNU\etc\gnupg\trusted-certs").
;
;
; == 	Tests if two strings are equal. Case sensitive.
;		The left and right values are converted to strings if they are
;		not strings already. This operator should only be used if
;		string comparisons need to be case sensitive.
; <> 	Tests if two values are not equal. Case insensitive	when used
;		with strings. To do a case sensitive not equal comparison use
;		Not ("string1" == "string2")
;

#cs
	;**********************************************************************
	History:

	04.11.2016
	-	First release, adding, fixing...(can't remember)

	-	Update to version 1.0.0.1

	04.11.2016
	-	Update to version 1.0.0.2

	04.11.2021
	-	Update to version 1.0.0.3

	04.11.2021
	-	Update to version 1.0.0.4
	;**********************************************************************
#ce

#include <File.au3>
#include <string.au3>
#include <Constants.au3>
#include <StringConstants.au3>
#include <GuiConstants.au3>
#include <GUIConstantsEx.au3>
#include <FileConstants.au3>
#include <WinAPI.au3>
#include <Misc.au3>
#include <Date.au3>

; product name
Global $prod_name = "ExpandGNUPGHOME"

; generate dynamic name-instance from filename
Global $app_name = $prod_name
; retrieve short version of @ScriptName
Global $scriptname_short = StringTrimRight(@ScriptName, 4) ; Remove the 5 rightmost characters from the string.
If Not (StringLen($scriptname_short) = 0) Then
	$app_name = $scriptname_short
EndIf

Global $app_version = "1.0.0.4"
Global $app_copy = "Copyright 2021 © Veit Berwig"
Global $appname = $prod_name & " " & $app_version
Global $appGUID = $app_name & "-200359f6-1c29-4b53-ad17-91bdb38c3e2c"

Global $homedrive_, $homepath_
Global $sControllerpath, $sControllerpath_
Global $Config_File

Global $GNUPGHOME, $GNUPGHOMEloc, $GNUPGHOME_REG
Global $GNUPGCAROOT, $GNUPGCAROOTloc, $GNUPGCAROOT_REG
Global $GPGCAROOT = "GNU\etc\gnupg\trusted-certs"

; APPDATA in "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
; is under Windows 7:  %USERPROFILE%\AppData\Roaming
Global $AppDataUser, $iAppDataUserErr

; "Common AppData" in
; "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
;    "Common AppData"    REG_SZ         C:\ProgramData
;
; "Common AppData" in
; "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
;    "Common AppData"    REG_EXPAND_SZ  %ProgramData%
Global $AppDataSystem, $iAppDataSystemErr


; EnvMyUpdate
; Source:
; https://www.autoitscript.com/forum/topic/116484-solved-autoit3-encountered-a-problem-and-needs-to-close/#comment-813190
; #INDEX# ==========================================================
; Title .........: Environment Update
; AutoIt Version.: 3.2.12++
; Language.......: English
; Description ...: Refreshes the OS environment.
; Author ........: João Carlos (jscript)
; Support .......: trancexx, PsaltyDS, KaFu
; ==================================================================

; #VARIABLES# ======================================================
; 	Local Const $MAX_VALUE_NAME = 1024
;   Already defined in APISysConstants.au3, so locally defined here
; 	Local Const $HWND_BROADCAST = 0xffff
; 	Local Const $WM_MY_SETTINGCHANGE = 0x001A
;   Already defined in APISysConstants.au3, so locally defined here
; 	Local Const $SMTO_ABORTIFHUNG = 0x0002
;   Already defined in APISysConstants.au3, so locally defined here
; 	Local Const $SMTO_NORMAL = 0x0000
; 	Local Const $MSG_TIMEOUT = 5000

; #FUNCTION# =======================================================
; Name...........: _EnvUpdate
; Description ...: Refreshes the OS environment.
; Syntax.........: _EnvUpdate( ["envvariable" [, "value" [, CurrentUser [, Machine ]]]] )
; Parameters ....: envvariable  - [optional] Name of the environment variable to set.
;											 If no variable, refreshes all variables.
;                  value        - [optional] Value to set the environment variable to.
;											 If a value is not used the environment
;                                   		 variable will be deleted.
;                  CurrentUser  - [optional] Sets the variable in current user environment.
;                  Machine      - [optional] Sets the variable in the machine environment.
; Return values .: Success      - None
;                  Failure      - Sets @error to 1.
; Author ........: João Carlos (jscript)
; Support .......: trancexx, PsaltyDS, KaFu
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; _EnvMyUpdate("TEMP", @SystemDir & "TEMP", True, True)
; ==================================================================

Func _EnvMyUpdate($sEnvVar = "", $vValue = "", $fCurrentUser = True, $fMachine = False)
	Local Const $MAX_VALUE_NAME = 1024
	; Already defined in APISysConstants.au3, so locally defined here
	Local Const $HWND_BROADCAST = 0xffff
	Local Const $WM_MY_SETTINGCHANGE = 0x001A
	; Already defined in APISysConstants.au3, so locally defined here
	Local Const $SMTO_ABORTIFHUNG = 0x0002
	; Already defined in APISysConstants.au3, so locally defined here
	Local Const $SMTO_NORMAL = 0x0000
	Local Const $MSG_TIMEOUT = 5000

	Local $sREG_TYPE = "REG_SZ", $iRet1, $iRet2

	If $sEnvVar <> "" Then
		If StringInStr($vValue, "%") Then $sREG_TYPE = "REG_EXPAND_SZ"

		; If $vValue contains "" then mark $vValue for deletion ("-") and
		; delete $sEnvVar and $vValue from Environment key in registry.
		; I did this here, because "or"-comparison of two strings won't work
		; here.
		If Not $vValue <> "" Then $vValue = "-"

		If $vValue <> "-" Then
			If $fCurrentUser Then RegWrite("HKCU\Environment", $sEnvVar, $sREG_TYPE, $vValue)
			If $fMachine Then RegWrite("HKLM\System\CurrentControlSet\Control\Session Manager\Environment", $sEnvVar, $sREG_TYPE, $vValue)
		Else
			If $fCurrentUser Then RegDelete("HKCU\Environment", $sEnvVar)
			If $fMachine Then RegDelete("HKLM\System\CurrentControlSet\Control\Session Manager\Environment", $sEnvVar)
		EndIf

		; http://msdn.microsoft.com/en-us/library/ms686206%28VS.85%29.aspx
		; https://docs.microsoft.com/de-de/windows/win32/winmsg/wm-settingchange
		$iRet1 = DllCall("Kernel32.dll", "BOOL", "SetEnvironmentVariable", "str", $sEnvVar, "str", $vValue)
		If $iRet1[0] = 0 Then Return SetError(1)
	EndIf
	; http://msdn.microsoft.com/en-us/library/ms644952%28VS.85%29.aspx
	; https://docs.microsoft.com/de-de/windows/win32/winmsg/wm-settingchange
	$iRet2 = DllCall("user32.dll", "lresult", "SendMessageTimeoutW", _
			"hwnd", $HWND_BROADCAST, _
			"dword", $WM_MY_SETTINGCHANGE, _
			"ptr", 0, _
			"wstr", "Environment", _
			"dword", $SMTO_ABORTIFHUNG, _
			"dword", $MSG_TIMEOUT, _
			"dword_ptr*", 0)

	If $iRet2[0] = 0 Then Return SetError(1)
EndFunc   ;==>_EnvMyUpdate

Global $sDateTime = @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC

Opt("GUIOnEventMode", 1)
Opt("TrayOnEventMode", 1)
Opt("TrayAutoPause", 0) ; The script will not pause when selecting the tray icon.
Opt("TrayMenuMode", 2) ; Items are not checked when selected.

; Extend the behaviour of the script tray icon/menu.
; This can be done with a combination (adding) of the following values.
; 0 = default menu items (Script Paused/Exit) are appended to the usercreated menu;
;     usercreated checked items will automatically unchecked; if you double click
;     the tray icon then the controlid is returned which has the "Default"-style (default).
; 1 = no default menu
; 2 = user created checked items will not automatically unchecked if you click it
; 4 = don't return the menuitemID which has the "default"-style in the main
;     contextmenu if you double click the tray icon
; 8 = turn off auto check of radio item groups
Opt("TrayMenuMode", 10)

TrayItemSetText($TRAY_ITEM_EXIT, $app_name & " beenden ...") ; Set the text of the default 'Exit' item.
TrayItemSetText($TRAY_ITEM_PAUSE, $app_name & " anhalten ...") ; Set the text of the default 'Pause' item.

TraySetClick(16)
TraySetToolTip($app_name)

$sControllerpath = FileGetLongName(@ScriptDir)
; If whe have only 3 chars, then we are in the root-dir with a
; additional backslash at the end of the pathname. this will
; result in \\; so we have to fix this here.
If (StringLen($sControllerpath) = 3) Then
	$sControllerpath_ = StringRegExpReplace($sControllerpath, "([\\])", "")
Else
	$sControllerpath_ = $sControllerpath
EndIf

; debug-info
;MsgBox(0, "Controllerpath is:", $sControllerpath_)

; Check for running only one instance of process (in Misc.au3)
; $sOccurenceName String to identify the occurrence of the script.
; This string may not contain the \ character unless you are placing the
; object in a namespace (See Remarks).
;
; $iFlag [optional] Behavior options.
; 0 - Exit the script with the exit code -1 if another instance already exists.
; 1 - Return from the function without exiting the script.
; 2 - Allow the object to be accessed by anybody in the system. This is useful
;     if specifying a "Global\" object in a multi-user environment.
; You can place the object in a namespace by prefixing your object name with
; either "Global\" or "Local\". "Global\" objects combined with the flag 2 are
; useful in multi-user environments.
If _Singleton($appGUID, 1) = 0 Then
	MsgBox(16, $appname, "Eine Instanz dieses Programmes:" & @CRLF & '"' & $appname & '"' & @CRLF & "läuft schon im Hauptspeicher !" & @CRLF & @CRLF & "Bitte das Programm erst beenden !", 10)
	Exit
EndIf

; Read "AppData" from Registry User Shell Folders
$AppDataUser = RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "AppData")
$iAppDataUserErr = @error
If $iAppDataUserErr = 0 Then
	$GNUPGHOME_REG = $AppDataUser & "\" & "GnuPG"
Else
	$GNUPGHOME_REG = "%USERPROFILE%\AppData\Roaming\GnuPG"
EndIf

; Read "Common AppData" from System Registry "User Shell Folders" => VALUE: %ProgramData%
$AppDataSystem = RegRead("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Common AppData")
; Read "Common AppData" from System Registry "Shell Folders"      => VALUE: C:\ProgramData
; $AppDataSystem = RegRead("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Common AppData")
$iAppDataSystemErr = @error
If $iAppDataSystemErr = 0 Then
	$GNUPGCAROOT_REG = $AppDataSystem & "\" & $GPGCAROOT
Else
	$GNUPGCAROOT_REG = "%ProgramData%" & "\" & $GPGCAROOT
EndIf

; Build absolute file-pathname
$Config_File = FileGetLongName($sControllerpath_ & "\" & $app_name & ".ini")

; Install the ini-file if no ini-file is existent
If (FileExists($Config_File) <> 1) Then
	; Write the value of 'Value' to the key 'Key' in the section labelled 'Section'.
	; IniWrite("INI-File", "Section", "Key", "Value")
	IniWrite($Config_File, "Main Prefs", "GNUPGHOME", $GNUPGHOME_REG)
	IniWrite($Config_File, "Main Prefs", "GNUPGCAROOT", $GNUPGCAROOT_REG)
EndIf


; ------------ Read ini-file
$GNUPGHOME = IniRead($Config_File, "Main Prefs", "GNUPGHOME", $GNUPGHOME_REG)
; ------------ Update environment
If $GNUPGHOME <> "" Then
	If $GNUPGHOME = "-" Then
		_EnvMyUpdate("GNUPGHOME", "", True, False)
	Else
		$GNUPGHOMEloc = FileGetLongName($GNUPGHOME)
		_EnvMyUpdate("GNUPGHOME", $GNUPGHOMEloc, True, False)
	EndIf
EndIf


; ------------ Read ini-file
$GNUPGCAROOT = IniRead($Config_File, "Main Prefs", "GNUPGCAROOT", $GNUPGCAROOT_REG)
; ------------ Update environment
If $GNUPGCAROOT <> "" Then
	If $GNUPGCAROOT = "-" Then
		_EnvMyUpdate("GNUPGCAROOT", "", True, False)
	Else
		$GNUPGCAROOTloc = FileGetLongName($GNUPGCAROOT)
		_EnvMyUpdate("GNUPGCAROOT", $GNUPGCAROOTloc, True, False)
	EndIf
EndIf

Exit

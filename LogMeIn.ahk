#Include TF.ahk ;https://github.com/hi5/TF
#Include Crypt.ahk ;http://autohotkey.com/board/topic/67155-ahk-l-crypt-ahk-cryptography-class-encryption-hashing/page-1
#Include CryptConst.ahk
#Include CryptFoos.ahk
#Include RunAsAdmin.ahk ;http://autohotkey.com/board/topic/46526-run-as-administrator-xpvista7-a-isadmin-params-lib/

;**************************************************************************************************
; Declare/define variables and Auto Execute stuff
;**************************************************************************************************

Global hotkeyDir := A_ProgramFiles . "\Hotkeys"
Global dataDir := hotkeyDir . "\LogMeIn"
Global exePath := hotkeyDir . "\LogMeIn.exe"
Global hotStringFile := dataDir . "\LogMeInHotstrings.ahk"
Global hotStringExe := hotkeyDir . "\LogMeInHotstrings.exe"
Global configFile := dataDir . "\LogMeIn.ini"
Global refreshIconPath  := dataDir . "\Reload.png"
Global compilerDir := dataDir . "\AHK Compiler"
Global clientList := ""

FileGetTime, versionPlain, %exePath%, M
FormatTime, versionFormatted, %versionPlain%, yMMdd
Global version := "v" . versionFormatted
Global windowTitle := "LogMeIn (" . version . ")"

IfNotExist, %hotkeyDir%
	FileCreateDir, %hotkeyDir%
	
IfNotExist, %dataDir%
	FileCreateDir, %dataDir%

IfNotExist, %A_Startup%\LogMeIn.lnk
	FileCreateShortcut, %exePath%, %A_Startup%\LogMeIn.lnk
	
If A_IsCompiled
{
	IfNotExist, %exePath%
	{
		FileMove, %A_ScriptFullPath%, %exePath%
		Run, %exePath%
		ExitApp
	}
}

IniRead, wantReturn, %configFile%, settings, wantReturn
If(wantReturn = "ERROR")
{
	wantReturn := "1"
	IniWrite, %wantReturn%, %configFile%, settings, wantReturn
}

If A_IsCompiled
	FileInstall, Refresh.png, %refreshIconPath%

If not(A_IsCompiled)
	Menu, Tray, Icon, LogMeIn.ico

Menu, Tray, NoStandard
Menu, Tray, Add, Reload LogMeIn, reloadLogMeIn
Menu, Tray, Default, Reload LogMeIn
Menu, Tray, Add
Menu, Tray, Add, Reload HotStrings, reloadLogMeInHotstrings
Menu, Tray, Add, Suspend HotStrings, SuspendHotStrings
Menu, Tray, Add
;Menu, Tray, Add, Help, HelpFile
;Menu, Tray, Add, Email Developer, EmailDev
Menu, Tray, Add, Exit, ExitLogMeIn
Menu, Tray, Tip, %windowTitle%

Global LogMeInHotkeyTrigger
IniRead, LogMeInHotkeyTrigger, %configFile%, settings, LogMeInTrigger
if(not(LogMeInHotkeyTrigger) or LogMeInHotkeyTrigger = "ERROR")
	LogMeInHotkeyTrigger := "#F2"
Hotkey, %LogMeInHotkeyTrigger%, LogMeInHotkey, On

IfExist, %hotStringFile%
	FileDelete, %hotStringFile%

IfExist, %hotStringExe%
	Run, %hotStringExe%
Else
	UpdateHotStringsExe()
Return

;**************************************************************************************************
; Functions 
;**************************************************************************************************

;Removes blank lines and trims begin/end white space per line 
trimWhitespacePerLine(str)
{
	trimmedStr =
	Loop, parse, str, `n, `r
		if(A_LoopField) ;remove blank/duplicate lines
			trimmedStr .= regexreplace(regexreplace(A_LoopField, "\s+$"), "^\s+") . "`n" ;trim beginning/end whitespace
	StringTrimRight, trimmedStr, trimmedStr, 1			
	return trimmedStr
}

;Gets an IE object by the window title or the active window if no title is passed in.
IEGet(name="")
{
   IfEqual, Name,, WinGetTitle, Name, ahk_class IEFrame     ;Get active window if no parameter
   Name := (Name="New Tab - Windows Internet Explorer")? "about:Tabs":RegExReplace(Name, " - ((Windows|Microsoft) Internet Explorer|Internet Explorer)")
   for wb in ComObjCreate("Shell.Application").Windows
      if wb.LocationName=Name and InStr(wb.FullName, "iexplore.exe")
         return wb
}

;This is used for tooltips. Variable just needs to be defined with _TT at the end of the name in the GUI section that its used in.
WM_MOUSEMOVE()
{
    static CurrControl, PrevControl, _TT  ; _TT is kept blank for use by the ToolTip command below.
    CurrControl := A_GuiControl
    If (CurrControl <> PrevControl and not InStr(CurrControl, " "))
    {
        ToolTip  ; Turn off any previous tooltip.
        SetTimer, DisplayToolTip, 1000
        PrevControl := CurrControl
    }
    return

    DisplayToolTip:
    SetTimer, DisplayToolTip, Off
    ToolTip % %CurrControl%_TT  ; The leading percent sign tell it to use an expression.%
    SetTimer, RemoveToolTip, 3000
    return

    RemoveToolTip:
    SetTimer, RemoveToolTip, Off
    ToolTip
    return
}

;Loads the list of clients from the configfile excluding the settings and CWxLogins keys.
loadClientList(file="")
{
	clientList :=
	IniRead, rawClientList, %file%
	replacedClientList := RegExReplace(rawClientList, "`n", "|")
	Loop, parse, replacedClientList, "|"
	{
		If A_LoopField contains settings
			Continue
		Else
			clientList .= A_LoopField . "|"
	}
	Return clientList
}

;Checks to see if there is already a section in the configfile with the same trigger defined.  If so, spit back error to user.
duplicateTriggerCheck(trigger="")
{
	Loop, read, %configFile%
	{
		If A_LoopReadLine contains clientTrigger=
		{
			StringTrimLeft, triggerToCheck, A_LoopReadLine, 14
			If (trigger=triggerToCheck)
			{
				triggerError := "1"
				MsgBox, 16, Duplicate Trigger, This trigger is already set to a different client.`nPlease enter a different trigger.
				Break
			}
		}
		Else
			Continue
	}
	Return triggerError
}

;Checks when user tries to add a new Client to see if the Client Mnemonic (INI Section) already exists.
;If it does, spit out error. Otherwise they will be overwriting the existing credentials.
duplicateClientCheck(file="", newClientMnemonic="")
{
	IniRead, rawClientList, %file%
	replacedClientList := RegExReplace(rawClientList, "`n", "|")
	Loop, parse, replacedClientList, "|"
	{
		If A_LoopField contains %newClientMnemonic%
			{
				newClientError := "1"
				MsgBox, 16, Duplicate Client Mnemonic, This Client Mnemonic is already defined.`nPlease enter something different to identify these logins.
				Break
			}
		Else
			Continue
	}
	Return newClientError
}

;Retrieves info to populate LogMeIn GUI when a client is selected in the drop down list.  This fires any time the drop down selection changes.
retrieveClientInfo(client="", configFile="", ByRef URL=0, ByRef trigger=0, ByRef appUserName=0, ByRef password=0)
{
	IniRead, appUserName, %configFile%, %client%, appUserName
	IniRead, password, %configFile%, %client%, password
	IniRead, trigger, %configFile%, %client%, clientTrigger
	IniRead, URL, %configFile%, %client%, clientURL
	password := Crypt.Encrypt.StrDecrypt(password, "xtensible", 5, 1)
}

;Checks for various common edit IDs on webpages to enter the defined username and password.
scanAndEnterInfo(clientURL="", userID="", userPass="")
{
	userIDFields := "login-userid|login-account|email|user|userid|szEmail-login|username|login_username|accountName"
	passwordIDFields := "login-pass|login-password|password|pass|szPassword-login|login_password"
	buttonFound := 0
	IE := ComObjCreate("InternetExplorer.Application")
	IE.Visible := True
	WinMaximize, % "ahk_id " IE.HWND 			;%
	IE.Navigate(clientURL)
	While IE.busy or IE.ReadyState != 4
		Sleep 100
	Loop, parse, userIDFields, |
	{
		IE.document.getElementByID(A_LoopField).value := userID
	}
	Loop, parse, passwordIDFields, |
	{
		IE.document.getElementByID(A_LoopField).value := userPass
	}
	inputs := IE.document.getElementsByTagName("input")
	Loop % inputs.length ;%
	{
		index := A_Index - 1
		inputButtonValue := inputs[index].value
		If inputButtonValue in Log In,Sign In,Go,Login
		{
			inputs[index].click()
			buttonFound := 1
			Break
		}
	}
	If not(buttonFound)
	{
		buttons := IE.document.getElementsByTagName("button")
		Loop % buttons.length ;%
		{
			index := A_Index - 1
			buttonInnerHTML := buttons[index].innerHTML
			buttonValue := buttons[index].value
			If buttonInnerHTML contains Log In,Sign In,Go,Login
			{
				buttons[index].click()
				Break
			}
			Else If buttonValue in Log In,Sign In,Go,Login
			{
				buttons[index].click()
				Break
			}
		}
	}
}

;Bread and butter of the hotstrings. This creates the .ahk from the configfile, compiles it, and then deletes the .ahk file.
UpdateHotStringsExe()
{
	Global
TrayTip, LogMeIn, Updating your HotStrings...
	FileAppend,
	(
#Include Crypt.ahk
#Include CryptConst.ahk
#Include CryptFoos.ahk
#Persistent
#NoTrayIcon
#SingleInstance force

configFile := A_ProgramFiles . "\Hotkeys\LogMeIn\LogMeIn.ini"

sendCredentials(file="", trigger="", ByRef appUserName=0, ByRef password=0)
{
	Loop, read, `%file`%
	{
		If A_loopreadline contains clientTrigger=`%trigger`%
		{
			lineNumber := A_Index - 2
			FileReadLine, sectionNumberRaw, `%file`%, `%lineNumber`%
			sectionName := RegExReplace(sectionNumberRaw, ".*\[(.+?)\].*","$1")
			IniRead, appUserName, `%file`%, `%sectionName`%, appUserName
			IniRead, password, `%file`%, `%sectionName`%, password
			password := Crypt.Encrypt.StrDecrypt(password, "xtensible", 5, 1)
			Break
		}
		Else
			Continue
	}
	SendRaw, `%appUserName`%
	SendInput, {tab}
	SendRaw, `%password`%
	SendInput, {enter}
}`n`n
), %hotStringFile%

IniRead, wantReturnCheckBox, %configFile%, settings, wantReturn

Loop, read, %configFile%
{
	If A_LoopReadLine contains clientTrigger=
	{
		StringTrimLeft, triggerToWrite, A_LoopReadLine, 14
		If (WantReturnCheckBox="1")
		{
			FileAppend,
			(
::%triggerToWrite%::
sendCredentials(configFile, "%triggerToWrite%", appUserName, password)
Return`n`n
			), %hotStringFile%
		}
		Else If (WantReturnCheckBox="0")
		{
			FileAppend,
			(
:*:%triggerToWrite%::
sendCredentials(configFile, "%triggerToWrite%", appUserName, password)
Return`n`n
			), %hotStringFile%
		}
	}
	Else
		Continue
}
FileInstall, LogMeIn.ico, %dataDir%\LogMeIn.ico, 1
FileInstall, Crypt.ahk, %dataDir%\Crypt.ahk, 1
FileInstall, CryptConst.ahk, %dataDir%\CryptConst.ahk, 1
FileInstall, CryptFoos.ahk, %dataDir%\CryptFoos.ahk, 1
FileCreateDir, %compilerDir%
FileInstall, Compiler\Ahk2Exe.exe, %compilerDir%\Ahk2Exe.exe, 1
FileInstall, Compiler\ANSI 32-bit.bin, %compilerDir%\ANSI 32-bit.bin, 1
FileInstall, Compiler\AutoHotkeySC.bin, %compilerDir%\AutoHotkeySC.bin, 1
FileInstall, Compiler\Unicode 32-bit.bin, %compilerDir%\Unicode 32-bit.bin, 1
FileInstall, Compiler\Unicode 64-bit.bin, %compilerDir%\Unicode 64-bit.bin, 1
Process, Close, LogMeInHotStrings.exe
RunWait, "%compilerDir%\Ahk2Exe.exe" /in "%hotStringFile%" /icon "%dataDir%\LogMeIn.ico"
FileDelete, %dataDir%\LogMeIn.ico
FileDelete, %dataDir%\Crypt.ahk
FileDelete, %dataDir%\CryptConst.ahk
FileDelete, %dataDir%\CryptFoos.ahk
FileRemoveDir, %compilerDir%, 1
FileMove, %dataDir%\LogMeInHotStrings.exe, %hotkeyDir%, 1
Run, %hotStringExe%
TrayTip, LogMeIn, HotStrings updated successfully!
FileDelete, %hotStringFile%
TrayTip
}

;****************************************************************************************
; Main GUI
;****************************************************************************************

LogMeInHotkey:

SetTitleMatchMode, 3

;Various info that must be loaded each time the GUI is opened to populate fields.
clientListDDL := loadClientList(configFile) ;Load client drop down list from configFile
ReloadButton_TT := "Reload hotstrings"
GoToIcon_TT := "Navigate to App/Site"
WantReturnCheckBox_TT := "This option, when checked, will make your`nhotstrings only send upon hitting the enter, tab, or space key`ninstead of sending immediately after typing the trigger."

IniRead, wantReturn, %configFile%, settings, wantReturn
if(wantReturn = "ERROR")
	wantReturn := "1"

IfWinNotExist, LogMeIn (%version%)
{
	Gui, MainLogMeIn:-MaximizeBox
	Gui, MainLogMeIn:Color, Silver
	Gui, MainLogMeIn:Margin, 10, 10
	Gui, MainLogMeIn:Font, s10, Verdana
	Gui, MainLogMeIn:Add, Button, w50 h25 x+25 y+15 Default Section gGoToSite vGoToIcon, GO
	Gui, MainLogMeIn:Add, Text, w60 h20 x+20 ys Right, App/Site
	Gui, MainLogMeIn:Add, DropDownList, ys-2 w175 h50 r10 Center Sort gViewCredentials vClient Uppercase, %clientListDDL%
	Gui, MainLogMeIn:Add, Button, w50 h25 ys-2 gAddClientButton, Add
	Gui, MainLogMeIn:Add, Button, w75 h25 ys-2 gRemoveClientButton, Remove
	If A_IsCompiled
		Gui, MainLogMeIn:Add, Picture, x+20 ys+2 gReloadLogMeInHotstrings BackgroundTrans vReloadButton, %refreshIconPath%
	Gui, MainLogMeIn:Add, GroupBox, w475 h55 xs Center Section
	Gui, MainLogMeIn:Add, Text, h20 w75 xs+10 ys+20 Right Section, URL/Path
	Gui, MainLogMeIn:Add, Edit, hp w200 r1 ys-2 +ReadOnly -Background -TabStop +Center vclientURL
	Gui, MainLogMeIn:Add, Button, w150 h25 ys-2 gChangeURLPath vChangeURLPath, Change URL/Path
	Gui, MainLogMeIn:Add, GroupBox, w475 h55 xs-10 Center Section
	Gui, MainLogMeIn:Add, Text, h20 w75 xs+10 ys+20 Right Section, Trigger
	Gui, MainLogMeIn:Add, Edit, hp w200 r1 ys-2 +ReadOnly -Background -TabStop +Center vclientTrigger
	Gui, MainLogMeIn:Add, Button, w150 h25 ys-2 gChangeTrigger vChangeTrigger, Change Trigger
	Gui, MainLogMeIn:Add, GroupBox, w475 h90 xs-10 ys+40 Center section
	Gui, MainLogMeIn:Add, Text, h20 w75 xs+10 ys+20 Right section, User Name
	Gui, MainLogMeIn:Add, Edit, hp w200 r1 ys +ReadOnly -Background -TabStop +Center vappUserName
	Gui, MainLogMeIn:Add, Text, h20 w75 xs ys+35 Right section, Password
	Gui, MainLogMeIn:Add, Edit, hp w200 r1 ys +ReadOnly -Background -TabStop +Center +Password vPassword
	Gui, MainLogMeIn:Add, Button, w150 h25 ys-15 Section gChangeCredentials vChangeCredentials, Change Credentials
	Gui, MainLogMeIn:Font, s10 Bold, Verdana
	Gui, MainLogMeIn:Add, Text, Center xs-325 y+35 Section cBlue gChangeHotkey vChangeHotKey, Change Hotkey
	Gui, MainLogMeIn:Add, Checkbox, x+200 ys Checked%wantReturn% vWantReturnCheckBox gUpdateWantReturn, Require Enter/Tab/Space
	Gui, MainLogMeIn:Show, Center w550 h285, LogMeIn (%version%)
}
Else
{
	WinActivate, LogMeIn (%version%)
}

GuiControl, MainLogMeIn:Focus, Client

;Used for Tooltips.
OnMessage(0x200, "WM_MOUSEMOVE")
Return

GoToSite:
Gui, MainLogMeIn:Submit
Gui, MainLogMeIn:Destroy
If not(clientURL)
	Return
Else If clientURL contains www,http
	scanAndEnterInfo(clientURL, appUserName, Password)
Else If clientURL contains .exe
	Run, %clientURL%
Return

;Updates the WantReturn setting in the ini config file.
UpdateWantReturn:
Sleep 300
GuiControlGet, WantReturnCheckBox
Gui, MainLogMeIn:+Disabled
If (WantReturnCheckBox="1")
	IniWrite, 1, %configFile%, settings, wantReturn
Else If (WantReturnCheckBox="0")
	IniWrite, 0, %configFile%, settings, wantReturn
Sleep 300
UpdateHotStringsExe()
Gui, MainLogMeIn:-Disabled
If (WantReturnCheckBox="1")
	MsgBox, 64, LogMeIn (%version%), Your hotstrings now require the Tab/Enter/Space key to be pressed to send., 3
Else If (WantReturnCheckBox="0")
	MsgBox, 64, LogMeIn (%version%), Your hotstrings now send immediately after being typed., 3
Return

;Changes the hotkey associated with LogMeIn.
ChangeHotkey:
IniRead, CurrentHotkey, %configFile%, settings, LogMeInTrigger
If(CurrentHotkey = "ERROR")
	CurrentHotkey := "#F2"
Gui, ChangeHotkey:-MaximizeBox +OwnerMainLogMeIn
Gui, ChangeHotkey:Color, Silver
Gui, ChangeHotkey:Font, s8, Verdana
Gui, ChangeHotkey:Margin, 15,15
Gui, ChangeHotkey:Add, Text, x+15 y+15 Center Section, Your current hotkey is: 
Gui, ChangeHotkey:Add, Edit, ys w45 r1 +ReadOnly -Background Center vCurrentHotkey, %CurrentHotkey%
Gui, ChangeHotkey:Font, s6, Verdana
Gui, ChangeHotkey:Add, Text, xsCenter Section, (HINT: # = Windows key, ^ = Ctrl key, ! = Alt key)
Gui, ChangeHotkey:Font, s8, Verdana
Gui, ChangeHotkey:Add, Text, xs Center Section, Enter your new hotkey: 
Gui, ChangeHotkey:Add, Edit, ys w45 r1 Center vLogMeInNewHotkey
Gui, ChangeHotkey:Add, Button, Default w60 h25 xs+80 gSubmitLogMeInHotkey, Submit
Gui, ChangeHotkey:Show, Center w275 h150, Change Hotkey - LogMeIn
Return

;Confirms settings prior to changing the hotkey.
SubmitLogMeInHotkey:
Gui, ChangeHotkey:Submit, NoHide
If (LogMeInNewHotkey="")
{
	MsgBox, 16, LogMeIn (%version%), You must enter a hotkey to change to first.
	Return
}
Gui, ChangeHotkey:Destroy
WinActivate, LogMeIn (%version%)
Hotkey, %LogMeInHotkeyTrigger%, LogMeInHotkey, Off
IniWrite, %LogMeInNewHotkey%, %configFile%, settings, LogMeInTrigger
LogMeInHotkeyTrigger := LogMeInNewHotkey
Hotkey, %LogMeInHotkeyTrigger%, LogMeInHotkey, On
Return

ChangeURLPath:
Process, Close, LogMeInHotStrings.exe
Gui, MainLogMeIn:Submit, NoHide
GuiControlGet, Client
If (Client = "")
{
	MsgBox, 16, Change URL/Path, Select an App/Site first., 4
	GuiControl, MainLogMeIn:Focus, Client
	Return
}
Gui, URLChange: -MaximizeBox +OwnerMainLogMeIn
Gui, URLChange:Color, Silver
Gui, URLChange:Font, s10, Verdana
Gui, URLChange:Margin, 10, 10
Gui, URLChange:Add, Text, h20 Section, New URL/Path
Gui, URLChange:Add, Edit, h25 w200 ys +Center vNewURL
Gui, URLChange:Add, Button, Default w100 h30 xs+150 ys+40 gSetURL vSetURL, Set URL/Path
Gui, URLChange:Show, Center AutoSize, Change URL/Path
Return

SetURL:
Gui, URLChange:Submit, NoHide
NewURL := trimWhitespacePerLine(NewURL)
Gui, URLChange:Destroy
IniWrite, %NewURL%, %configFile%, %Client%, clientURL
GuiControl, MainLogMeIn:, clientURL, %NewURL%
Return

;GUI to change the trigger on the selected item in the drop down list.
ChangeTrigger:
Process, Close, LogMeInHotStrings.exe
Gui, MainLogMeIn:Submit, NoHide
GuiControlGet, Client
If (Client = "")
{
	MsgBox, 16, Change Trigger, Select an App/Site first., 4
	GuiControl, MainLogMeIn:Focus, Client
	Return
}
Gui, TriggerChange: -MaximizeBox +OwnerMainLogMeIn
Gui, TriggerChange:Color, Silver
Gui, TriggerChange:Font, s10, Verdana
Gui, TriggerChange:Margin, 10, 10
Gui, TriggerChange:Add, Text, h20 Section, New trigger
Gui, TriggerChange:Add, Edit, h25 w200 ys +Center vNewTrigger
Gui, TriggerChange:Add, Button, Default w100 h30 xs+150 ys+40 gSetTrigger vSetTrigger, Set Trigger
Gui, TriggerChange:Show, Center AutoSize, Change Trigger
Return

;Checks to confirm the trigger is not a duplicate and then set it if not.
SetTrigger:
Gui, TriggerChange:Submit, NoHide
triggerError := duplicateTriggerCheck(NewTrigger)
If (triggerError)
{
	GuiControl, TriggerChange:, NewTrigger,
	triggerError :=
	Return
}
NewTrigger := trimWhitespacePerLine(NewTrigger)
Gui, TriggerChange:Destroy
IniWrite, %NewTrigger%, %configFile%, %Client%, clientTrigger
GuiControl, MainLogMeIn:, clientTrigger, %NewTrigger%
Gui, MainLogMeIn:+Disabled
UpdateHotStringsExe()
Gui, MainLogMeIn:-Disabled
Return

;Fires whenever the selection in the drop down list changes.
ViewCredentials:
Gui, MainLogMeIn:Submit, NoHide
retrieveClientInfo(Client, configFile, URL, trigger, appUserName, password)
GuiControl, MainLogMeIn:, clientURL, %URL%
GuiControl, MainLogMeIn:, clientTrigger, %trigger%
GuiControl, MainLogMeIn:, appUserName, %appUserName%
GuiControl, MainLogMeIn:, Password, %password%
Password_TT = %password%
clientURL_TT = %URL%
Return

;Removes client credentials and trigger.
RemoveClientButton:
Gui, MainLogMeIn:Submit, NoHide
If (Client = "")
{
	MsgBox, 16, Remove client, Select an App/Site first., 4
	GuiControl, MainLogMeIn:Focus, Client
	Return
}
clientListDDL := ;Empty the drop down list first.
GuiControlGet, Client
GuiControlGet, clientTrigger
Gui, MainLogMeIn:+Disabled
IniDelete, %configFile%, %Client%
clientListDDL .= "|" . loadClientList(configFile) ;Reload the drop down list once the ini changes are made.  "|" in front makes the GuiControl command overwrite the existing variable in the dropdown list.
GuiControl, MainLogMeIn:, Client, %clientListDDL%
GuiControl, MainLogMeIn:, clientURL,
GuiControl, MainLogMeIn:, clientTrigger,
GuiControl, MainLogMeIn:, appUserName,
GuiControl, MainLogMeIn:, password,
UpdateHotStringsExe()
Gui, MainLogMeIn:-Disabled
Return

;Opens GUI to add a new client.
AddClientButton:
Process, Close, LogMeInHotStrings.exe
Gui, AddClient:-MaximizeBox +OwnerMainLogMeIn
Gui, AddClient:Color, Silver
Gui, AddClient:Margin, 10, 10
Gui, AddClient:Font, s10, Verdana
Gui, AddClient:Add, Text, w100 Right, New App/Site
Gui, AddClient:Add, Edit, w200 ys r1 +Uppercase +Center vNewClientMnemonic
Gui, AddClient:Add, Text, w100 xs Right Section, URL/Path
Gui, AddClient:Add, Edit, w200 ys r1 +Center vNewURL
Gui, AddClient:Add, Text, w100 xs Right Section, Trigger
Gui, AddClient:Add, Edit, w200 ys r1 +Center vNewTrigger
Gui, AddClient:Add, Text, w100 xs Right Section, User Name
Gui, AddClient:Add, Edit, w200 ys r1 +Center vNewClientappUserName,
Gui, AddClient:Add, Text, w100 xs Right Section, Password
Gui, AddClient:Add, Edit, w200 ys r1 +Center vNewClientPassword
Gui, AddClient:Add, Button, Default w150 h25 xs+100 gAddClient vAddCredentials, Add Credentials
Gui, AddClient:Show, Center Autosize, Add New App/Site Credentials
Return

;Confirms all parameters are set, checks for duplicate triggers, prior to adding client.
AddClient:
clientListDDL := ;Clear drop down list variable.
Gui, AddClient:Submit, NoHide
GuiControlGet, WantReturnCheckBox, MainLogMeIn:
If (NewClientMnemonic="" or NewTrigger="" or NewClientappUserName="" or NewClientPassword="")
{
	MsgBox, 16, Uh-Oh, All fields are required., 3
	If not(NewClientMnemonic)
		GuiControl, AddClient:Focus, NewClientMnemonic
	Else If not(NewURL)
		GuiControl, AddClient:Focus, NewURL
	Else If not(NewTrigger)
		GuiControl, AddClient:Focus, NewTrigger
	Else If not(NewClientappUserName)
		GuiControl, AddClient:Focus, NewClientappUserName
	Else If not(NewClientPassword)
		GuiControl, AddClient:Focus, NewClientPassword
	Return
}
newClientError := duplicateClientCheck(configFile, NewClientMnemonic) ;Check if the client mnemonic entered is already in use.  If so, toss out an error msgbox and the value 1.
If (newClientError) ;If duplicate client mnemonic, wipe out the client mnemonic they entered and set focus to that control.
{
	GuiControl, AddClient:, NewClientMnemonic,
	GuiControl, AddClient:Focus, NewClientMnemonic
	triggerError =
	Return
}
triggerError := duplicateTriggerCheck(NewTrigger) ;Check if the trigger entered is already in use.  If so, toss out an error msgbox and the value 1.
If (triggerError) ;If duplicate trigger, wipe out the trigger they entered and set focus to that control.
{
	GuiControl, AddClient:, NewTrigger,
	GuiControl, AddClient:Focus, NewTrigger
	triggerError =
	Return
}
NewClientPassword := Crypt.Encrypt.StrEncrypt(NewClientPassword, "xtensible", 5, 1)
IniWrite, %NewURL%, %configFile%, %NewClientMnemonic%, clientURL
IniWrite, %NewTrigger%, %configFile%, %NewClientMnemonic%, clientTrigger
IniWrite, %NewClientappUserName%, %configFile%, %NewClientMnemonic%, appUserName
IniWrite, %NewClientPassword%, %configFile%, %NewClientMnemonic%, password
Gui, AddClient:Destroy
clientListDDL .= "|" . loadClientList(configFile)
GuiControl, MainLogMeIn:, client, %clientListDDL%
GuiControl, MainLogMeIn:, clientURL,
GuiControl, MainLogMeIn:, clientTrigger,
GuiControl, MainLogMeIn:, appUserName,
GuiControl, MainLogMeIn:, password,
Gui, MainLogMeIn:+Disabled
UpdateHotStringsExe()
Gui, MainLogMeIn:-Disabled
Return

;GUI to change appUserName and password for selected client in drop down list.
ChangeCredentials:
Gui, MainLogMeIn:Submit, NoHide
If (Client = "")
{
	MsgBox, 16, Change Credentials, Select an App/Site first., 4
	GuiControl, MainLogMeIn:Focus, Client
	Return
}
Gui, ChangeCreds:-MaximizeBox +OwnerMainLogMeIn
Gui, ChangeCreds:Color, Silver
Gui, ChangeCreds:Margin, 10, 10
Gui, ChangeCreds:Font, s10, Verdana
Gui, ChangeCreds:Add, Text, w125 Right Section, New User Name
Gui, ChangeCreds:Add, Edit, w200 ys r1 +Center vNewClientappUserName, %appUserName%
Gui, ChangeCreds:Add, Text, w125 xs Right Section, New Password
Gui, ChangeCreds:Add, Edit, w200 ys r1 +Center vnewClientPassword
Gui, ChangeCreds:Add, Button, Default w150 h25 xs+160 ys+40 gUpdateCredentials vUpdateCredentials, Update Credentials
Gui, ChangeCreds:Show, Center Autosize, Update Client Credentials
Return

;Only updates the configFile as the hotstrings exe pulls the logins from that so no update to it is necessary.
UpdateCredentials:
Gui, ChangeCreds:Submit, NoHide
Gui, ChangeCreds:Destroy
GuiControl, MainLogMeIn:, Password, %newClientPassword%
Password_TT = %newClientPassword%
GuiControl, MainLogMeIn:, appUserName, %newClientappUserName%
newEncryptedPassword := Crypt.Encrypt.StrEncrypt(newClientPassword, "xtensible", 5, 1)
IniWrite, %newEncryptedPassword%, %configFile%, %Client%, password
IniWrite, %newClientappUserName%, %configFile%, %Client%, appUserName
MsgBox, 64, Update Credentials, Your information for %Client% has been updated successfully., 4
Return

reloadLogMeInHotstrings:
UpdateHotStringsExe()
Return

reloadLogMeIn:
Run, %exePath%
ExitApp

SuspendHotStrings:
Menu, Tray, ToggleCheck, Suspend LogMeIn HotStrings
Process, Exist, LogMeInHotStrings.exe
If (ErrorLevel)
	Process, Close, LogMeInHotStrings.exe
Else
	Run, %hotStringExe%
Return

ExitLogMeIn:
Process, Close, LogMeInHotStrings.exe
ExitApp

MainLogMeInGuiClose:
MainLogMeInGuiEscape:
Gui, MainLogMeIn:Destroy
Exit

ChangeHotkeyGuiClose:
ChangeHotkeyGuiEscape:
Gui, ChangeHotkey:Destroy
Exit

;We have to run the hotStringExe when either of the following two GUIs are closed by escape or the X
;since we close the process to avoid the hotstrings from triggering while making changes in those windows.
AddClientGuiClose:
AddClientGuiEscape:
Run, %hotStringExe%
Gui, AddClient:Destroy
Exit

URLChangeGuiClose:
URLChangeGuiEscape:
Run, %hotStringExe%
Gui, URLChange:Destroy
Exit

TriggerChangeGuiClose:
TriggerChangeGuiEscape:
Run, %hotStringExe%
Gui, TriggerChange:Destroy
Exit

ChangeCredsGuiClose:
ChangeCredsGuiEscape:
Gui, ChangeCreds:Destroy
Exit 

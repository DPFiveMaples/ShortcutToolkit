﻿/*;========================================================
;==  INI Values (DO NOT ADJUST THE LINE SPACING!!!)
;==========================================================
[INI_Section]
version=27


*/
;==========================================================
;==  Boilerplate
;==========================================================

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases, prevents checking empty variables
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Persistent
#SINGLEINSTANCE force
MouseSpeed = 30 ; This is a slowing modifier - the higher it is, the slower the down button goes, default is 15
MouseSUp = 15
MouseSDown = 30
CurrentVer = 0
NewVer = 0
;SetBatchLines, -1	; Runs the script at max speed, default is 10 or 20 ms

;==========================================================
;==  GENERAL NOTES
;==========================================================

; Key	Syntax
; Alt       !
; Ctrl      ^
; Shift     +
; Win Logo  #



;===========================================================
;==  Update Module
;===========================================================
; Files if hosted on Github    : https://raw.githubusercontent.com/MarvinFiveMaples/ShortcutToolkit/master/ShortcutToolkit.ahk?raw=true

SetTimer UpdateCheck, 60000 ; Check each minute
;Return

ButtonRestartToolkit:
^+#r:: ;c 🌟 Restart Marvin's ShortCutToolkit ⌨️ Ctrl+Shift+Win+r | Pressing Ctrl+Shift+Win+r will restart the ShortCutToolkit - use this if it freezes up on you.
{
	Progress, w250,,, Hold yer ponies,  I'm restarting…
	vRestart := 0
	loop, 45
	{
        Progress, %vRestart%
        vRestart := vRestart + 10
        sleep, 1
	}
	Progress,  Off
	MsgBox, 0, , Rebooted - please click 'OK' to proceed., 1
	Reload
	ExitApp
}


UpdateCheck:
If (A_Hour = 01 And A_Min = 12)
{
	Progress, w250,,, Hold yer ponies,  I'm updating…
	Sleep, 2000
	Progress, off
	gosub VersionCheck
	}
Return

; The following is for testing ONLY, and shouldn't be used otherwise.
/*
SetTimer UpdateCheckTest, 1000 ; Check each second
Return

UpdateCheckTest:
If (A_Hour = 13)
	{
	Msgbox, Hey!
	Progress, w250,,, Hold yer ponies,  I'm updating…
	Sleep, 10000
	Progress, off
	gosub VersionCheck
	}
Return
*/



VersionCheck:
^+#t:: ; c If you didn't trigger this window by pressing 'Win+F2', then it must mean that you had an update last night! Yay! (you can click OK)
{
; IniRead, OutputVar, C:\Temp\myfile.ini, section2, key
; MsgBox, The value is %OutputVar%.
; FileReadLine, %CurrentVer%, ShortcutToolkit.ahk, 5
; FileReadLine, %NewVer%, //raw.githubusercontent.com/MarvinFiveMaples/ShortcutToolkit/master/ShortcutToolkit.ahk?raw=true, 5
IniRead, CurrentVer, ShortcutToolkit.ahk, INI_Section, version
UrlDownloadToFile, https://raw.githubusercontent.com/MarvinFiveMaples/ShortcutToolkit/master/ShortcutToolkit.ahk?raw=true, JunkKit.ahk ;*[ShortcutToolkit]
IniRead, NewVer, JunkKit.ahk, INI_Section, version
FileDelete, JunkKit.ahk
if (CurrentVer < NewVer)
	{
	gosub UpdateScript ; Insert the following above if you need to check anything in the future: MsgBox, %CurrentVer% & %NewVer%
	Return
	}
Return
}

ButtonForceUpdateToolkit:
UpdateScript:
^+#u:: ;c 🌟 Update Script ⌨️ Ctrl+Shift+Win+u | Typing Ctrl+Shift+Win+u will trigger an update of the script - also automatically triggered every morning at 1:15am
{
	UrlDownloadToFile, https://raw.githubusercontent.com/MarvinFiveMaples/ShortcutToolkit/master/ShortcutToolkit.ahk?raw=true, ShortcutToolkit.ahk ;*[ShortcutToolkit]
	;Progress, w250,,, Hold yer ponies,  I'm updating…
	MsgBox If you see me, I either just updated when you triggered me to, or I updated last night. Either way, please click 'OK', and go about your day! Also, press 'Win+F2' to open up a quick help cheat-sheet.
	Reload
	ExitApp
}


; Stub for future work:
; https://www.autohotkey.com/download/AutoCorrect.ahk

^+#a::
{
    Run AutoCorrect.ahk
    return
}


;==========================================================
;==  Help File Section
;==========================================================

ButtonHelp:
#F2:: ;c 🌟 Activate "Help" Menu ⌨️ WinKey+F2 | will bring up this help file, which attempts to automatically document these functions. It may go without saying, but I'm still working on it. :D
{
fileread content, %a_scriptfullpath%
comment=
loop parse, content,`n
{
 ifinstring a_loopfield,;c  ;;
  ifnotinstring a_loopfield,;;
 {
  position := InStr(a_loopfield,";c") ;;
  position +=1
  stringtrimleft com,a_loopfield,%position%
  comment =%comment%%com%`n`n
 }
}
  gui, submit
  msgbox %comment%
Return
}

;==========================================================
;==  Useful Hotkeys Section
;==========================================================


::Jam::James Chase`t8023875157`t110 ;c 🌟 Populate Warehouse Name & Number ⌨️ Type "jam" & Hit Tab OR just press the Warehouse NDC SCF Info button | This will auto-populate the warehouse info on the NDC/SCF 8125 screen
; :*:jam`t::James Chase`t8023875157`t110 ; Old - replaced by above



ButtonWarehouseNDCSCFInfo:
{
Gui, Submit
SendInput James Chase`t8023875157`t110
Return
}

:*:shruggie`t::¯\_(ツ)_/¯ ;c 🌟 Insert ¯\_(ツ)_/¯ into your text ⌨️ Type "shruggie" & Hit Tab OR just press the button

Button¯\_(ツ)_/¯: ; This does the same thing as the above command, but FROM A BUTTON!!!
{
    gui, submit
    sleep, 69
    Send, ¯\_(ツ)_/¯
    Return
}


^!'::’ ; Left in for compatibility with earlier versions of the program.
^+'::’ ;c 🌟 Dartmouth Apostrophe ⌨️ Ctrl+Shift+' | Ctrl+Shift+' (or just pressing the "Dartmouth Apostrophe" button) will insert a Dartmouth apostrophe, instead of a regular one.

ButtonDartmouthApostrophe:
{
    gui, submit
    sleep, 69
    Send, ’
    Return
}


ButtonLaunchWiki:
^!w:: ;c 🌟 Open Wiki to Index ⌨️ Ctrl+Alt+W | Ctrl+Alt+W (or just pressing the "Launch Wiki" button) will open the Five Maples Wiki in a new tab to the Index of All Pages
{
    gui, submit
    Run, http://watson/mediawiki/index.php/Special:AllPages
    Return
}

^F1::gosub, ButtonLogin ; Note: "Help" (Win+F2) info included under primary label ButtonLogin.


^!k::
{
test := A_UserName
If test = "kendrak"
	{
		Msgbox, %test%
		return
	}
Else
	{
	MsgBox, Kendra is Kewl!
	return
	}
}

;===========================================================
;==  Exit/Escape Section
;===========================================================


#IfWinActive ahk_class AutoHotkeyGUI

;/* ; Fix below with context sensitive : IfWinActive ahk_class AutoHotkeyGUI ; or something of the like
GuiEscape:
GuiClose:
CloseAllWindows: ; Close all and reset the GUI number
;^space::
{
	;Reload
	;ExitApp
    IfWinExist DP Dashboard
    {
        Gui DPDashboard:Cancel
        Gui, submit
    }
	Return
}

#IfWinActive

;===========================================================
;==  DP Dashboard
;===========================================================


^space:: ;c 🌟 DP Dashboard ⌨️ Ctrl+Space | Ctrl+Space will launch what I like to call the "DP Dashboard", which contains most of the clickable commands
{
    IfWinNotExist DP Dashboard
	{
        Gui, New,,DP Dashboard
        Gui Add, Button, x13 y19 w80 h23, &Std Dedupe
        Gui Add, Button, x13 y71 w80 h23, &Clipboard Manager
        Gui Add, Button, x13 y45 w80 h23, &Just Dupe Indexes
        Gui Add, Button, x13 y96 w80 h23, &Login
        Gui Add, Button, x13 y122 w80 h23, &Warehouse NDC SCF Info
        Gui Add, Button, x13 y150 w80 h23, &Launch Wiki
        Gui Add, Button, x107 y19 w80 h23, &Restart Toolkit
        Gui Add, Button, x107 y150 w80 h23, &Help
        Gui Add, Button, x107 y71 w80 h23, ¯\&_(ツ)_/¯
        Gui Add, Button, x107 y96 w80 h23, &Dartmouth Apostrophe
        Gui Add, Button, x107 y122 w80 h23, &Thank You Emailer
        Gui Add, Button, x107 y45 w80 h23, &Force Update Toolkit
        Gui, Show,,DP Dashboard
        Return
    }
    Else
    {
        Gui DPDashboard:Cancel
        Gui, submit
        Return
    }
    return
}



;===========================================================
;==  Clipboard Management
;===========================================================

TestClipCapture:
~^x::
~^c::
{
sleep, 100
#ClipboardTimeout 3000
if clipboard = ""
{
	; it's a picture, file etc.
	return
}
if clipboard = %clpb1%
{
	; already exists, stay in slot 1
}
else
{
	; new content
	clpb5 := clpb4
	clpb4 := clpb3
	clpb3 := clpb2
	clpb2 := clpb1
	clpb1 := clipboard
}
; finally
clpb1len := StrLen(clpb1)
StringTrimRight, clpbs1, clpb1, clpb1len - 50
clpb2len := StrLen(clpb2)
StringTrimRight, clpbs2, clpb2, clpb2len - 50
clpb3len := StrLen(clpb3)
StringTrimRight, clpbs3, clpb3, clpb3len - 50
clpb4len := StrLen(clpb4)
StringTrimRight, clpbs4, clpb4, clpb4len - 50
clpb5len := StrLen(clpb5)
StringTrimRight, clpbs5, clpb5, clpb5len - 50
; tooltip, 1: %clpbs1%`n2: %clpbs2%`n3: %clpbs3%`n4: %clpbs4%`n5: %clpbs5% ; Uncomment this to get a "popup" of what was copied� every time someone copies.
sleep, 300
; tooltip Uncomment if the above tooltip is also uncommented
#ClipboardTimeout 1000
return
}

ButtonClipboardManager:
^+v:: ;c 🌟 Launch Clipboard Manager ⌨️ Ctrl+Shift+V OR just press the Clipboard Manager button | This will let you select one of the last 5 things you've copied, assuming they were text.`nIt will time out and default back to your current clipboard after 3 seconds.
{
	;varProgressMeter := 0
	clpb1len := StrLen(clpb1)
	StringTrimRight, clpbs1, clpb1, clpb1len - 50
	clpb2len := StrLen(clpb2)
	StringTrimRight, clpbs2, clpb2, clpb2len - 50
	clpb3len := StrLen(clpb3)
	StringTrimRight, clpbs3, clpb3, clpb3len - 50
	clpb4len := StrLen(clpb4)
	StringTrimRight, clpbs4, clpb4, clpb4len - 50
	clpb5len := StrLen(clpb5)
	StringTrimRight, clpbs5, clpb5, clpb5len - 50
	; inputbox, clpbNbr, which clipboard number do you want to have in your clipboard now?, 1: %clpbs1% 2: %clpbs2%`n3: %clpbs3% 4: %clpbs4%`n5: %clpbs5%`n	; In place originally, replaced with my GUI below
	Gui, New,,Choose your Clipboard
	Gui, Add, Radio,vclpbNbr Checked,Clipboard #&1: %clpbs1%
	Gui, Add, Radio,,Clipboard #&2: %clpbs2%
	Gui, Add, Radio,,Clipboard #&3: %clpbs3%
	Gui, Add, Radio,,Clipboard #&4: %clpbs4%
	Gui, Add, Radio,,Clipboard #&5: %clpbs5%
	Gui, Add, Text,w300, You will auto-select #1 after 3 seconds,Please select desired clipboard content and click "go".
	Gui, Add, Progress, w300 h20 cBlue vVarProgressMeter
	Gui, Add, Button,default ym,&Go
	Gui, Show,,Choose your Clipboard

    ; ==== Start Test Code
    ; This is ONLY used during debugging, but please don't touch it if you don't need to.

    if DebugVar = 1
    {
        WinActivate, Choose your Clipboard
        SendInput, 2{ENTER}
        gosub, ButtonGo
        return
    }
    ; ==== End Test Code

	Loop{
			GuiControl,, varProgressMeter, +1
			sleep, 30
		} Until A_Index > 99
	gosub, ButtonGo
	return


	ButtonGo:
		{
			gui, submit
			;if errorlevel = 1  ; old code, not sure what this was for or how to use it
			;	return			; old code, not sure what this was for or how to use it
			clipboard := clpb%clpbNbr%
			; Sleep, 1000
			; MsgBox, %clpbNbr% Test
			return
		}
Return
}


;===========================================================
;==  Postal One! Login Button (label: ButtonLogin)
;===========================================================

ButtonLogin:
{
    Gui, Submit
    User := ""
    IniRead,UnparsedCreds,\\Watson\Production\X\DP Use\Shared AHKs\Toolkit.ini, PostalOneLogins, %A_Username%, "ERROR|ERROR"  ; Consider making the \\Watson\Prod... path a path to X:\DP Use\etc
    UserCreds := Object("User")
    Loop, Parse, UnparsedCreds,|
    {
        If A_Index = 1
        {
        User := A_LoopField
        }
        Else If A_Index = 2
        {
        Pass := A_LoopField
        }
        Else
        {
        MsgBox, Error! Danger Will Robinson! See Marvin, and Mention ButtonLogin Loop
        }
    }



;}    ; <---Comment out the bracket once I'm working on this again
;/* ; <---Comment out the slashasterisk once I'm working on this again

; OK, this WHOLE THING should be triggered by either Kendra's hotkey, or pressing "Scroll Lock" when in the Postal One uploader.

MDClientVar := ""      ; MDRVar := ""
If FileExist("C:\Postal1\run-mdclient.bat")
{
MDClientVar := "C:\Postal1\"
}
Else If FileExist("C:\Postal1\MDRClient-win64-PROD\run-mdclient.bat")
{
        MDClientVar := "C:\Postal1\MDRClient-win64-PROD\"
    }
    Else If FileExist("C:\Program Files (x86)\Postal1\MDRClient-win64-PROD\run-mdclient.bat")
    {
        MDClientVar := "C:\Program Files (x86)\Postal1\MDRClient-win64-PROD\"
    }
    Else If FileExist("C:\Program Files\Postal1\MDRClient-win64-PROD\run-mdclient.bat")
    {
        MDClientVar := "C:\Program Files\Postal1\MDRClient-win64-PROD\"
    }
    Else If FileExist("C:\Program Files (x86)\MDRClient-win64-PROD\run-mdclient.bat")
    {
        MDClientVar := "C:\Program Files (x86)\MDRClient-win64-PROD\"
    }
    Else
    {
        MDClientVar := "ERROR!!!"
        MsgBox, Please see Marvin - something went wrong. Please screenshot and send him the next screen after this one.
        Msgbox, Error Data: %A_Username% | %MDClientVar%
        Reload
        Return
    }
    If MDClientVar <> ""
    {

        ;Msgbox, %MDClientVar%run-mdclient.bat  ; For testing only - uncomment out next three lines if testing
        ;Reload
        ;return


        Run, %MDClientVar%run-mdclient.bat, %MDClientVar%, Max
        SplashTextOn, 400, 100, Wait for it…, Please Left Click the "Username" field once the Uploader is open and ready for input then hit "Scroll Lock" to have your credentials entered and login.
        KeyWait, ScrollLock, D
        Sleep, 100
        SplashTextOff
        Sleep, 100
        Send, {BackSpace 3}%User%
        Sleep, 100
        Send, {Tab}
        Sleep, 100
        Send, %Pass%{Enter}
        WinWait, PostalOne! Mail.dat Client Application
        WinWaitClose
    }
    Else
    {
        MsgBox, Please tell Marvin - something MAY not be quite right with your Mail.Dat Client Install.
    }
    Sleep, 200


    Run, chrome.exe --incognito https://gateway.usps.com/eAdmin/view/signin
    Sleep 5000
    Send, {ShiftDown}{Tab}{Tab}{ShiftUp}%User%{TAB}%Pass%{ENTER}
    Return
}
    ;MsgBox, Hopefull this doesn't popup until I'm ready for it...
    ;MsgBox, This is under construction. At the moment, it just launches the login screen in incogneto mode.
	;Run, C:\Program Files (x86)\Google\Chrome\Application\chrome.exe -incognito https://gateway.usps.com/eAdmin/view/signin
    ;IEPointer := IEGet("usps")
    ;SplashTextOn, 400, 100, Wait for it…, Please press "Space" once the page is done loading.
    ;KeyWait, Space, D
    ;SplashTextOff
    ;Send, {ShiftDown}{Tab}{Tab}{ShiftUp}%User%
    ;Sleep, 100
    ;Send, {Tab}%Pass%{Enter}
    ;;MsgBox, %User% | %Pass%
	;Return
;}



;used to be WBGet renamed though
IEGet(WinTitle="ahk_class IEFrame", Svr#=1)
{ ; based on ComObjQuery docs
   static   msg := DllCall("RegisterWindowMessage", "str", "WM_HTML_GETOBJECT")
   ,   IID := "{0002DF05-0000-0000-C000-000000000046}" ; IID_IWebBrowserApp
;   ,   IID := "{332C4427-26CB-11D0-B483-00C04FD90119}" ; IID_IHTMLWindow2
   SendMessage msg, 0, 0, Internet Explorer_Server%Svr#%, %WinTitle%
   if (ErrorLevel != "FAIL") {
      lResult:=ErrorLevel, VarSetCapacity(GUID,16,0)
      if DllCall("ole32\CLSIDFromString", "wstr","{332C4425-26CB-11D0-B483-00C04FD90119}", "ptr",&GUID) >= 0 {
         DllCall("oleacc\ObjectFromLresult", "ptr",lResult, "ptr",&GUID, "ptr",0, "ptr*",pdoc)
         return ComObj(9,ComObjQuery(pdoc,IID,IID),1), ObjRelease(pdoc)
      }
   }
}

/*
ButtonLogin:
{
	Gui, Submit
	IfNotExist, %A_MyDocuments%\FMAHK.config
	{
		GetUnP("USPSLogin")
		FileRead, vP1Blob, %A_MyDocuments%\FMAHK.config
	}

	While(vP1Blob ="")
	{
		GetUnP("USPSLogin")
		FileRead, vP1Blob, %A_MyDocuments%\FMAHK.config
	}

	FileRead, vP1Blob, %A_MyDocuments%\FMAHK.config

	Loop
	{
	 Loop, read, %A_MyDocuments%\FMAHK.config
		  last_line := USPSLogin|_|_|
	 if InStr(last_line,"login:")
		  break
	}
MsgBox Found!

	MsgBox, %vP1Blob%


	Return
}
*/


GetUnP(vPurposeKey) ; vPurposeKey will be used to differentiate between the different records in the DB.
{
	While(vP1user ="")
		{
		InputBox, vP1user, Username Entry, Enter your Postal One Username,,,,,,,,Enter your PostalOne! Username
		}
	While(vP1pass ="")
		{
		InputBox, vP1pass, Password Entry, Enter your Postal One Password,,,,,,,,Enter your PostalOne! Password
		}
	FileDelete, %A_MyDocuments%\FMAHK.config
	FileAppend, %vPurposeKey%|_|_|%vP1User%|||||%vP1Pass%, %A_MyDocuments%\FMAHK.config ; If I need new lines for some reason, add with this: `n
	Return
}




;===========================================================
;==  Warehouse Department - Auto "Thank you" Emailer
;===========================================================

ButtonThankYouEmailer:
#+^m:: ;c 🌟 Warehouse Macro ⌨️ Ctrl+Shift+Win+m | Ctrl+Shift+Win+m will trigger the Warehouse Macro - this will launch the Google Sheet and (after a couple of clicks) send an email to Heather.
{
    Gui, Submit
	Run, C:\Program Files (x86)\Google\Chrome\Application\chrome.exe https://docs.google.com/spreadsheets/d/1l35If337LGq5pjDBIJ9lBj_UkVYxzVGyzLtMkGcwIqk/edit#gid=0
	Sleep,  50
	Progress, zh0 fs18, Please click the cell containing the Enter Date of the row `n that you wish to email about and then press 'Pause' to proceed. `n This notice will remain until you do. :D
	KeyWait, PAUSE, D
	Progress, Off
	settitlematchmode 2
	winactivate, Google Chrome
	Sleep, 50
	Send, ^c
	sleep,  50

	If clipboard != %A_MM%/%A_DD%/%A_YYYY% ; If they are different...
	{
		MsgBox, 4, Proceed with Wrong Date?, %clipboard% FAILS to match today's date (%A_MM%/%A_DD%/%A_YYYY%), 10
		IfMsgBox, No
			Return  ; User pressed the "No" button.
		IfMsgBox, Timeout
			Return ; i.e. Assume "No" if it timed out.
		; Otherwise, continue:
	}
	winactivate, Google Chrome
	; =============== This section should be refactored into an array-object loop].
	sleep,  50
	EnterDate := clipboard
	Send, {TAB}
	sleep,  20
	Send, ^c
	sleep,  50
	ClientName := clipboard
	Send, {TAB}
	sleep,  20
	Send, ^c
	sleep,  50
	Package := clipboard
	Send, {TAB}
	sleep,  20
	Send, ^c
	sleep,  50
	FileName := clipboard
	Send, {TAB}
	sleep,  20
	Send, ^c
	sleep,  50
	MailQty := clipboard
	Send, {TAB}
	sleep,  20
	Send, ^c
	sleep,  50
	MailDate := clipboard
	Send, {TAB}
	sleep,  20
	Send, ^c
	sleep,  50
	Salutation := clipboard
	;Send, {TAB}
	sleep,  20
	;Send, ^c
	;sleep,  50
	; RM ClientEmail := clipboard
	;Send, {TAB}
	;sleep,  20
	;Send, ^c
	;sleep,  50
	; =============== End of the section that should be refactored into an array-object loop.
	; Variables Used: %MsgBox%, %EnterDate%,  %ClientName% %Package% %FileName% %MailQty% %MailDate% %Salutation% %ClientEmail%



	; The following section is what actually SENDS the email
	m := ComObjCreate("Outlook.Application").CreateItem(0)
	m.Subject := "Thank you!"
	m.To := "hmetzger@bib-arch.org" ;This will provide a variable option, but will have to be redone: ClientEmail
	; Original - for reference: m.Body :="Here is the body... `n`n And the really cool thing about using this method, `n`n`n`n`n`n is, you can have what ever you want as the "body" and `n`n`n`n`n`n not worry about how long it is...or worry about the non-formatting issues that come from the mailto: command`n`n`n`n ...yes, that is a whole bunch of "new Lines" to show you how you can format this however you want...`n`n`n`n`n`n`n`n AND IT WORKS
	m.Body := "Dear " Salutation " `n`n Good news: your thank you letter file has been mailed. `n`n File Name: " FileName " `n`n Number Mailed: " MailQty " `n`n Date Received: " EnterDate " `n`n Date Mailed: " MailDate " `n`n Package Number:  " Package " `n`n Sincerely, `n`n`n The Five Maples Team"
	m.Display ;to display the email message...and the really cool part, if you leave this line out, it will not show the window............... but the m.send below will still send the email!!!
	; m.Send ;to automatically send and CLOSE that new email window...
	MsgBox, In the future, the email will send automatically - this just gives you a chance to review the first few times. Please let marvin@fivemaples.com know when to 'pull the switch'.
	Return
}




;===========================================================
;==  Dedupe Stuff
;===========================================================


ButtonJustDupeIndexes:
{
Gui, Submit
WinWait, ahk_class TMailList_Form,
IfWinNotActive, ahk_class TMailList_Form, , WinActivate, ahk_class TMailList_Form,
WinWaitActive, ahk_class TMailList_Form
SetupJustDupeIndexes()
Return
}

ButtonStdDedupe:
{
Gui, Submit
WinWait, ahk_class TMailList_Form,
IfWinNotActive, ahk_class TMailList_Form, , WinActivate, ahk_class TMailList_Form,
WinWaitActive, ahk_class TMailList_Form
SetupStdDedupe()
Return
}





SetupJustDupeIndexes()
{
WinWait, ahk_class TMailList_Form,
IfWinNotActive, ahk_class TMailList_Form, , WinActivate, ahk_class TMailList_Form,
WinWaitActive, ahk_class TMailList_Form
sleep, 100
Gui, Submit  ; Save the input from the user to each control's associated variable.

Progress, w250,,, Creating Indexes - Please Hold Your Ponies

; Start of Script, this chunk of text gets us INTO the indexes - RUN FROM BASE MM2010 WINDOW!
Send, {CTRLDOWN}a{CTRLUP}
WinWait, Duplicate Records Processing,
IfWinNotActive, Duplicate Records Processing, , WinActivate, Duplicate Records Processing,
WinWaitActive, Duplicate Records Processing,
Send, {ALTDOWN}n{ALTUP}
WinWait, Indexes,
IfWinNotActive, Indexes, , WinActivate, Indexes,
WinWaitActive, Indexes,

Progress, 12

; Creates ZCLF2, sets it to an expression and then inputs it.
Send, {ALTDOWN}n{ALTUP}
WinWait, Index Name,
IfWinNotActive, Index Name, , WinActivate, Index Name,
WinWaitActive, Index Name,
Send, ZCLF2{ENTER}
WinWait, Indexes,
IfWinNotActive, Indexes, , WinActivate, Indexes,
WinWaitActive, Indexes,
Send, {ALTDOWN}xp{ALTUP}
SendInput {Raw}left([ZIP+4_],5)+left([Company_],5)+left([Last Name_],5)+left([First Name_],5)+left([Add2_],12) ; Sending ZCLF2

Progress, 25

; Creates ZLF2, sets it to an expression and then inputs it.
Send, {ALTDOWN}n{ALTUP}
WinWait, Index Name,
IfWinNotActive, Index Name, , WinActivate, Index Name,
WinWaitActive, Index Name,
Send, ZLF2{ENTER}
WinWait, Indexes,
IfWinNotActive, Indexes, , WinActivate, Indexes,
WinWaitActive, Indexes,
Send, {ALTDOWN}xp{ALTUP}
SendInput {Raw}left([ZIP+4_],5)+left([Last Name_],5)+left([First Name_],5)+left([Add2_],5) ; Sending ZLF2

Progress, 37

; Creates ZLF, sets it to an expression and then inputs it.
Send, {ALTDOWN}n{ALTUP}
WinWait, Index Name,
IfWinNotActive, Index Name, , WinActivate, Index Name,
WinWaitActive, Index Name,
Send, ZLF{ENTER}
WinWait, Indexes,
IfWinNotActive, Indexes, , WinActivate, Indexes,
WinWaitActive, Indexes,
Send, {ALTDOWN}xp{ALTUP}
SendInput {Raw}left([ZIP+4_],5)+left([Last Name_],5)+[First Name_] ; Sending ZLF

Progress, 50

; Creates ZL2, sets it to an expression and then inputs it.
Send, {ALTDOWN}n{ALTUP}
WinWait, Index Name,
IfWinNotActive, Index Name, , WinActivate, Index Name,
WinWaitActive, Index Name,
Send, ZL2{ENTER}
WinWait, Indexes,
IfWinNotActive, Indexes, , WinActivate, Indexes,
WinWaitActive, Indexes,
Send, {ALTDOWN}xp{ALTUP}
SendInput {Raw}left([ZIP+4_],5)+left([Last Name_],5)+[Add2_] ; Sending ZL2

Progress, 62

; Creates ZCL, sets it to an expression and then inputs it.
Send, {ALTDOWN}n{ALTUP}
WinWait, Index Name,
IfWinNotActive, Index Name, , WinActivate, Index Name,
WinWaitActive, Index Name,
Send, ZCL{ENTER}
WinWait, Indexes,
IfWinNotActive, Indexes, , WinActivate, Indexes,
WinWaitActive, Indexes,
Send, {ALTDOWN}xp{ALTUP}
SendInput {Raw}Left([ZIP+4_],5)+Left([Company_],10)+Left([Last Name_],10) ; Sending ZCL

Progress, 75

; Creates ZDPL, sets it to an expression and then inputs it.
Send, {ALTDOWN}n{ALTUP}
WinWait, Index Name,
IfWinNotActive, Index Name, , WinActivate, Index Name,
WinWaitActive, Index Name,
Send, ZDPL{ENTER}
WinWait, Indexes,
IfWinNotActive, Indexes, , WinActivate, Indexes,
WinWaitActive, Indexes,
Send, {ALTDOWN}xp{ALTUP}
SendInput {Raw}[ZIP+4_] + [D.P. CODE] + left([Last Name_],3) ; Sending ZDPL

Progress, 98

; Creates Z2 (AKA ZA2), sets it to an expression and then inputs it.
Send, {ALTDOWN}n{ALTUP}
WinWait, Index Name,
IfWinNotActive, Index Name, , WinActivate, Index Name,
WinWaitActive, Index Name,
Send, Z2{ENTER}
Progress, 99
WinWait, Indexes,
IfWinNotActive, Indexes, , WinActivate, Indexes,
WinWaitActive, Indexes,
Send, {ALTDOWN}xp{ALTUP}
SendInput {Raw}Left([ZIP+4_],5)+left([Add2_],15) ; Sending Z2

; End of Script, exits back to "Indexes"
Send, {ALTDOWN}o{ALTUP}{ENTER}
Progress, 100
Sleep, 10
Progress, Off
Return
}


SetupStdDedupe()
{
	WinWait, ahk_class TMailList_Form,
	IfWinNotActive, ahk_class TMailList_Form, , WinActivate, ahk_class TMailList_Form,
	WinWaitActive, ahk_class TMailList_Form
	sleep, 100
	Gui, Submit  ; Save the input from the user to each control's associated variable.

	;/*

	Progress, w250,,, Setting up your Dedupe - Please Hold Your Ponies

	; Start of Script, this chunk of text gets us INTO the indexes - RUN FROM BASE MM2010 WINDOW!
	Send, {CTRLDOWN}d{CTRLUP}
	WinWait, Distribution Report,
	IfWinNotActive, Distribution Report, , WinActivate, Distribution Report,
	WinWaitActive, Distribution Report,
	Send, {ALTDOWN}n{ALTUP}
	WinWait, Indexes,
	IfWinNotActive, Indexes, , WinActivate, Indexes,
	WinWaitActive, Indexes,

	Progress, 12

	; Possible modularity chance from here...

	Send, {ALTDOWN}n{ALTUP}
	WinWait, Index Name,
	IfWinNotActive, Index Name, , WinActivate, Index Name,
	WinWaitActive, Index Name,
	Send, ZCLF2{ENTER}
	WinWait, Indexes,
	IfWinNotActive, Indexes, , WinActivate, Indexes,
	WinWaitActive, Indexes,
	Send, {ALTDOWN}xp{ALTUP}
	SendInput {Raw}left([ZIP+4_],5)+left([Company_],5)+left([Last Name_],5)+left([First Name_],5)+left([Add2_],12) ; Sending ZCLF2

	Progress, 25

	; Creates ZLF2, sets it to an expression and then inputs it.
	Send, {ALTDOWN}n{ALTUP}
	WinWait, Index Name,
	IfWinNotActive, Index Name, , WinActivate, Index Name,
	WinWaitActive, Index Name,
	Send, ZLF2{ENTER}
	WinWait, Indexes,
	IfWinNotActive, Indexes, , WinActivate, Indexes,
	WinWaitActive, Indexes,
	Send, {ALTDOWN}xp{ALTUP}
	SendInput {Raw}left([ZIP+4_],5)+left([Last Name_],5)+left([First Name_],5)+left([Add2_],5) ; Sending ZLF2

	Progress, 37

	; Creates ZLF, sets it to an expression and then inputs it.
	Send, {ALTDOWN}n{ALTUP}
	WinWait, Index Name,
	IfWinNotActive, Index Name, , WinActivate, Index Name,
	WinWaitActive, Index Name,
	Send, ZLF{ENTER}
	WinWait, Indexes,
	IfWinNotActive, Indexes, , WinActivate, Indexes,
	WinWaitActive, Indexes,
	Send, {ALTDOWN}xp{ALTUP}
	SendInput {Raw}left([ZIP+4_],5)+left([Last Name_],5)+[First Name_] ; Sending ZLF

	Progress, 50

	; Creates ZL2, sets it to an expression and then inputs it.
	Send, {ALTDOWN}n{ALTUP}
	WinWait, Index Name,
	IfWinNotActive, Index Name, , WinActivate, Index Name,
	WinWaitActive, Index Name,
	Send, ZL2{ENTER}
	WinWait, Indexes,
	IfWinNotActive, Indexes, , WinActivate, Indexes,
	WinWaitActive, Indexes,
	Send, {ALTDOWN}xp{ALTUP}
	SendInput {Raw}left([ZIP+4_],5)+left([Last Name_],5)+[Add2_] ; Sending ZL2

	Progress, 62

	; Creates ZCL, sets it to an expression and then inputs it.
	Send, {ALTDOWN}n{ALTUP}
	WinWait, Index Name,
	IfWinNotActive, Index Name, , WinActivate, Index Name,
	WinWaitActive, Index Name,
	Send, ZCL{ENTER}
	WinWait, Indexes,
	IfWinNotActive, Indexes, , WinActivate, Indexes,
	WinWaitActive, Indexes,
	Send, {ALTDOWN}xp{ALTUP}
	SendInput {Raw}Left([ZIP+4_],5)+Left([Company_],10)+Left([Last Name_],10) ; Sending ZCL

	Progress, 75

	; Creates ZDPL, sets it to an expression and then inputs it.
	Send, {ALTDOWN}n{ALTUP}
	WinWait, Index Name,
	IfWinNotActive, Index Name, , WinActivate, Index Name,
	WinWaitActive, Index Name,
	Send, ZDPL{ENTER}
	WinWait, Indexes,
	IfWinNotActive, Indexes, , WinActivate, Indexes,
	WinWaitActive, Indexes,
	Send, {ALTDOWN}xp{ALTUP}
	SendInput {Raw}[ZIP+4_] + [D.P. CODE] + left([Last Name_],3) ; Sending ZDPL


	; Creates Z2 (AKA ZA2), sets it to an expression and then inputs it.
	Send, {ALTDOWN}n{ALTUP}
	WinWait, Index Name,
	IfWinNotActive, Index Name, , WinActivate, Index Name,
	WinWaitActive, Index Name,
	Send, Z2{ENTER}
	WinWait, Indexes,
	IfWinNotActive, Indexes, , WinActivate, Indexes,
	WinWaitActive, Indexes,
	Send, {ALTDOWN}xp{ALTUP}
	SendInput {Raw}Left([ZIP+4_],5)+left([Add2_],15) ; Sending Z2

	; to here.

	Progress, 90

	Send, {ALTDOWN}n{ALTUP}
	WinWait, Index Name,
	IfWinNotActive, Index Name, , WinActivate, Index Name,
	WinWaitActive, Index Name,
	Send, Dedupe{ENTER}
	WinWait, Indexes,
	IfWinNotActive, Indexes, , WinActivate, Indexes,
	WinWaitActive, Indexes,
	Send, {ALTDOWN}xp{ALTUP}
	SendInput {Raw}[Source List] ; Naming Source List
	Send, {ALTDOWN}o{ALTUP}
	WinWait, Confirm,
	IfWinNotActive, Confirm, , WinActivate, Confirm,
	WinWaitActive, Confirm,
	Send, y
	WinWait, Distribution Report,
	IfWinNotActive, Distribution Report, , WinActivate, Distribution Report,
	WinWaitActive, Distribution Report,
	Progress, 94
	Send, {SHIFTDOWN}{TAB}{SHIFTUP}dedupe{ALTDOWN}b{ALTUP}
	WinWait, Save Procedure Information,
	IfWinNotActive, Save Procedure Information, , WinActivate, Save Procedure Information,
	WinWaitActive, Save Procedure Information,
	Sleep, 100
	Send, {ALTDOWN}d{ALTUP}

	Progress, 98

	;Printing the distribution report - this may need to change later

	WinWait, Distribution Report - Print,
	IfWinNotActive, Distribution Report - Print, , WinActivate, Distribution Report - Print,
	WinWaitActive, Distribution Report - Print,
	;Send, {ALTDOWN}n{ALTUP}adobe{ALTDOWN}o{ALTUP}{ENTER}

	; End of Script, exits back to "Indexes"
	Progress, 100
	Sleep, 10
	Progress, Off
	*/
	MsgBox, 262144,, Your dedupe is setup and ready to go - just choose where you want the distribution report to be printed!
	Return

}


; #singleinstance force ; Needed, but already in file


SetTitleMatchMode, 2 ; I could likely make the whole document this, but just to be safe I'm setting it here


;===========================================================
;==  Dedupe Module (Always active, but only where needed)
;===========================================================



#IfWinActive Merge/Edit Duplicate Records


^NumpadAdd::
{
InputBox, MouseSpeed, Enter Speed Below, Current Speed: %MouseSpeed%
If MouseSpeed < 1
	MouseSpeed = 20
MouseSUp = (%MouseSpeed% / 2)
MouseSDown = %MouseSpeed%
Return
}






XButton1:: ;c 🌟 Move Up & Down Using Mouse Buttons While Deduping ⌨️ The Foward & Backward Mouse buttons (if available), or roll the Mouse Wheel | This will navagate up and down your dedupe screen, and the middle mouse button will UNHIDE all dupes in the selected group. The Mouse Wheel should work too.
while GetKeyState("XButton1", "P") ; this loop only executes when the XButton is pressed down
{
	 MouseSpeed = %MouseSpeed%
	 Send, {Down} ; this presses down only while the above condition is met
	 sleep, %MouseSpeed%
}
;Send, {Alt Up}{LButton Up} ; this liberates when the above condition is not met anymore
return


XButton2::
while GetKeyState("XButton2", "P") ; this loop only executes when the XButton is pressed down
{
	 Send, {Up} ; this presses up only while the above condition is met
	 ; Try this next time Marv:
	 ;MouseSpeed = (%MouseSpeed% / 2)
	 Send, {Up} ; this presses down only while the above condition is met
	 sleep, %MouseSpeed%   ; OLD VERSION: sleep, %MouseSpeed%   ;sleep, (%MouseSpeed% / 2)
	 ;MouseSpeed = (%MouseSpeed% / 0.5)
}
;Send, {Alt Up}{LButton Up} ; this liberates when the above condition is not met anymore
return


MButton::
suspend, On
send {RButton}
sleep, 50
send g
sleep, 50
send u
suspend, Off
return

;====================================================
;=== The following section totally WORKS - it's just
;=== been disabled because I don't need it. Simply
;=== delete the two indicated comment marks (or the
;=== whole lines) to re-enable. Thanks!  -Past Marv
;====================================================

;/* ; <-- Delete this line to re-enable as well as the bottom one
WheelDown::
;while GetKeyState("WheelDown", "P") ; this loop only executes when the XButton is pressed down
Loop, 10
{
    MouseSpeed = %MouseSpeed%
    Send, {Down} ; this presses down only while the above condition is met
    sleep, %MouseSpeed%
}
;Send, {Alt Up}{LButton Up} ; this liberates when the above condition is not met anymore
return

WheelUp::
;while GetKeyState("WheelUp", "P") ; this loop only executes when the XButton is pressed down
Loop, 5
{
    Send, {Up} ; this presses up only while the above condition is met
}
;Send, {Alt Up}{LButton Up} ; this liberates when the above condition is not met anymore
return
;*/ ; <-- Delete this line to re-enable as well as the top one



;====================================================================
;==  TEST SUITE (All Tests - if fail, pump failure to TestResults)
;====================================================================

#IfWinActive ShortcutToolkit.ahk

+!r:: ; Ctrl + Alt + R - Instantly reload & reboot script
{
    msgbox, , , Reloading Script, 0.30
    Reload
    ExitApp
    return
}

#IfWinActive

ClipboardTestFunc(TestClip1, TestClip2)
{
	if (TestClip1 == TestClip2)
    {
		return
    }
    else
    {
        throw Exception "Fail - Clipboard management not working"
    }
}

TestResults := ""
DebugVar := 0
^+b:: ;c 🌟 Run ShortcutToolkit self-test suite ⌨️ Ctrl+Shift+B - ONLY USE IF YOU'RE KENDRA OR MARVIN

; This is the "Test Suite" - run it and it will complain of issues. Once run, it reboots the script for a clean working state.
{
    ; TODO Add a "do you REALLY wanna run tests and reset the whole damn thing" button
    DebugVar := 1
    ;TEST - This tests for UNICODE encoding being changed to ASCII - if yes, then shruggie won't work no mo.
    if ("(?)" == "(ツ)")
    {
        TestResults := "Hey - someone screwed up the ASCII/Unicode encoding - see Marvin."
    }

    ;TEST - This tests the clipboard manager (TODO: Make it more comprehensive)
    ClipHolder := Clipboard
    ClipHolder1 := clpb1
    Clipboard := "Goat"
    clpb1 := "Pony"
    gosub TestClipCapture
    gosub ButtonClipboardManager
    try
    {
        ClipboardTestFunc("Goat", clpb1)
	}
	catch e
	{
		if (TestResults <> "")
		{
			TestResults := TestResults . " | " . e
		}
		else
		{
			TestResults := e
		}
	}
    Clipboard := ClipHolder
    clpb1 := ClipHolder1

    ;Displays test results
    if (TestResults <> "")
    {
        msgbox Some tests failed, click OK to commence reboot. Failed tests: %TestResults%
        Reload
        ExitApp
    }
    Else
        msgbox, , , TESTS PASSED!, 1.00
        DebugVar := 0
        return
}

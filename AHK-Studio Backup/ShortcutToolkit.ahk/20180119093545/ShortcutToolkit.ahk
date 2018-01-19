#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases, prevents checking empty variables
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Persistent
#SINGLEINSTANCE force
;SetBatchLines, -1   ; Runs the script at max speed, default is 10 or 20 ms

; Consider putting a "First Run" quick tip upon first running the file - a quick 3 second splash screen that indicates to hit the "help" key to find out more.

/*
vFirstRun := 0
If (vFirstRun <> 0)
{
MsgBox, Hit Win+F2 to learn moar!
}
vFirstRun := 1
*/

; Where ALL of this is stored on Dropbox (update there to trigger propigation): https://www.dropbox.com/sh/rytpiimr5snjnbk/AAB_nORecls7kcCB5ldlV-tAa?dl=0



:*:Jam`t::James Chase`t8023875157`t110 ;c Typing Jam and then hitting tab or space will auto-populate the warehouse info on the NDC/SCF 8125 screen


;===========================================================
;==  DP Dashboard
;===========================================================


^space:: ; !#d:: ;Used to be Win+Alt+D ;c Hitting Ctrl+Space will launch what I like to call the "DP Dashboard", which contains most of the clickable commands (will also close it if it's open)
{

Gui, New,,DP Dashboard
Gui, Add, Text,, Standard Dedupe Indexes && Dist Report:
Gui, Add, Text,, JUST Dedupe Indexes:
Gui, Add, Text,, Clipboard Manager:
; Gui, Add, Text,, Postal One! Login:    ; Come back for this one after release
Gui, Add, Text,, WH NDC/SCF Info:
Gui, Add, Button, ym, &Std Dedupe  ; The ym option starts a new column of controls, and the label ButtonStdDedupe (if it exists) will be run when the button is pressed.
Gui, Add, Button,, &Just Dupe Indexes
Gui, Add, Button, default, &ClipboardMgr
;Gui, Add, Button,, &Login ; Come back for this one after release
Gui, Add, Button,, &WHIMB
Gui, Show,,DP Dashboard
return  ; End of auto-execute section. The script is idle until the user does something.

}



;===========================================================
;==  Clipboard Management
;===========================================================

~^c::
{
sleep, 100
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
; tooltip, 1: %clpbs1%`n2: %clpbs2%`n3: %clpbs3%`n4: %clpbs4%`n5: %clpbs5% ; Uncomment this to get a "popup" of what was copied… every time someone copies.
sleep, 1000
tooltip
return
}

ButtonClipboardMgr:
^+v:: ;c Ctrl+Shift+V will let you select one of the last 5 things you've copied, assuming they were text.`nIt will time out and default to your current clipboard after 3 seconds.
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
	; inputbox, clpbNbr, which clipboard number do you want to have in your clipboard now?, 1: %clpbs1% 2: %clpbs2%`n3: %clpbs3% 4: %clpbs4%`n5: %clpbs5%`n   ; In place originally, replaced with my GUI below
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
	Loop{
			GuiControl,, varProgressMeter, +1
			sleep, 30
		} Until A_Index > 99
	GoSub, ButtonGo
	Return
	
	
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






ButtonWHIMB:
{
Gui, Submit 
SendInput James Chase`t8023875157`t110
Return
}

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
	Send, {ALTDOWN}n{ALTUP}adobe{ALTDOWN}o{ALTUP}{ENTER}

	; End of Script, exits back to "Indexes"
	Progress, 100
	Sleep, 10
	Progress, Off
	*/
	MsgBox, 262144,, Your dedupe is setup and ready to go - just choose where you want the distribution report to be saved!
	Return

}


; #singleinstance force ; Needed, but already in file

#F2:: ;c WinKey+F2 will bring up a help file, which attempts to automatically document the functions found here-in. It may go without saying, but I'm still working on it. :D
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
  msgbox %comment%
Return
}












SetTitleMatchMode, 2 ; I could likely make the whole document this, but just to be safe I'm setting it here


;===========================================================
;==  Dedupe Module (Always active, but only when needed
;===========================================================


; #IfWinActive WinCrypt_Sync.ahk ; This is just here since I know it works,
; it should be a way to verify that what I'm doing works
#IfWinActive Merge/Edit Duplicate Records
MouseSpeed = 3 ; This is a slowing modifier - the higher it is, the slower the down button goes, default is 15

^NumpadAdd:: InputBox, MouseSpeed, Enter Speed Below, Current Speed: %MouseSpeed%







XButton1:: ;c The Foward & Backward Mouse buttons (if available) will navagate up and down your dedupe screen, and the middle mouse button will UNHIDE all dupes in the selected group.
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


#IfWinActive DP Dashboard ahk_class AutoHotkeyGUI
^space::
{
	Gui DPDashboard:Destroy
	Return	
}
#IfWinActive

;/* ; Fix below with context sensitive : IfWinActive ahk_class AutoHotkeyGUI ; or something of the like
;Escape::
GuiClose:
CloseAllWindows: ; Close all and reset the GUI number
{
	Reload
	ExitApp
}



/* 
;===========================================================
;==  NOTES
;===========================================================

~ = Passthrough, at least when used as a prefix for a button combo, so ~^c would be Ctrl+c that also sends the same command to the native system
^ = Ctrl
# = Win key (I think?)
! = Alt Key (I think)


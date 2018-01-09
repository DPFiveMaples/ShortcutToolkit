#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Persistent
#SINGLEINSTANCE force

:*:Jam`t::James Chase`t8023875157`t110


!#d::
{

; Example: A simple input-box that asks for first name and last name:

Gui, Add, Text,, Standard Dedupe Indexes & Dist Report:
Gui, Add, Text,, JUST Dedupe Indexes:
Gui, Add, Text,, Postal One! Login:
Gui, Add, Text,, WH NDC/SCF Info:
Gui, Add, Button, default ym, &StdDedupe  ; The ym option starts a new column of controls, and the label ButtonStdDedupe (if it exists) will be run when the button is pressed.
Gui, Add, Button,, JustDupeIndexes
Gui, Add, Button,, Login
Gui, Add, Button,, WHIMB
Gui, Show,, Five Maples Dashboard
return  ; End of auto-execute section. The script is idle until the user does something.

}

~Esc::Gui Cancel

GuiClose:
{
Gui, Submit
Gui Cancel
return
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

	MsgBox, %vP1Blob%

	;WinWait, Index Name, 
	;IfWinNotActive, Index Name, , WinActivate, Index Name, 
	;WinWaitActive, Index Name, 


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

/*

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



; Random Notes that likely aren't needed:
/*

GetUnP() ; In the future, I can make this more general by making FMAHK.config a passed argument
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
	FileAppend, %vP1User%|||||%vP1Pass%, %A_MyDocuments%\FMAHK.config ; If I need new lines for some reason, add with this: `n
	Return
}


:*:Jam`t::James Chase`t8023875157`t110
:*:gh`t::Gary Henricksen
:*:db`t::Dawn Brennan
:*:gr`t::Ginny Ricker
:*:cm`t::Carol Martin
:*:bm`t::Brett Morrison
:*:ec`t::Elyse Carter
:*:ad`t::Anna Duca
:*:jh`t::Jillian Hobday
:*:kh`t::Kayla Hornbrook
:*:shruggie`t::¯\_(ツ)_/¯
:*:nonus`t::(Not ValidUSPSZIPCode(MainAddrGrp)) AND {ENTER}(Not InTable([State_], "Z:\Data Files\States.rpl"))
SendMode Input
:*:note`t::Note - Highlighted record(s) contain PO Box(s) - we're looking for a physical address and will get it to you ASAP
:*:tabret`t::{CTRLDOWN}h{CTRLUP}\t{Tab}\r\n{ALTDOWN}ta{ALTUP}{ESCAPE}
:*:prfsel`t::(Not Empty([Proofs])) AND {ENTER}([Proofs] Excludes "NOT A PROOF")
; */

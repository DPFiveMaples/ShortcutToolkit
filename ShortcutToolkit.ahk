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

~Esc::
{
Gui, Submit
Gui Cancel
Check_ForUpdate(1)
return
}
GuiClose:
{
Gui, Submit
Gui Cancel
Check_ForUpdate(1)
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

	Loop 
	{
    Loop, read, %A_MyDocuments%\FMAHK.config
        last_line := USPSLogin|_|_|  
    if InStr(last_line,"login:")
        break
	}
MsgBox Found!

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











;This is the function-call. The part that tells autohotkey to execute the function below.
;Check_ForUpdate(1)

;The 1 tells the function to use 1 as the first paramater - setting "_ReplaceCurrentSCript" to 1.

;This is the actual function.
Check_ForUpdate(_ReplaceCurrentScript = 0, _SuppressMsgBox = 0, _CallbackFunction = "", ByRef _Information = "")
{
	; MsgBox, TEST!
	;Version.ini file format
	;
	;[Info]
	;Version=1.4
	;URL=https://gitlab.com/Marvin_FiveMaples/Dedupe-Index-Script/raw/master/ShortcutToolkit.ahk or .exe
	;MD5=00000000000000000000000000000000 or omit this key completly to skip the MD5 file validation
	
	Static Script_Name := "ShortcutToolkit" ;Your script name
	, Version_Number := 0.3 ;The script's version number
	, Update_URL := "https://gitlab.com/Marvin_FiveMaples/Dedupe-Index-Script/raw/master/Version.ini" ;The URL of the version.ini file for your script
	, Retry_Count := 3 ;Retry count for if/when anything goes wrong
	
	Random,Filler,10000000,99999999
	Version_File := A_Temp . "\" . Filler . ".ini"
	, Temp_FileName := A_Temp . "\" . Filler . ".tmp"
	, VBS_FileName := A_Temp . "\" . Filler . ".vbs"
	
	Loop,% Retry_Count
	{
		_Information := ""
		
		UrlDownloadToFile,%Update_URL%,%Version_File%
		
		IniRead,Version,%Version_File%,Info,Version,N/A
		
		If (Version = "N/A"){
			FileDelete,%Version_File%
			
			If (A_Index = Retry_Count)
				_Information .= "The version info file doesn't have a ""Version"" key in the ""Info"" section or the file can't be downloaded."
			Else
				Sleep,500
			
			Continue
		}
		
		If (Version > Version_Number){
			If (_SuppressMsgBox != 1 and _SuppressMsgBox != 3){
				MsgBox,0x4,New version available,There is a new version of %Script_Name% available.`nCurrent version: %Version_Number%`nNew version: %Version%`n`nWould you like to download it now?
				
				IfMsgBox,Yes
					MsgBox_Result := 1
			}
			
			If (_SuppressMsgBox or MsgBox_Result){
				IniRead,URL,%Version_File%,Info,URL,N/A
				
				If (URL = "N/A")
					_Information .= "The version info file doesn't have a valid URL key."
				Else {
					SplitPath,URL,,,Extension
					
					If (Extension = "ahk" And A_AHKPath = "")
						_Information .= "The new version of the script is an .ahk filetype and you do not have AutoHotKey installed on this computer.`r`nReplacing the current script is not supported."
					Else If (Extension != "exe" And Extension != "ahk")
						_Information .= "The new file to download is not an .EXE or an .AHK file type. Replacing the current script is not supported."
					Else {
						IniRead,MD5,%Version_File%,Info,MD5,N/A
						
						Loop,% Retry_Count
						{
							UrlDownloadToFile,%URL%,%Temp_FileName%
							
							IfExist,%Temp_FileName%
							{
								If (MD5 = "N/A"){
									_Information .= "The version info file doesn't have a valid MD5 key."
									, Success := True
									Break
								} Else {
									H := DllCall("CreateFile","Str",Temp_FileName,"UInt",0x80000000,"UInt",3,"UInt",0,"UInt",3,"UInt",0,"UInt",0)
									, VarSetCapacity(FileSize,8,0)
									, DllCall("GetFileSizeEx","UInt",H,"Int64",&FileSize)
									, FileSize := NumGet(FileSize,0,"Int64")
									, FileSize := FileSize = -1 ? 0 : FileSize
									
									If (FileSize != 0){
										VarSetCapacity(Data,FileSize,0)
										, DllCall("ReadFile","UInt",H,"UInt",&Data,"UInt",FileSize,"UInt",0,"UInt",0)
										, DllCall("CloseHandle","UInt",H)
										, VarSetCapacity(MD5_CTX,104,0)
										, DllCall("advapi32\MD5Init",Str,MD5_CTX)
										, DllCall("advapi32\MD5Update",Str,MD5_CTX,"UInt",&Data,"UInt",FileSize)
										, DllCall("advapi32\MD5Final",Str,MD5_CTX)
										
										FileMD5 := ""
										Loop % StrLen(Hex:="123456789ABCDEF0")
											N := NumGet(MD5_CTX,87+A_Index,"Char"), FileMD5 .= SubStr(Hex,N>>4,1) . SubStr(Hex,N&15,1)
										
										VarSetCapacity(Data,FileSize,0)
										, VarSetCapacity(Data,0)
										
										If (FileMD5 != MD5){
											FileDelete,%Temp_FileName%
											
											If (A_Index = Retry_Count)
												_Information .= "The MD5 hash of the downloaded file does not match the MD5 hash in the version info file."
											Else										
												Sleep,500
											
											Continue
										} Else
											Success := True
									} Else {
										DllCall("CloseHandle","UInt",H)
										Success := True
									}
								}
							} Else {
								If (A_Index = Retry_Count)
									_Information .= "Unable to download the latest version of the file from " %URL% "."
								Else
									Sleep,500
								Continue
							}
						}
					}
				}
			}
		} Else
			_Information .= "No update was found."
		
		FileDelete,%Version_File%
		Break
	}
	
	If (_ReplaceCurrentScript And Success){
		SplitPath,URL,,,Extension
		Process,Exist
		MyPID := ErrorLevel
		
		VBS_P1 =
		(LTrim Join`r`n
			On Error Resume Next
			Set objShell = CreateObject("WScript.Shell")
			objShell.Run "TaskKill -f -im %MyPID%", WindowStyle, WaitOnReturn
			WScript.Sleep 1000
			Set objFSO = CreateObject("Scripting.FileSystemObject")
		)
		
		If (A_IsCompiled){
			If (Extension = "exe"){
				VBS_P2 =
				(LTrim Join`r`n
					objFSO.CopyFile "%Temp_FileName%", "%A_ScriptFullPath%", True
					objFSO.DeleteFile "%Temp_FileName%", True
					objShell.Run """%A_ScriptFullPath%"""
				)
				
				Return_Val :=  Temp_FileName
			} Else { ;Extension is ahk
				SplitPath,A_ScriptFullPath,,FDirectory,,FName
				FileMove,%Temp_FileName%,%FDirectory%\%FName%.ahk,1
				FileDelete,%Temp_FileName%
				
				VBS_P2 =
				(LTrim Join`r`n
					objFSO.DeleteFile "%A_ScriptFullPath%", True
					objShell.Run """%FDirectory%\%FName%.ahk"""
				)
				
				Return_Val := FDirectory . "\" . FName . ".ahk"
			}
		} Else {
			If (Extension = "ahk"){
				FileMove,%Temp_FileName%,%A_ScriptFullPath%,1
				If (Errorlevel)
					_Information .= "Error (" Errorlevel ") unable to replace current script with the latest version."
				Else {
					VBS_P2 = 
					(LTrim Join`r`n
						objShell.Run """%A_ScriptFullPath%"""
					)
					
					Return_Val :=  A_ScriptFullPath
				}
			} Else If (Extension = "exe"){
				SplitPath,A_ScriptFullPath,,FDirectory,,FName
				FileMove,%Temp_FileName%,%FDirectory%\%FName%.exe,1
				FileDelete,%A_ScriptFullPath%
				
				VBS_P2 =
				(LTrim Join`r`n
					objShell.Run """%FDirectory%\%FName%.exe"""
				)
				
				Return_Val :=  FDirectory . "\" . FName . ".exe"
			} Else {
				FileDelete,%Temp_FileName%
				_Information .= "The downloaded file is not an .EXE or an .AHK file type. Replacing the current script is not supported."
			}
		}
		
		VBS_P3 =
		(LTrim Join`r`n
			objFSO.DeleteFile "%VBS_FileName%", True
			Set objFSO = Nothing
			Set objShell = Nothing
		)
		
		If (_SuppressMsgBox < 2)
			VBS_P3 .= "`r`nWScript.Echo ""Update complected successfully."""
		
		FileDelete,%VBS_FileName%
		FileAppend,%VBS_P1%`r`n%VBS_P2%`r`n%VBS_P3%,%VBS_FileName%
		
		If (_CallbackFunction != ""){
			If (IsFunc(_CallbackFunction))
				%_CallbackFunction%()
			Else
				_Information .= "The callback function is not a valid function name."
		}
		
		RunWait,%VBS_FileName%,%A_Temp%,VBS_PID
		Sleep,2000
		
		Process,Close,%VBS_PID%
		_Information := "Error (?) unable to replace current script with the latest version.`r`nPlease make sure your computer supports running .vbs scripts and that the script isn't running in a pipe."
	}
	
	_Information := _Information = "" ? "None" : _Information
	
	Return Return_Val
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

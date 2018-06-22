#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


SetTitleMatchMode, 2 ; This let's any window that partially matches the given name get activated
#IfWinActive Excel
^+j::
XL := ComObjActive("Excel.Application") 
;handle to running application
qty := XL.Worksheets.Count
WBName = % XL.ActiveWorkbook.Name
SplitName := StrSplit(WBName, ".").1
SheetNum = 0
; ^this is a counter so I know which sheet the loop is on.
For sheets in XL.ActiveWorkbook.Worksheets{
SheetNum++
; ^each time the loop runs, it adds 1 to the counter
XL.Sheets(SheetNum).select
SheetName = % XL.ActiveSheet.Name
SplashTextOn, 200, 100, , Sheet Name: %SheetName% `n Sheet Number: %SheetNum%
Sleep, 2500
SplashTextOff 
Send, !FAY
Sleep, 500
Send, 3
Sleep, 50
Send, {shift down}%SheetName%-converted-{shift up}
Send,%SplitName%
Send,{tab}
Sleep, 500
if (SheetNum = 1){
Send, {down 12}
}
; ^this if statement will only change the file type on the first loop, it stays as text after the first save
Send, {enter}{tab}{enter}{enter}
Sleep, 2500

}
Return

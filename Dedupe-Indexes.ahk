﻿#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Persistent
#SINGLEINSTANCE force



!#d::
{


; Start of Script, this chunk of text gets us INTO the indexes - RUN FROM BASE MM2010 WINDOW!
Send, {CTRLDOWN}a{CTRLUP}
WinWait, Duplicate Records Processing, 
IfWinNotActive, Duplicate Records Processing, , WinActivate, Duplicate Records Processing, 
WinWaitActive, Duplicate Records Processing, 
Send, {ALTDOWN}n{ALTUP}
WinWait, Indexes, 
IfWinNotActive, Indexes, , WinActivate, Indexes, 
WinWaitActive, Indexes, 


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

; End of Script, exits back to "Indexes"
Send, {ALTDOWN}o{ALTUP}{ENTER}
}

; Random Notes that likely aren't needed:
/*
:*:Bru`t::Bruce McIntosh`t8023875157`t110
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
:*:sluggie`t::Ask1{Tab}Ask2{Tab}Ask3{Tab}Verify{Enter}=IF(A2<0,"",IF(A2<1,100,IF(A2<=10,50,IF(A2<=15,50,IF(A2<=20,50,IF(A2<=25,100,IF(A2<=30,100,IF(A2<=35,100,IF(A2<=40,100,IF(A2<=50,150,IF(A2<=60,150,IF(A2<=75,150,IF(A2<=100,150,IF(A2<=125,200,IF(A2<=150,250,IF(A2<=175,250,IF(A2<=200,500,IF(A2<=250,500,IF(A2<=300,500,IF(A2<=350,500,IF(A2<=400,750,IF(A2<=500,1000,IF(A2<=600,1000,IF(A2<=700,1250,IF(A2<=750,1250,IF(A2<=800,1250,IF(A2<=900,1250,IF(A2<=1000,2000,IF(A2<=1100,2000,IF(A2<=1200,2000,IF(A2<=1250,2000,IF(A2<=1500,2500,IF(A2<=1600,2500,IF(A2<=1700,2500,IF(A2<=1800,2500,IF(A2<=1900,2500,IF(A2<=2000,5000,IF(A2<=2500,10000,IF(A2<=3000,10000,IF(A2<=3500,10000,IF(A2<=4000,10000,IF(A2<=5000,10000,IF(A2<=6000,10000,IF(A2<=7000,15000,IF(A2<=8000,15000,IF(A2<=9000,15000,IF(A2<10000,15000,IF(A2=10000,20000,"")))))))))))))))))))))))))))))))))))))))))))))))){tab}=IF(A2<0,"",IF(A2<1,50,IF(A2<=10,25,IF(A2<=15,25,IF(A2<=20,25,IF(A2<=25,50,IF(A2<=30,50,IF(A2<=35,50,IF(A2<=40,50,IF(A2<=50,100,IF(A2<=60,100,IF(A2<=75,100,IF(A2<=100,125,IF(A2<=125,150,IF(A2<=150,200,IF(A2<=175,200,IF(A2<=200,250,IF(A2<=250,300,IF(A2<=300,400,IF(A2<=350,400,IF(A2<=400,500,IF(A2<=500,750,IF(A2<=600,750,IF(A2<=700,1000,IF(A2<=750,1000,IF(A2<=800,1000,IF(A2<=900,1000,IF(A2<=1000,1500,IF(A2<=1100,1500,IF(A2<=1200,1500,IF(A2<=1250,1500,IF(A2<=1500,2000,IF(A2<=1600,2000,IF(A2<=1700,2000,IF(A2<=1800,2000,IF(A2<=1900,2000,IF(A2<=2000,2500,IF(A2<=2500,5000,IF(A2<=3000,5000,IF(A2<=3500,5000,IF(A2<=4000,5000,IF(A2<=5000,7500,IF(A2<=6000,7500,IF(A2<=7000,10000,IF(A2<=8000,10000,IF(A2<=9000,10000,IF(A2<10000,12500,IF(A2=10000,15000,"")))))))))))))))))))))))))))))))))))))))))))))))){tab}=IF(A2<0,"",IF(A2<1,25,IF(A2<=10,10,IF(A2<=15,15,IF(A2<=20,20,IF(A2<=25,25,IF(A2<=30,30,IF(A2<=35,35,IF(A2<=40,40,IF(A2<=50,50,IF(A2<=60,60,IF(A2<=75,75,IF(A2<=100,100,IF(A2<=125,125,IF(A2<=150,150,IF(A2<=175,175,IF(A2<=200,200,IF(A2<=250,250,IF(A2<=300,300,IF(A2<=350,350,IF(A2<=400,400,IF(A2<=500,500,IF(A2<=600,600,IF(A2<=700,700,IF(A2<=750,750,IF(A2<=800,800,IF(A2<=900,900,IF(A2<=1000,1000,IF(A2<=1100,1100,IF(A2<=1200,1200,IF(A2<=1250,1250,IF(A2<=1500,1500,IF(A2<=1600,1600,IF(A2<=1700,1700,IF(A2<=1800,1800,IF(A2<=1900,1900,IF(A2<=2000,2000,IF(A2<=2500,2500,IF(A2<=3000,3000,IF(A2<=3500,3500,IF(A2<=4000,4000,IF(A2<=5000,5000,IF(A2<=6000,6000,IF(A2<=7000,7000,IF(A2<=8000,8000,IF(A2<=9000,9000,IF(A2<10000,10000,IF(A2=10000,10000,"")))))))))))))))))))))))))))))))))))))))))))))))){tab}=IF(OR(IF(AND(A2<1,A2>=0,B2=100,C2=50,D2=25),TRUE,FALSE),IF(AND(A2>=1,A2<=10,B2=50,C2=25,D2=10),TRUE,FALSE),IF(AND(A2>10,A2<=15,B2=50,C2=25,D2=15),TRUE,FALSE),IF(AND(A2>15,A2<=20,B2=50,C2=25,D2=20),TRUE,FALSE),IF(AND(A2>20,A2<=25,B2=100,C2=50,D2=25),TRUE,FALSE),IF(AND(A2>25,A2<=30,B2=100,C2=50,D2=30),TRUE,FALSE),IF(AND(A2>30,A2<=35,B2=100,C2=50,D2=35),TRUE,FALSE),IF(AND(A2>35,A2<=40,B2=100,C2=50,D2=40),TRUE,FALSE),IF(AND(A2>40,A2<=50,B2=150,C2=100,D2=50),TRUE,FALSE),IF(AND(A2>50,A2<=60,B2=150,C2=100,D2=60),TRUE,FALSE),IF(AND(A2>60,A2<=75,B2=150,C2=100,D2=75),TRUE,FALSE),IF(AND(A2>75,A2<=100,B2=150,C2=125,D2=100),TRUE,FALSE),IF(AND(A2>100,A2<=125,B2=200,C2=150,D2=125),TRUE,FALSE),IF(AND(A2>125,A2<=150,B2=250,C2=200,D2=150),TRUE,FALSE),IF(AND(A2>150,A2<=175,B2=250,C2=200,D2=175),TRUE,FALSE),IF(AND(A2>175,A2<=200,B2=500,C2=250,D2=200),TRUE,FALSE),IF(AND(A2>200,A2<=250,B2=500,C2=300,D2=250),TRUE,FALSE),IF(AND(A2>250,A2<=300,B2=500,C2=400,D2=300),TRUE,FALSE),IF(AND(A2>300,A2<=350,B2=500,C2=400,D2=350),TRUE,FALSE),IF(AND(A2>350,A2<=400,B2=750,C2=500,D2=400),TRUE,FALSE),IF(AND(A2>400,A2<=500,B2=1000,C2=750,D2=500),TRUE,FALSE),IF(AND(A2>500,A2<=600,B2=1000,C2=750,D2=600),TRUE,FALSE),IF(AND(A2>600,A2<=700,B2=1250,C2=1000,D2=700),TRUE,FALSE),IF(AND(A2>700,A2<=750,B2=1250,C2=1000,D2=750),TRUE,FALSE),IF(AND(A2>750,A2<=800,B2=1250,C2=1000,D2=800),TRUE,FALSE),IF(AND(A2>800,A2<=900,B2=1250,C2=1000,D2=900),TRUE,FALSE),IF(AND(A2>900,A2<=1000,B2=2000,C2=1500,D2=1000),TRUE,FALSE),IF(AND(A2>1000,A2<=1100,B2=2000,C2=1500,D2=1100),TRUE,FALSE),IF(AND(A2>1100,A2<=1200,B2=2000,C2=1500,D2=1200),TRUE,FALSE),IF(AND(A2>1200,A2<=1250,B2=2000,C2=1500,D2=1250),TRUE,FALSE),IF(AND(A2>1250,A2<=1500,B2=2500,C2=2000,D2=1500),TRUE,FALSE),IF(AND(A2>1500,A2<=1600,B2=2500,C2=2000,D2=1600),TRUE,FALSE),IF(AND(A2>1600,A2<=1700,B2=2500,C2=2000,D2=1700),TRUE,FALSE),IF(AND(A2>1700,A2<=1800,B2=2500,C2=2000,D2=1800),TRUE,FALSE),IF(AND(A2>1800,A2<=1900,B2=2500,C2=2000,D2=1900),TRUE,FALSE),IF(AND(A2>1900,A2<=2000,B2=5000,C2=2500,D2=2000),TRUE,FALSE),IF(AND(A2>2000,A2<=2500,B2=10000,C2=5000,D2=2500),TRUE,FALSE),IF(AND(A2>2500,A2<=3000,B2=10000,C2=5000,D2=3000),TRUE,FALSE),IF(AND(A2>3000,A2<=3500,B2=10000,C2=5000,D2=3500),TRUE,FALSE),IF(AND(A2>3500,A2<=4000,B2=10000,C2=5000,D2=4000),TRUE,FALSE),IF(AND(A2>4000,A2<=5000,B2=10000,C2=7500,D2=5000),TRUE,FALSE),IF(AND(A2>5000,A2<=6000,B2=10000,C2=7500,D2=6000),TRUE,FALSE),IF(AND(A2>6000,A2<=7000,B2=15000,C2=10000,D2=7000),TRUE,FALSE),IF(AND(A2>7000,A2<=8000,B2=15000,C2=10000,D2=8000),TRUE,FALSE),IF(AND(A2>8000,A2<=9000,B2=15000,C2=10000,D2=9000),TRUE,FALSE),IF(AND(A2>9000,A2<10000,B2=15000,C2=12500,D2=10000),TRUE,FALSE),IF(AND(A2=10000,B2=20000,C2=15000,D2=10000),TRUE,FALSE),IF(A2>10000,TRUE,FALSE)),"","ECode1: Either the Last Gift Amount is not a number, or it's a negative number.") & IF(ISBLANK(A2),"ECode2: You have a blank Last Gift Amount - that may be OK, but it may not too.","") & IF(AND(ISNUMBER(A2),A2>10000),"ECode3: The Last Gift is greater than $10,000. This is likely OK, but should be noted.","") & IF(AND(NOT(ISNUMBER(A2)),NOT(ISBLANK(A2))),"ECode1: Either the Last Gift Amount is not a number, or it's a negative number.","") & IF(AND(A2>=0,A2<=10000,ISNUMBER(A2)),"OK!",""){enter}
:*:tabret`t::{CTRLDOWN}h{CTRLUP}\t{Tab}\r\n{ALTDOWN}ta{ALTUP}{ESCAPE}
:*:prfsel`t::(Not Empty([Proofs])) AND {ENTER}([Proofs] Excludes "NOT A PROOF")
*/

WinWait, 99999MB-BetaTemplateAdjust_2017_10-20, 
IfWinNotActive, 99999MB-BetaTemplateAdjust_2017_10-20, , WinActivate, 99999MB-BetaTemplateAdjust_2017_10-20, 
WinWaitActive, 99999MB-BetaTemplateAdjust_2017_10-20, 
MouseClick, left,  972,  22
Sleep, 100


; Start of Script, RUN FROM BASE MM2010 WINDOW!
Send, {CTRLDOWN}a{CTRLUP}
WinWait, Duplicate Records Processing, 
IfWinNotActive, Duplicate Records Processing, , WinActivate, Duplicate Records Processing, 
WinWaitActive, Duplicate Records Processing, 
Send, {ALTDOWN}n{ALTUP}
WinWait, Indexes, 
IfWinNotActive, Indexes, , WinActivate, Indexes, 
WinWaitActive, Indexes, 
Send, {ALTDOWN}n{ALTUP}
WinWait, Index Name, 
IfWinNotActive, Index Name, , WinActivate, Index Name, 
WinWaitActive, Index Name, 
Send, zclf2{ENTER}
WinWait, Indexes, 
IfWinNotActive, Indexes, , WinActivate, Indexes, 
WinWaitActive, Indexes, 
Send, {ALTDOWN}xp{ALTUP}


; Below chunk needs to be condensed and fixed - it's going to just send a chunk of text out.
/* WinWait, INDEXES.txt [L:\] - Notepad2-mod, 
IfWinNotActive, INDEXES.txt [L:\] - Notepad2-mod, , WinActivate, INDEXES.txt [L:\] - Notepad2-mod, 
WinWaitActive, INDEXES.txt [L:\] - Notepad2-mod, 
MouseClick, right,  40,  104
Sleep, 100
MouseClick, left,  27,  112
Sleep, 100
Send, {CTRLDOWN}c{CTRLUP}
WinWait, Indexes, 
IfWinNotActive, Indexes, , WinActivate, Indexes, 
WinWaitActive, Indexes, 
MouseClick, left,  185,  302
Sleep, 100
Send, {CTRLDOWN}v{CTRLUP}
*/
SendInput {Raw}



Send, {ALTDOWN}n{ALTUP}
WinWait, Index Name, 
IfWinNotActive, Index Name, , WinActivate, Index Name, 
WinWaitActive, Index Name, 
Send, xz{BACKSPACE}{BACKSPACE}zlf2{ENTER}
WinWait, Indexes, 
IfWinNotActive, Indexes, , WinActivate, Indexes, 
WinWaitActive, Indexes, 
Send, {ALTDOWN}xp{ALTUP}


; Below chunk needs to be condensed and fixed - it's going to just send a chunk of text out.
WinWait, INDEXES.txt [L:\] - Notepad2-mod, 
IfWinNotActive, INDEXES.txt [L:\] - Notepad2-mod, , WinActivate, INDEXES.txt [L:\] - Notepad2-mod, 
WinWaitActive, INDEXES.txt [L:\] - Notepad2-mod, 
MouseClick, left,  31,  166
Sleep, 100
Send, {CTRLDOWN}c{CTRLUP}
WinWait, Indexes, 
IfWinNotActive, Indexes, , WinActivate, Indexes, 
WinWaitActive, Indexes, 
MouseClick, left,  150,  317
Sleep, 100
Send, {CTRLDOWN}v{CTRLUP}{ALTDOWN}n{ALTUP}
WinWait, Index Name, 
IfWinNotActive, Index Name, , WinActivate, Index Name, 
WinWaitActive, Index Name, 
Send, zlf{ENTER}
WinWait, Indexes, 
IfWinNotActive, Indexes, , WinActivate, Indexes, 
WinWaitActive, Indexes, 
Send, {ALTDOWN}xp{ALTUP}


; Below chunk needs to be condensed and fixed - it's going to just send a chunk of text out.
WinWait, INDEXES.txt [L:\] - Notepad2-mod, 
IfWinNotActive, INDEXES.txt [L:\] - Notepad2-mod, , WinActivate, INDEXES.txt [L:\] - Notepad2-mod, 
WinWaitActive, INDEXES.txt [L:\] - Notepad2-mod, 
MouseClick, left,  31,  228
Sleep, 100
Send, {CTRLDOWN}c{CTRLUP}
WinWait, Indexes, 
IfWinNotActive, Indexes, , WinActivate, Indexes, 
WinWaitActive, Indexes, 
MouseClick, left,  206,  333
Sleep, 100
Send, {CTRLDOWN}v{CTRLUP}{ALTDOWN}n{ALTUP}
WinWait, Index Name, 
IfWinNotActive, Index Name, , WinActivate, Index Name, 
WinWaitActive, Index Name, 
Send, zl2{ENTER}
WinWait, Indexes, 
IfWinNotActive, Indexes, , WinActivate, Indexes, 
WinWaitActive, Indexes, 
Send, {ALTDOWN}xp{ALTUP}


; Below chunk needs to be condensed and fixed - it's going to just send a chunk of text out.
WinWait, INDEXES.txt [L:\] - Notepad2-mod, 
IfWinNotActive, INDEXES.txt [L:\] - Notepad2-mod, , WinActivate, INDEXES.txt [L:\] - Notepad2-mod, 
WinWaitActive, INDEXES.txt [L:\] - Notepad2-mod, 
MouseClick, left,  23,  289
Sleep, 100
Send, {CTRLDOWN}c{CTRLUP}
WinWait, Indexes, 
IfWinNotActive, Indexes, , WinActivate, Indexes, 
WinWaitActive, Indexes, 
MouseClick, left,  189,  345
Sleep, 100
Send, {CTRLDOWN}v{CTRLUP}{ALTDOWN}n{ALTUP}
WinWait, Index Name, 
IfWinNotActive, Index Name, , WinActivate, Index Name, 
WinWaitActive, Index Name, 
Send, zdpl{ENTER}
WinWait, Indexes, 
IfWinNotActive, Indexes, , WinActivate, Indexes, 
WinWaitActive, Indexes, 
Send, {ALTDOWN}xp{ALTUP}


; Below chunk needs to be condensed and fixed - it's going to just send a chunk of text out.
WinWait, INDEXES.txt [L:\] - Notepad2-mod, 
IfWinNotActive, INDEXES.txt [L:\] - Notepad2-mod, , WinActivate, INDEXES.txt [L:\] - Notepad2-mod, 
WinWaitActive, INDEXES.txt [L:\] - Notepad2-mod, 
MouseClick, left,  23,  293
Sleep, 100
Send, {CTRLDOWN}c{CTRLUP}
WinWait, Indexes, 
IfWinNotActive, Indexes, , WinActivate, Indexes, 
WinWaitActive, Indexes, 
MouseClick, left,  201,  309
Sleep, 100
Send, {CTRLDOWN}v{CTRLUP}{ALTDOWN}n{ALTUP}
WinWait, Index Name, 
IfWinNotActive, Index Name, , WinActivate, Index Name, 
WinWaitActive, Index Name, 
Send, zcl{ENTER}
WinWait, Indexes, 
IfWinNotActive, Indexes, , WinActivate, Indexes, 
WinWaitActive, Indexes, 
Send, {ALTDOWN}xp{ALTUP}


; Below chunk needs to be condensed and fixed - it's going to just send a chunk of text out.
WinWait, INDEXES.txt [L:\] - Notepad2-mod, 
IfWinNotActive, INDEXES.txt [L:\] - Notepad2-mod, , WinActivate, INDEXES.txt [L:\] - Notepad2-mod, 
WinWaitActive, INDEXES.txt [L:\] - Notepad2-mod, 
MouseClick, left,  28,  228
Sleep, 100
Send, {CTRLDOWN}c{CTRLUP}
WinWait, Indexes, 
IfWinNotActive, Indexes, , WinActivate, Indexes, 
WinWaitActive, Indexes, 
MouseClick, left,  202,  336
Sleep, 100
Send, {CTRLDOWN}v{CTRLUP}{ALTDOWN}n{ALTUP}
WinWait, Index Name, 
IfWinNotActive, Index Name, , WinActivate, Index Name, 
WinWaitActive, Index Name, 
Send, z2{ENTER}
WinWait, Indexes, 
IfWinNotActive, Indexes, , WinActivate, Indexes, 
WinWaitActive, Indexes, 
Send, {ALTDOWN}xp{ALTUP}


; Below chunk needs to be condensed and fixed - it's going to just send a chunk of text out.
WinWait, INDEXES.txt [L:\] - Notepad2-mod, 
IfWinNotActive, INDEXES.txt [L:\] - Notepad2-mod, , WinActivate, INDEXES.txt [L:\] - Notepad2-mod, 
WinWaitActive, INDEXES.txt [L:\] - Notepad2-mod, 
MouseClick, left,  15,  288
Sleep, 100
Send, {CTRLDOWN}c{CTRLUP}
WinWait, Indexes, 
IfWinNotActive, Indexes, , WinActivate, Indexes, 
WinWaitActive, Indexes, 
MouseClick, left,  198,  345
Sleep, 100
Send, {CTRLDOWN}v{CTRLUP}


; End of Script, exits back to "Indexes"
Send, {ALTDOWN}o{ALTUP}{ENTER}
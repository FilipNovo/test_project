#NoEnv
#SingleInstance, Force
#Include, spinner.ahk



Gui,  -Caption
Gui, Add, Text, w50 h50 x0 y0 +HWNDhTXT1
Gui, Show, Center w60 h60
WinSet, TransColor, FFFFFF 0
Spinner1 := new Spinner(cfg1, hTXT1)


Spinner1.Start()
Sleep, 3000
Spinner1.Stop()
Sleep, 3000
Spinner1.Start()


	
Return
GuiClose:
ExitApp

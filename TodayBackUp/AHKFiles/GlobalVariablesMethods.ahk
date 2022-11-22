#NoEnv
#Include %A_ScriptDir%\IncludesFiles.ahk
#SingleInstance, force

;Programs handeling
OutputTestCasesResults(testFileLocaion){
   CygnetTestCasesLookUptable := new CygnetTestCasesLookUptable(testFileLocaion)
   return CygnetTestCasesLookUptable
}

CygnetTestCasesTable(){
    destination := false
	while (!destination){
	FileSelectFile, destination, 1, D:\, Select File,
	}
   TestFileLocaion := destination
   return TestFileLocaion
}

runAndWait(cygnet, program,location,testName){
   RunWait, %program%  %location%\%testName%.ahk 
   objShell := ComObjCreate("WScript.Shell")
   objEnv := objShell.Environment("Volatile")
   cygnet.SetResultByTestName(testName, objEnv.Item("result"))
   cygnet.SetTimeByTestName(testName)
   return
}

runAndWaitSpinner(program,location,testName){
   spinner := addSpinner()
   RunWait, %program%  %location%\%testName%.ahk 
   return
}

runAndWaitNew(cygnet, program,location,testName){
   RunWait, %program%  %location%\%testName%.ahk 
   objShell := ComObjCreate("WScript.Shell")
   objEnv := objShell.Environment("Volatile")
   return objEnv.Item("result")
}

RunPowershellCommand(command){
   RunWait PowerShell.exe -ExecutionPolicy Bypass -Command %command%,,hide
}

openProgram(ProgramName, executionFile, programLocation){
   ControlSend, , {F5}, ahk_class Progman
   Sleep 500
   Send #d
   position := getDeskIconPosition(ProgramName)
   if (position.1 and position.2){
      Sleep 500
      MouseMove % position[1],% position[2], 100
      Send, {Click 2}
   }else  {
      Run, *RunAs %executionFile%, %programLocation%
   }
}


;Basic functions
GetCurrentTime(){
   FormatTime, OutputVar, R L0x0009
   return OutputVar
}

GetCurrentTimeNumbers(){
   FormatTime, OutputVar, R L0x0009, ddMMyyyyhhmmss
   return OutputVar
}

splitToArray(string, seprator){
   output := []
   Loop, Parse, string, %seprator%
   {
      output.push(A_LoopField)
   }
   return output
}

LookUptableLength(lookUpTable){
   return  lookUptable.MaxIndex()
}

chooseCygnetElementsLookupTable(){
   loc := false
	while (!loc){
	FileSelectFile, loc, 1, D:\, Select File,
	}
   fileLoc := loc
   return loc
}

Obj2String(Obj,FullPath:=1,BottomBlank:=0){
	static String,Blank
	if(FullPath=1)
		String:=FullPath:=Blank:=""
	if(IsObject(Obj)){
		for a,b in Obj{
			if(IsObject(b))
				Obj2String(b,FullPath "." a,BottomBlank)
			else{
				if(BottomBlank=0)
					String.=FullPath "." a " = " b "`n"
				else if(b!="")
					String.=FullPath "." a " = " b "`n"
				else
					Blank.=FullPath "." a " =`n"
			}
	}}
	return String Blank
}

m(x*){
	static List:={BTN:{OC:1,ARI:2,YNC:3,YN:4,RC:5,CTC:6},ico:{X:16,"?":32,"!":48,I:64}},Msg:=[]
	static Title
	List.Title:="AutoHotkey",List.Def:=0,List.Time:=0,Value:=0,TXT:="",Bottom:=0
	WinGetTitle,Title,A
	for a,b in x{
		Obj:=StrSplit(b,":"),(Obj.1="Bottom"?(Bottom:=1):""),(VV:=List[Obj.1,Obj.2])?(Value+=VV):(List[Obj.1]!="")?(List[Obj.1]:=Obj.2):TXT.=(IsObject(b)?Obj2String(b,,Bottom):b) "`n"
	}
	Msg:={option:Value+262144+(List.Def?(List.Def-1)*256:0),Title:List.Title,Time:List.Time,TXT:TXT}
	Sleep,120
	/*
		SetTimer,Move,-1
	*/
	MsgBox,% Msg.option,% Msg.Title,% Msg.TXT,% Msg.Time
	/*
		SetTimer,ActivateAfterm,-150
	*/
	for a,b in {OK:Value?"OK":"",Yes:"YES",No:"NO",Cancel:"CANCEL",Retry:"RETRY"}
		IfMsgBox,%a%
			return b
	return
	Move:
	TT:=List.Title " ahk_class #32770 ahk_exe AutoHotkey.exe"
	WinGetPos,x,y,w,h,%TT%
	WinMove,%TT%,,2000,% Round((A_ScreenHeight-h)/2)
	/*
		ToolTip,% A_ScriptFullPath
		USE THIS TO SAVE LAST POSITIONS FOR MSGBOX'S
	*/
	return
}

ReadFileSearchText(fileLoc, searchString){
   line := ""
   Loop, Read, %fileLoc%
   {
      If InStr(A_LoopReadLine, searchString)
      If !line
      {
         line := A_LoopReadLine
         Continue ;second concerned line
      } else  {
         line .= "`r`n" . A_LoopReadLine
         Break
      }}
   return line
}

addTextFile(location, data){
   FormatTime, mcurrentTime, %A_Now% L1033, yyyy-MM-dd HH-mm-ss
   FileAppend, %mcurrentTime%:  %data%`n, %location%
   return
}

SearchWordInString(Line, word){
	if(InStr(Line, word))
		return true
	else
		return false
}

Sleep(seconds){
   Sleep, % seconds * 1000
   return
}

;Window handeling
WinWait(window, timeout){
   WinWait, %window%,, %timeout%
   if ErrorLevel
      return false
   else
      return true
}

WinExistActivate(window){
   IfWinExist, %window%
   {
       WinRestore
       WinActivate
       WinMove, %window%, , ,  , 300, 300
       return true
   } else 
      return false
}

WinExistActivateClose(window){
   IfWinExist, % window
   {
       WinRestore
       WinActivate
       WinClose
       return true
   } else 
      return false
}

winExistClose(windowName){
if WinExist(windowName){
      WinClose
      return true
} else {
   return false
}
}

CloseActiveWindow(){
   WinClose, A
}

WinActivate(window, ahkClassahkClass){
   WinActivate, %window% ahk_class %ahkClass%
   WinMove, %window%, , ,  , 300, 300
}

WinClose(window, ahkClass){
   WinClose, %window% ahk_class %ahkClass%
}

;System handeling
isMute(){
   SoundGet, MuteState, Master, Mute
   if(MuteState="On"){
      return true
   }else  {
      return false
   }
}

RefreshDesktop(){
   ControlSend, , {F5}, ahk_class Progman
}

ChangeResolution(Screen_Width := 1920, Screen_Height := 1080, Color_Depth := 32)
{
	VarSetCapacity(Device_Mode,156,0)
	NumPut(156,Device_Mode,36) 
	DllCall( "EnumDisplaySettingsA", UInt,0, UInt,-1, UInt,&Device_Mode )
	NumPut(0x5c0000,Device_Mode,40) 
	NumPut(Color_Depth,Device_Mode,104)
	NumPut(Screen_Width,Device_Mode,108)
	NumPut(Screen_Height,Device_Mode,112)
	Return DllCall( "ChangeDisplaySettingsA", UInt,&Device_Mode, UInt,0 )
}

getDeskIconPosition(var)
{
   Critical
   static MEM_COMMIT := 0x1000, PAGE_READWRITE := 0x04, MEM_RELEASE := 0x8000
   static LVM_GETITEMPOSITION := 0x00001010, LVM_SETITEMPOSITION := 0x0000100F, WM_SETREDRAW := 0x000B
   ControlGet, hwWindow, HWND,, SysListView321, ahk_class Progman
   if !hwWindow ; #D mode
   ControlGet, hwWindow, HWND,, SysListView321, A
   IfWinExist ahk_id %hwWindow% ; last-found window set
   WinGet, iProcessID, PID
   hProcess := DllCall("OpenProcess"    , "UInt",   0x438           ; PROCESS-OPERATION|READ|WRITE|QUERY_INFORMATION
   , "Int", FALSE           ; inherit = false
   , "ptr", iProcessID)
   if hwWindow and hProcess
   {
      ControlGet, list, list,Col1   
      ;StringGetPos, pos, list, %var%
      if !coords
      {
         VarSetCapacity(iCoord, 8)
         pItemCoord := DllCall("VirtualAllocEx", "ptr", hProcess, "ptr", 0, "UInt", 8, "UInt", MEM_COMMIT, "UInt", PAGE_READWRITE)
          
         Loop, Parse, list, `n
         {
            if ( A_LoopField = var ){
               SendMessage, %LVM_GETITEMPOSITION%, % A_Index-1, %pItemCoord%
               DllCall("ReadProcessMemory", "ptr", hProcess, "ptr", pItemCoord, "UInt", &iCoord, "UInt", 8, "UIntP", cbReadWritten)
               x .=(NumGet(iCoord,"Int")) + 10
               y .=((Numget(iCoord, 4,"Int"))) + 10
               ret := [x,y]
               break
            }}
          
         DllCall("VirtualFreeEx", "ptr", hProcess, "ptr", pItemCoord, "ptr", 0, "UInt", MEM_RELEASE)
      }}
   DllCall("CloseHandle", "ptr", hProcess)
   return ret
}

;Click handeling
writeWord(word){
send, %word%
}

ClickActiveWindowControl(windowControl){
   MouseGetPos, , , WhichWindow,
   ControlGetPos, x, y, w, h, %windowControl%, ahk_id %WhichWindow%
   MouseClick, left, % x + w/2, % y + h/2, , 50
   return
}

DragClickActiveWindowControl(windowControl){
  MouseGetPos, , , WhichWindow,
  ControlGetPos, x, y,w ,h , %windowControl%, ahk_id %WhichWindow%
  MouseClickDrag, left, %x%, % y + h/2, % x+ w/2, % y + h/2, 50
  return
}

MoveMouseDragClick(x,y,w,h){
  MouseClickDrag, left, %x%, % y + h/2, % x+ w/2, % y + h/2, 50
  return
}

MoveMouseToWindowControl(windowControl){
   MouseGetPos, , , WhichWindow,
   ControlGetPos, x, y, w, h, %windowControl%, ahk_id %WhichWindow%
   MouseMove, % x , % y 
   return
}

MoveMouseClick(x,y){
   MouseMove, x , y 
   MouseClick, left, x, y
}

MoveMouseClickElement(element){
   MoveMouseClick(element[1],element[2])
 }
 
MoveMouseDragClickElement(element){
   x := element[1] - (element[3]/2)
   y := element[2] - (element[4]/2)
   MouseClickDrag, left, %x%, % y + element[4]/2, % x+ element[3]/2, % y + element[4]/2, 50
return
}
MoveMouseDragClickElementRandom(element, dragValue){
   x := element[1] - (element[3]/2)
   y := element[2] - (element[4]/2)
   if(dragValue = 0){
      MouseClickDrag, left, % x+ element[3]/2, % y + element[4]/2, %x%, % y + element[4]/2, 50
      return
   }
   MouseClickDrag, left, %x%, % y + element[4]/2, % x+ dragValue, % y + element[4]/2, 50
return
}
 
MoveMouseElement(element){
   MoveMouse(element[1],element[2])
}
 
 
pressRandomKey(){
   Random, oVar, 33, 122
   key := % Chr(oVar)
   Send,{%key%}
}

 
clickKeyboardButton(count, key){
   loop, %count%
   {
      Send,{%key%}
      Sleep(1)
   }
}

clickSelectAll(){
   Send, ^a
}

MoveMouse(x,y){
   MouseMove, x , y 
}

;Image handeling
getTextInImage(element){
   xStart := element[1]
   yStart := element[2]
   xEnd := element[1] + element[3]
   yEnd := element[2] + element[4]
   coordinates = "%xStart% %yStart% %yEnd% %yEnd%"
   RunWait, C:\Users\BEE Medic GmbH\Downloads\Test automation\autoHotKeyScripts\BEE_AHKTests\CygnetAHK\ahkLibraries\Capture2Text\Capture2Text_CLI.exe --screen-rect "1400 900 400 200" --clipboard
   MsgBox, %clipboard%
   shit := clipboard
   return shit
}

changeImagePosition(windowName, CygnetElementsLookUptable, index, consoleFile, console){
   x := CygnetElementsLookUptable.GetXCoordinates(index)
   y := CygnetElementsLookUptable.GetYCoordinates(index)
   w := CygnetElementsLookUptable.GetWidth(index) 
   h := CygnetElementsLookUptable.GetHeight(index)
   changed := false
   cancel := false
   element := CygnetElementsLookUptable.GetNameID(index)	
   ElementPosition := new ElementPosition(windowName, element, x, y, index, consoleFile, console)
   result := [0,0,index]
   if( x and y ) {
      ElementPosition.getElementInPosition()
      changed := ElementPosition.getPositionChanged()
   }
   else if(CygnetElementsLookUptable.GetCalssNNByNameID(element)) {
     if(WinExistActivate(windowName)){
         Sleep(1)
         coordinates := getCoordinatesByWindowControl(CygnetElementsLookUptable.GetCalssNNByNameID(element))
         x := coordinates[1] + coordinates[3]/2
         y := coordinates[2] + coordinates[4]/2
         ElementPosition := new ElementPosition(windowName, element, x, y, index, consoleFile, console)
         ElementPosition.getElementInPosition()
         changed := true
     }
   } else {
      MsgBox, You need to set %element% Position
      ElementPosition := new ElementPosition(windowName, element, 0, 0, index, consoleFile, console)
      ElementPosition.setElementInPosition(false)
      changed := true
   }

   if(changed){
      mess = Wait saving %element%  new position
      addTextFile(consoleFile, mess)
      spinner := addSpinner()
      x := ElementPosition.getCordinates()[1]
      y := ElementPosition.getCordinates()[2]
      w := ElementPosition.getSize()[1]
      h := ElementPosition.getSize()[2]
      CygnetElementsLookUptable.SetXCoordinates(index, x)
      CygnetElementsLookUptable.SetYCoordinates(index, y)
      CygnetElementsLookUptable.SetWidth(index, w)
      CygnetElementsLookUptable.SetHeight(index, h)
      mess = Done saving %element%  new position
      addTextFile(consoleFile, mess)
      console.minmize()
      console.restore()
      removeSpinner(spinner)
   }
   MoveMouse(x,y)
   if(ElementPosition.getCancel()){
      cancel := 1
   }
   value := ElementPosition.getIndex()
   result := [changed, cancel, value]
   Sleep, 1000
   return result
}

GetElementPosition(cygnetElementsLookupTable, element){
   x  := CygnetElementsLookUptable.GetXCordinateByNameID(element)
   y := CygnetElementsLookUptable.GetYCordinateByNameID(element)
   w := CygnetElementsLookUptable.GetWidthByNameID(element)
   h := CygnetElementsLookUptable.GetHeightByNameID(element)
   return [x,y,w,h]
}

checkImageElementsPosition(windowName, cygnetElementsLookUptable, consoleFile, console){
   changed := false 
   start := 1
   end := CygnetElementsLookUptable.ExcelArray.MaxIndex()   
   while(start <= end)
   {
      printElements(cygnetElementsLookUptable.GetListOfElements(start),start)
      if( WinExist(windowName)){
         elementPoistionChanged := changeImagePosition(windowName ,CygnetElementsLookUptable, start, consoleFile, console)
         if(elementPoistionChanged[1]){
            changed := true
         }
         if(elementPoistionChanged[2]){
            return changed
         }
         if(elementPoistionChanged[3] != start){
            start := elementPoistionChanged[3]
            continue
         }
         start += 1
      } else {
         m("Please open cygnet first")
         return
      }
   }
   return changed
}

getCoordinatesByWindowControl(windowControl){
   MouseGetPos, , , WhichWindow,
   ControlGetPos, x, y, w, h, %windowControl%, ahk_id %WhichWindow%
   ret := [x,y,w,h]
   return ret 
}

printElements(arrayOfElements, index) {
    if(index=arrayOfElements.MaxIndex()){
      ;Gui,3: Destroy
      winExistClose("Elements List")
      return
   }
   Gui,3: Destroy
   Gui, 3:-MinimizeBox +HWNDmsg +hwndHGUI +Resize
   Gui, 3:font, s10 w700 cblack, Segoe UI
   Loop % arrayOfElements.MaxIndex() {
      element =  %A_Index%
      elementName := arrayOfElements[A_Index]
      Gui, 3: Add, Text, , %element%. %elementName%
   }
   Gui, 3:+Border
   ;~ Gui, 3: Show, x0 y0 AutoSize, Elements List
  ; Create ScrollGUI1 with both horizontal and vertical scrollbars and scrolling by mouse wheel
   Global SG1 := New ScrollGUI(HGUI, 400, 400, "+Resize +LabelGui1", 3, 4)
   ; Show ScrollGUI1
   SG1.Show("ScrollGUI1", "y0 x0")
   return
}

compareImages(img1, img2) {
	pToken := Gdip_Startup() 
	image1 := Gdip_CreateBitmapFromFile(img1) 
	image2 := Gdip_CreateBitmapFromFile(img2)
	RET := Gdip_ImageSearch(image1,image2,LIST,0,0,0,0,0)
	Gdip_DisposeImage(image1)
	Gdip_DisposeImage(image2)
	return RET
}

takeScreenShot(imgName, element){
   Sleep(2)
   pToken := Gdip_Startup()
   element[1] := Round(element[1] - element[3]/2)
   element[2] := Round(element[2]- element[4]/2)
   imgName.= ".png"
   WinGetActiveStats, Title, Width, Height, X, Y
   screenShot := Gdip_BitmapFromScreen(element[1]+X "|" element[2]+Y "|" element[3] "|" element[4])
   Gdip_SaveBitmapToFile(screenShot, imgName)
   Gdip_ShutDown(pToken)
   return imgName  
}

deleteScreenShot(imgName){
   FileDelete, ./%imgName%
}

; Spinner handler

addSpinner(){
   Gui, Spinner: Destroy
   Gui, Spinner:  -Caption
   Gui, Spinner: Add, Text, w50 h50 x0 y0 +HWNDhTXT1
   Gui, Spinner: Show, Center w60 h60
   spinner := new Spinner(cfg1, hTXT1)
   spinner.Start()
   return spinner
}

removeSpinner(spinner){
   spinner.__Delete()
   Gui, Spinner: Destroy
   return
}

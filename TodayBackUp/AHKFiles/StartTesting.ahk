#SingleInstance, force
#WinActivateForce
SetWorkingDir %A_ScriptDir%
#Include %A_ScriptDir%\IncludesFiles.ahk
;Fix Resolution
;ChangeResolution(1920, 1080)

;import constat variables
consoleFile := objEnv.Item("consoleFile")
console := new BensConsole(consoleFile)
addTextFile(consoleFile, "Main application script started")
objShell := ComObjCreate("WScript.Shell")
objEnv := objShell.Environment("Volatile")
windowName := objEnv.Item("CYGNETHD_WINDOW_NAME")
cygnetTestCases := objEnv.Item("CygnetTestCasesDefaultFile")
addTextFile(consoleFile, "Creating object from Test cases file to use and copy of it to write the result of running excel file")

cygnet := OutputTestCasesResults(cygnetTestCases)
addTextFile(consoleFile, "Done Creating object from Test cases file to use and copy of it to write the result of running excel file")
cygnetElementsLookupTableFileLoc := objEnv.Item("CygnetElementsLookupTableDefaultFile")
addTextFile(consoleFile, "Creating object from elements Lookup excel table")
CygnetElementsLookUptable := new CygnetElementsLookupTable(cygnetElementsLookupTableFileLoc)
addTextFile(consoleFile, "Done Creating object from elements Lookup table")

;Global Attributes
result := ""
testNumber := 0
testName := ""
chossenTests := []
pause := ".\Images\play_down.png"
changeElementPosition :=".\Images\changeElementPosition.png"
play := ".\Images\pause_over.png"
consolePic := ".\Images\console.png"
noOfTests := LookUptableLength(cygnet.ExcelArray)
testParentprev := ""
currentParentID :=""
treeData := []
pauseClicked = 0

winExistClose("TestApplication.ahk")


;Gui Elements
Gui, +Resize
Gui, Add, Picture, hwndBut0  gShowConsole, %consolePic% 
Gui, Margin, 150
Gui, Add, Picture, hwndBut1 vpausePic gplay  ys, %pause%
Gui, Add, Picture, hwndBut2 xp yp wp hp vplayPic gpause Hidden , %play%
Gui, Add, Picture, hwndBut3 gChangeElementPosition ys, %changeElementPosition% 
Gui, Add, TreeView, section xs checked gTV_Clicked vTV_Clicked HwndTV_Hwnd AltSubmit 
treeStructure()
Gui, Show, Center  w450 h200
WinGetTitle, Title, A
if(Title = "StartTesting.ahk"){
	OnMessage(0x200, "Hover")
}
console.log("Terminal Started")
console.minimize()
console.restore()


return

;Gui controls
ChangeElementPosition:
{
	global CygnetElementsLookUptable
	global cygnetElementsLookupTableFileLoc
	global windowName
	global consoleFile
	global console
	

	addTextFile(consoleFile, "Elements checking gui is open")
	console.minmize()
	console.restore()
	ret := checkImageElementsPosition(windowName, CygnetElementsLookUptable, consoleFile, console)
	if(ret){
		CygnetElementsLookUptable := new CygnetElementsLookupTable(cygnetElementsLookupTableFileLoc)
	}
	return
}

ShowConsole:
{
	global consoleFile
	global consoleShown
	global objEnv
	global console
	
	count = 0
	Loop, Read, %consoleFile%
	{
		count++
		break
	}
	
	if(count = 0) {
		addTextFile(consoleFile, "No errors till now")
	}
	console.minmize()
	console.restore()
	return
}

play:
{
	global testNumber
	global chossenTests
	global console
	global consoleFile
	global pauseClicked
	global objEnv

    if(testNumber >= chossenTests.MaxIndex()){
		testNumber = 0 
	}
	
	while(testNumber < chossenTests.MaxIndex() && pauseClicked != 1){
		testNumber+=1
		testName := chossenTests[testNumber]
		objEnv.Item("testName") := testName
		testName .= " Started"
		addTextFile(consoleFile, testName)
		console.minmize()
		console.restore()
		PlayOn()
		PlayTest()
		testName := chossenTests[testNumber]
		testName .= " Ended"
		addTextFile(consoleFile, testName)
		console.minmize()
		console.restore()
	}
	pauseClicked = 0 
	PauseOn()
	return
}

pause:
{   
	global testNumber
	global chossenTests
	global consoleFile
	global console
	global objEnv
	global pauseClicked
	
	testName := chossenTests[testNumber]
	objEnv.Item("testName") := testName
	pauseClicked = 1 
    StopTest()
	PauseOn()
	testName .= " Stopped"
	addTextFile(consoleFile, testName)
	console.minmize()
	console.restore()
	
	return	
}

TV_Clicked:
{
	if(A_GuiEvent <> "Normal")
		return
	else {
	global cygnet
	global testNumber
	global noOfTests
	global chossenTests := []
	TV_Modify(A_EventInfo, "Select")
	
	TV_GetText(OutputVar, A_EventInfo)
	CheckRecursive(A_EventInfo)
	CheckRecursive(TV_ID)
	{
		If TV_GetChild(TV_ID) ; This is a parent with children.
		{
			RetrievedText := TV.GetText(ItemID)
			Status = 0
			If TV_Get(TV_ID, "Check") 			; The parent is checked
				Status = 1
			ChildID := TV_GetChild(TV_ID) 		; Get first child
			Loop							; Loop until no more children, then break
			{
				If (ChildID = 0)
					break
				TV_Modify(ChildID, "Check" . Status)
				CheckRecursive(ChildID)           ; Call myself on every childID. This makes me recurse all the nodes.
				ChildID := TV_GetNext(ChildID)
			}
		}
	}
	isParentUpdate()
Return
}
}

GuiSize:  ; Expand/shrink the ListView and TreeView in response to Client's resizing of window.
if (A_EventInfo = 1)  ; The window has been minimized. No action needed.
    return
; Otherwise, the window has been resized or maximized. Resize the controls to match.
GuiControl, Move, TV_Clicked, % "H" . (A_GuiHeight - 90)  ; -30 for StatusBar and margins.
GuiControl, Move, TV_Clicked, % "W" . (A_GuiWidth - 20)
return

GuiClose:
{
	for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
		if (process.Name = "EXCEL.EXE")
			process, close, % process.processId
	ExitApp
}


;Functions
PlayTest(){
	global testNumber
	global cygnet
	global chossenTests
	global console 
	
	if(chossenTests.MaxIndex() > 0 ){
		if(testNumber=0){
			Gui, Starter: Destroy
			location =  "%A_ScriptDir%"
			program = %A_AhkPath%
			runAndWait(cygnet,program,location,"TestApplication")
		} else {
			testName := chossenTests[testNumber]
			location =  "%A_ScriptDir%"
			program = %A_AhkPath%
			runAndWait(cygnet,program,location,testName)
			return
			}
	}else {
	m("Please selecte desired TestCases first")
}
	return
}

StopTest(){
	global testNumber
	global cygnet
	global chossenTests
	
	testName := chossenTests[testNumber]
	fullScriptPath = %A_ScriptDir%\%testName%.ahk  ; edit with your full script path
	DetectHiddenWindows, On 
	WinKill  %fullScriptPath% 
	if(testNumber > 0){
		testNumber-=1
	}
}

PlayOn(){
	GuiControl, show, playPic
	GuiControl, hide, pausePic
}

PauseOn(){	
	GuiControl, hide, playPic
	GuiControl, show, pausePic
	return
}

; Tree Functions
isParentUpdate(){
	global testNumber
	global cygnet
	global noOfTests
	global treeData
	global chossenTests := []
	Loop, % treeData.MaxIndex()
	{
		children := treeData[A_Index][2]
		Loop, % children.MaxIndex()
		{
			if TV_Get(children[A_Index][2], "Checked")
			{
				chossenTests.push(children[A_Index][1])
			}
		}
	}
	noOfTests := chossenTests.MaxIndex()
	return
}


treeStructure(){
	global cygnet
	global noOfTests
	global treeData
	Loop, %noOfTests% {
		testName :=  cygnet.GetTestName(A_Index)
		children :=  splitToArray(cygnet.GetChildren(A_Index)," `,")
		childrenAdded := false
		alternativeWay := []
		i:=0
		bestBranch := 0
		Loop, % treeData.MaxIndex()
		{	
			j=0
			branch := treeData[A_Index][1]
			Loop, % children.MaxIndex()
			{
				if(children[A_Index] == branch[A_Index][2]) {
						j+=1
				}  else {
					break
					}
			}
			if(j == children.MaxIndex()){
				childrenAdded := true	
			    testID := TV_Add(testName, branch[children.MaxIndex()][1])
				treeData[A_Index][2].push([testName, testID])
				break
			}  else if(i<j) {
					i := j
					bestBranch := branch
				}
		}
		if(!childrenAdded and i > 0)
		{
			Loop, %i%
				alternativeWay.push(branch[A_Index])
			n := children.MaxIndex() - i
			Loop, %n%
			{
				testID := TV_Add(children[i+1], alternativeWay[i][1])
				parentID := testID
				alternativeWay.push([parentID, children[i+1]])
				i+=1
			}
			testID := TV_Add(testName, parentID)
			treeData.push([alternativeWay, [[testName, testID]]])
			childrenAdded := true
		}
		if(!childrenAdded)
		{
			newBranch := []
			parentID  :=  TV_Add(children[1])
			newBranch.push([parentID, children[1]])
			Loop, % children.MaxIndex()-1
			{
				testID := TV_Add(children[A_Index+1], parentID)
				newBranch.push([testID, children[A_Index+1]])
				parentID := testID
			}
			testID := TV_Add(testName, parentID)
			t := [newBranch,[[testName, testID]]]
			treeData.push(t)
		}
	}
	 return
}

	
Hover(){
	WinGetTitle, Title, A
	if(Title = "StartTesting.ahk"){
		cursor:=DllCall( "LoadCursor"
				,"UInt",""
				,"Int",32649)
		MouseGetPos,,,,control
		
		If control=Static1
		{
			Help := "Show Console"
			DllCall("SetCursor","UInt",cursor)
		}
		If control=Static2 
		{
			Help := "Play"
			DllCall("SetCursor","UInt",cursor)
		}
		If control=Static3 
		{
			Help := "Pause"
			DllCall("SetCursor","UInt",cursor)
		}
		If control=Static4 
		{
			Help := "Change element position"
			DllCall("SetCursor","UInt",cursor)
		}
		If control=SysTreeView321
		{
			;~ Help := "Learn & troubleshoot"
			DllCall("SetCursor","UInt",cursor)
		}
		
		ToolTip % Help
	}
}
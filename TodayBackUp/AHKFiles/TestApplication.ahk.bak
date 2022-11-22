#SingleInstance, force
#WinActivateForce
SetWorkingDir %A_ScriptDir%
#Include %A_ScriptDir%\IncludesFiles.ahk


consoleFile := objEnv.Item("consoleFile")
addTextFile(consoleFile, "Test Application started")

cygnet := ""
cygnetElementsLookupTableFileLoc := ""
objShell := ComObjCreate("WScript.Shell")
objEnv := objShell.Environment("Volatile")
testCases := objEnv.Item("CygnetTestCasesDefaultFile")
SplitPath, testCases, name
testCases := name
testElements := objEnv.Item("CygnetElementsLookupTableDefaultFile")
SplitPath, testElements, name
testElements := name
CloseAllExcelFiles()


Gui, Starter: New 
Gui, Starter:font, s10 w700 cblack, Segoe UI

Gui, Starter: Add,Text,vTestCases , %testCases%
Gui, Starter: Add, Button, ym w200 gTestCasesFile, Change Default TestCases File
Gui, Starter: Add,Text, vTestElements section xs, %testElements%
Gui, Starter: Add, Button,   ys w200 gTestElementsFile, Chane Default Elements File
Gui, Starter: Add, Button,  vStartTesting xm   w700 gStartTesting, Start testing

Gui, Starter: Show, AutoSize Center
addTextFile(consoleFile, "Test application started")
winExistClose("CygnetTestApplication.ahk")
return

TestCasesFile:
{
	global consoleFile
	global cygnet := CygnetTestCasesTable()
	
	addTextFile(consoleFile, "Starting to change TestCases file location")
	objEnv.Item("CygnetTestCasesDefaultFile") := cygnet
	SplitPath, cygnet, name
	GuiControl,,TestCases,%name%
	addTextFile(consoleFile, "Done changing TestCases file location")
	return
}
 
TestElementsFile:
{
	global consolFile
		
	addTextFile(consoleFile, "Starting to change ElementsLookupTable file location")
	cygnetElementsLookupTableFileLoc := chooseCygnetElementsLookupTable()
	objEnv.Item("CygnetElementsLookupTableDefaultFile") := cygnetElementsLookupTableFileLoc
	SplitPath, cygnetElementsLookupTableFileLoc, name
	GuiControl,,TestElements,%name%
	addTextFile(consoleFile, "Done changing ElementsLookupTable file location")

	return

}

StartTesting:
{
	global consolFile
    
	addTextFile(consoleFile, "Starting to run main application for testing")
	location =  "%A_ScriptDir%"
	program = %A_AhkPath%
	Gui, Starter: Destroy
	spinner := addSpinner()
	runAndWaitSpinner(program,location,"StartTesting")
	return
}

GuiClose:
{
	ExitApp
}

CloseAllExcelFiles()
{
	global consolFile
		
	addTextFile(consoleFile, "Closing all opened excel files to prevent excel errors")
	for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
		if (process.Name = "EXCEL.EXE")
			process, close, % process.processId
	addTextFile(consoleFile, "Done closing all opened excel files")

}
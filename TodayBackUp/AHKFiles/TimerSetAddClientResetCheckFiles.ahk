;Task 8 TimeSetAddClientResetCheckFiles
#SingleInstance, force
#Include %A_ScriptDir%\IncludesFiles.ahk
objShell := ComObjCreate("WScript.Shell")
objEnv := objShell.Environment("Volatile")
console := objEnv.Item("consoleFile")
testName := objEnv.Item("testName")
objEnv.Item("testResult") := ""


clickButton("PlayIcon")
clickButton("SearchAddClients")
writeNewUniqueClient()
clickButton("AddClient")
clickButton("StopIcon")
Sleep(20)
checkNewClientPdf()
checkNewClientXdf()
CloseActiveWindow()
                  

testResult := objEnv.Item("testResult")
objEnv.Item("result") := testResult
consoleR = %testName% %testResult%
addTextFile(console, consoleR)
ExitApp
return
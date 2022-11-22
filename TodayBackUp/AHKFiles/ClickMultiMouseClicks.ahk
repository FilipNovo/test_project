;Task 38 Click Mutliple mouse clicks
#SingleInstance, force
#Include %A_ScriptDir%\IncludesFiles.ahk
objShell := ComObjCreate("WScript.Shell")
objEnv := objShell.Environment("Volatile")
console := objEnv.Item("consoleFile")
testName := objEnv.Item("testName")
objEnv.Item("testResult") := ""


clickElementsRandomly(["StopIcon", "PlayIcon", "RecordIcon"], 0.2, 50)
clickButton("StopIcon")
clickButton("PlayIcon")
clickButton("StopIcon")

testResult := objEnv.Item("testResult")
objEnv.Item("result") := testResult
consoleR = %testName% %testResult%
addTextFile(console, consoleR)
ExitApp
return
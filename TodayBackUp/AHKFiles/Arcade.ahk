﻿;Task 10 Game Manuel
#SingleInstance, force
#Include %A_ScriptDir%\IncludesFiles.ahk
objShell := ComObjCreate("WScript.Shell")
objEnv := objShell.Environment("Volatile")
console := objEnv.Item("consoleFile")
testName := objEnv.Item("testName")
objEnv.Item("testResult") := ""


clickButton("Help")
clickButton("GameManuels")
clickManuel("Arcade/ArcadeHelp.pdf", 2)


testResult := objEnv.Item("testResult")
objEnv.Item("result") := testResult
consoleR = %testName% %testResult%
addTextFile(console, consoleR)
ExitApp
return

#SingleInstance, force
#Include %A_ScriptDir%\GlobalVariablesMethods.ahk

spinner := addSpinner()

currentTime = % GetCurrentTimeNumbers()
consoleFile = %A_ScriptDir%\console%currentTime%.txt
objEnv.Item("consoleFile") := consoleFile
addTextFile(consoleFile, "Application started")
location =  "%A_ScriptDir%"
program = %A_AhkPath%
runAndWait(cygnet,program,location,"TestApplication")

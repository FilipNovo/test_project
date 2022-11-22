;Task 8 check Mute Functionality
#SingleInstance, force
#WinActivateForce
#Include %A_ScriptDir%\IncludesFiles.ahk

;#Include IncludesFiles.ahk
objShell := ComObjCreate("WScript.Shell")
objEnv := objShell.Environment("Volatile")
windowName := objEnv.Item("CYGNET_WINDOW_NAME")
programName := objEnv.Item("PROGRAM_NAME")
programLocation := objEnv.Item("PROGRAM_LOCATION")
feedbackWindowName := objEnv.Item("FEEDBACK_WINDOW_NAME")
executionFile := objEnv.Item("EXECTUION_FILE")
windowClass := objEnv.Item("CYGNET_CLASS_NAME")
console := objEnv.Item("consoleFile")

cygnetElementsLookupTable := new CygnetElementsLookupTable(objEnv.Item("cygnetElementsLookupTableFileLoc"))

SignatureValidation(programLocation){
   global console
   command = "Get-AuthenticodeSignature -FilePath  %programLocation% | Format-List | Out-File -FilePath .\Validation.txt"
   RunPowershellCommand(command)
   line := False
   searchString:="Status"
   validationFile:= ".\Validation.txt"
   line := ReadFileSearchText(validationFile, SearchString)
   if(line){
      valid := SearchWordInString(Line, "valid")
   if(valid){
      return "Passed"
   }else  {
      addTextFile(console, "Cygnet version is not valid anymore not")
      return "Failed"
   }
   Sleep(1)
   command := "Remove-Item -Path .\Validation.txt"
   RunPowershellCommand(command)
   }else  {
      addTextFile(console, "Can not find Cygnet AuthenticodeSignature")
      result := "Failed"
   }
   addTextFile(console, "Can not find Cygnet AuthenticodeSignature")
}



AddNewClient(){
   global windowName
   global programLocation
   global cygnetElementsLookupTable
   global console

   if(WinExistActivate(windowName))
   {      
      addClient := GetElementPosition(cygnetElementsLookupTable, "AddClient")
      img1 := takeScreenShot("img1", addClient)
      
      goBack := GetElementPosition(cygnetElementsLookupTable, "GoBack")
      img2 := takeScreenShot("img2", goBack)
      
      ;~ searchAddClients := GetElementPosition(cygnetElementsLookupTable, "SearchAddClients")

      ;~ MoveMouseClickElement(searchAddClients)
      
      ;~ currentTime := GetCurrentTimeNumbers()
      ;~ uniqueClient = TestClient_%currentTime%

      ;~ writeWord(uniqueClient)
      ;~ Sleep, 500
      
      img3 := takeScreenShot("img3", addClient)

      if(!compareImages(img1, img3)){ 
         MoveMouseClickElement(addClient)
         sleep(1)
         checkSession = \Session\%uniqueClient%
         newClientFolder = %programLocation%%checkSession%
         IfNotExist, %newClientFolder%
         {
            addTextFile(console, "New folder was not created for the new Client")
            return "Failed"
         }
         IfExist, %programLocation%
         {
            img4 := takeScreenShot("img4", goBack)
            if(!compareImages(img2, img4)){
               return "Passed"
            } else{
               addTextFile(console, "After adding new Client, the Client was not opened automatically")
               return "Failed"
            }
         }

      } else{
         return "Add button not visible"
      }
      
   } else {
      return "Failed - Cygnet is not open"
   }
}

;~ clickPlayButtonTrial(){
   ;~ global windowName
   ;~ global cygnetElementsLookupTable
   ;~ global console
  ;~ ; if(lastMethode){
   ;~ if(WinExistActivate(windowName))
   ;~ { 
      ;~ playIcon := GetElementPosition(cygnetElementsLookupTable, "PlayIcon")
      ;~ img1 := takeScreenShot("img1", playIcon)
      ;~ MoveMouseClickElement(playIcon)
      ;~ Sleep(1)
      ;~ img2:= takeScreenShot("img2", playIcon)
      ;~ if(!compareImages(img1, img2))
      ;~ {
         ;~ FileDelete, img1.png
         ;~ FileDelete, img2.png
         ;~ return "passed"
      ;~ } else {
         ;~ FileDelete, img1.png
         ;~ FileDelete, img2.png
         ;~ return "play icon doesnot work correctly"
      ;~ }
   ;~ }
   ;~ else {
      ;~ ;write to the log
      ;~ ;right in the log its skipped and if it passed or not then update last methode value
      ;~ return "failed"
   ;~ }
      
;~ ;} right in the log its skipped and if it passed or not

;~ }

PauseSession(){
   result := clickPlayButton()
   if(result = "Passed"){
      result := clickPlayButton()
      return result
   }
   return result
}

clickPlayButton(){
   global windowName
   global cygnetElementsLookupTable
   global console
   
   if(WinExistActivate(windowName))
   {
      playTimer := GetElementPosition(cygnetElementsLookupTable, "playTimer")
      timerBefore := takeScreenShot("timerBefore", playTimer)
      play := GetElementPosition(cygnetElementsLookupTable, "PlayIcon")
      img1 := takeScreenShot("img1", play)
      play := GetElementPosition(cygnetElementsLookupTable, "PlayIcon")
      MoveMouseClickElement(play)
      Sleep(5)
      play := GetElementPosition(cygnetElementsLookupTable, "PlayIcon")
      img2 := takeScreenShot("img2", play)
      if(!compareImages(img1, img2)){
         playTimer := GetElementPosition(cygnetElementsLookupTable, "playTimer")
         timerAfter := takeScreenShot("timerAfter", playTimer)
         if(!compareImages(timerBefore, timerAfter)){
            result := "Passed"
         } else{
            addTextFile(console, "Time is not updated after cicking play button")
            result := "failed"
         }
      } else{
         addTextFile(console, "Play button doesnot work correctly")
         result := "failed"
      }
   } else{
      addTextFile(console, "Cygnet is not open")
      result := "failed"
   }
   deleteScreenShot(img1)
   deleteScreenShot(img2)
   deleteScreenShot(timerBefore)
   deleteScreenShot(timerAfter)
   return result
}

TimeUpdatecheck() {
   global windowName
   global cygnetElementsLookupTable
   global console

   
   if(WinExistActivate(windowName))
   { 
      Sleep(1)
      playTimer := GetElementPosition(cygnetElementsLookupTable, "playTimer")
      timerBefore := takeScreenShot("timerBefore", playTimer)


      playButton := clickPlayButton()
      if(SearchWordInString(playButton, "passed")){
         Sleep(5)
         timerafter := takeScreenShot("timerafter", playTimer)
         if(!compareImages(timerBefore, timerafter))
         {
            result := "Passed"
         } else {
            addTextFile(console, "On click play timer doesnot change")
            result := "Failed"
         }
      } else {
         result := playButton
      }
   }
   else {
      addTextFile(console, "Cygnet is not open")
      return "Failed"
   }
   FileDelete, timerBefore.png
   FileDelete, timerafter.png
   return result
}

MuteCygnet() {
   global windowName
   global cygnetElementsLookupTable
   if(WinExistActivate(windowName))
   {
      volume := GetElementPosition(cygnetElementsLookupTable, "Volume")
      MoveMouseClickElement(volume)
      Sleep(0.5)
      RefreshDesktop()
      Sleep(0.5)
      return CygnetIsMute()
   } else  {
      addTextFile(console, "Failed - Cygnet is not open")
      return "Failed"
   }
}

StopResetCheck(){
   global windowName
   global cygnetElementsLookupTable
   global console
   
   if(WinExistActivate(windowName))
   { 
      playTimer := GetElementPosition(cygnetElementsLookupTable, "playTimer")
      timerBeforeStop := takeScreenShot("timerBeforeStop", playTimer)
      Sleep(0.5)
      stopIcon := GetElementPosition(cygnetElementsLookupTable, "StopIcon")
      MoveMouseClickElement(stopIcon)
      Sleep(2)
      timerAfterStop := takeScreenShot("timerAfterStop", playTimer)
      if(!compareImages(timerBeforeStop, timerAfterStop)){
         result := "passed"
      } else {
         result := "on click stop timer doesnot change"
      }
      
   } else {
      return "Failed - Cygnet is not open"
   }
   FileDelete, timerBeforeStop.png
   FileDelete, timerAfterStop.png
   return result
}

DragVolumeSlider(){
   global windowName
   global cygnetElementsLookupTable
   global console
   
   if(WinExistActivate(windowName))
   {
     Volume_Slider := GetElementPosition(cygnetElementsLookupTable, "Volume_Slider")
     MoveMouseDragClickElement(Volume_Slider)
     Sleep(0.5)
     RefreshDesktop()
     Sleep(0.5)
     return CygnetIsNotMute()
   } 
   addTextFile(console, "Cygnet is not open")
   return "Failed"
}

checkFiles(location, files)
{
   global console

   numberOfFiles := files.MaxIndex()
   Loop, %numberOfFiles%
   {
      fileExtention := location
      fileExtention .= files[A_Index]
      if (not FileExist(fileExtention))
      {
         addTextFile(console, "Files was not created correctly in the new Client folder")
         return "Failed"
      } 
   }
   return "passed"
}

OpenCygnet(){
   global programName
   global executionFile
   global programLocation
   global windowName
   global feedbackWindow
   global feedbackWindowName
   global console
   
   openProgram(programName, executionFile, programLocation)
   cygnetWindow := WinWait(windowName, 3*60)
   if (!cygnetWindow)
   {
      return "Failed - Cygnet Takes More Than 3 minutes to Lunch!"
   } 
   feedbackWindow := WinWait(feedbackWindowName, 2*60)
   if (!feedbackWindow)
   {
      addTextFile(console, "Select Feedback Window Takes More Than 3 minutes to Lunch!")
      return "Failed"
   } 
   WinActivate(windowName, windowClass)
   
   return "Passed"
}

CloseCygnet()
{
   global windowName
   global console
   
   if(winExistClose(windowName))
      return "Passed"
   else {
      addTextFile(console, "Cygnet is not open ")
      return "Failed"
   }
      
}

checkReportTabIsOPen() {
   global windowName
   global cygnetElementsLookupTable
   global console
   
   if(WinExistActivate(windowName))
   {
   reportTab := GetElementPosition(cygnetElementsLookupTable, "Report_Tab")
   reportTab[2] -= 15

   MoveMouseElement(reportTab)
   
   MouseGetPos, x, y
   PixelGetColor, color, x, y, RGB
   if(color = "0x268DD3"){
      return "Passed"
   } else {
      addTextFile(console, "Report function is not highlighted by default after starting  Cygnet")
      return "Failed"
   }

   } 
   addTextFile(console, "Cygnet is not open")
   return "Failed"
}

CygnetIsMute()
{
   global console

  if (isMute()) {
      return "Passed"
   } else {
      addTextFile(console, "System should be mute but it is not")
      return "Failed"
      }
}

CygnetIsNotMute()
{
  global console
  if (!isMute()) {
      return "Passed"
   } else {
      addTextFile(console, "System should be not mute but it is")
      return "Failed"
      }
}

CloseOpenisMute()
{
   result := CloseCygnet()
   if(result = "Passed"){
      result := OpenCygnet()
      if(result = "Passed"){
         result := CygnetIsMute()
         if(result = "Passed") {
            return result
         } else {
            return result
         }
         
      } else {
         return result
      }
   }
   return result
}


ManuelsOpenCorrectly(manuelName, index){
   global windowName
   global cygnetElementsLookupTable
   global console 
   
   if(WinExistActivate(windowName))
   {
       help := GetElementPosition(cygnetElementsLookupTable, "Help")
       MoveMouseClickElement(help)
       clickKeyboardButton(2, "down")
       clickKeyboardButton(1, "right")
       clickKeyboardButton(index, "down")
       clickKeyboardButton(1, "Enter")
       Sleep(2)
      splitManuelName:=StrSplit(manuelName,"/")
      pdfFile := splitManuelName[splitManuelName.MaxIndex()]
       WinGetTitle, Title, A
       if(SearchWordInString(Title, pdfFile)){
       Send, !d
       Sleep 3
       Send, ^c
       openedPdf := clipboard
       if(SearchWordInString(openedPdf, manuelName)){
            return "Passed"
       } else {
         addTextFile(console, "Not the right pdf is opened")
         return "Failed"
       }
       } else {
         addTextFile(console, "Title of opened pdf manuel not correct")
         return "Failed"
       }
       
   } 
   addTextFile(console, "Cygnet is not open")
   return "Failed"
}

playFromToolBar(){
   global windowName
   global cygnetElementsLookupTable
   global console
   if(WinExistActivate(windowName))
   {
      session := GetElementPosition(cygnetElementsLookupTable, "Session")
      play := GetElementPosition(cygnetElementsLookupTable, "PlayIcon")
      img1 := takeScreenShot("img1", play)
      MoveMouseClickElement(session)
      clickKeyboardButton(1, "down")
      clickKeyboardButton(1, "Enter")
      Sleep(5)
      play := GetElementPosition(cygnetElementsLookupTable, "PlayIcon")
      img2 := takeScreenShot("img2", play)
      if(!compareImages(img1, img2)){
         result := "Passed"
      } else {
      addTextFile(console, "Start in session tap does not make session start")
      result := "Failed"
      }
   } else {
      addTextFile(console, "Cygnet is not open")
      result := "Failed"
   }
   deleteScreenShot(img1)
   deleteScreenShot(img2)
   return result
}

pauseFromToolBar(){
   global windowName
   global cygnetElementsLookupTable
   global console

   if(WinExistActivate(windowName))
   {
      session := GetElementPosition(cygnetElementsLookupTable, "Session")
      play := GetElementPosition(cygnetElementsLookupTable, "PlayIcon")
      img1 := takeScreenShot("img1", play)
      MoveMouseClickElement(session)
      clickKeyboardButton(3, "down")
      clickKeyboardButton(1, "Enter")
      Sleep(5)
      play := GetElementPosition(cygnetElementsLookupTable, "PlayIcon")
      img2 := takeScreenShot("img2", play)
      if(!compareImages(img1, img2)){
         result := "Passed"
      }else {
         addTextFile(console, "Pause in session tap does not make session pause")
         result := "Failed"
      }
   } else {
      addTextFile(console, "Cygnet is not open")
      result := "Failed"
   }
   deleteScreenShot(img1)
   deleteScreenShot(img2)
   return result
}

stopFromToolBar(){
   global windowName
   global cygnetElementsLookupTable
   global console

   if(WinExistActivate(windowName))
   {
      session := GetElementPosition(cygnetElementsLookupTable, "Session")
      play := GetElementPosition(cygnetElementsLookupTable, "PlayIcon")
      img1 := takeScreenShot("img1", play)
      MoveMouseClickElement(session)
      clickKeyboardButton(5, "down")
      clickKeyboardButton(1, "Enter")
      Sleep(5)
      if(WinExistActivate(windowName)){
         Sleep(2)
         play := GetElementPosition(cygnetElementsLookupTable, "PlayIcon")
         img2 := takeScreenShot("img2", play)
         if(!compareImages(img1, img2)){
            result := "Passed"
         } else{
         addTextFile(console, "Stop in session tap does not make session pause")
         result := "Failed"
         }
      } else {
      addTextFile(console, "Cygnet is not open")
      result := "Failed"
      }
 
   } else {
      addTextFile(console, "Cygnet is not open")
      result := "Failed"
   }
   deleteScreenShot(img1)
   deleteScreenShot(img2)
   return result
}

SessionToolbarOptions(){
   result := playFromToolBar()
   if(result = "Passed"){
      result := pauseFromToolBar()
   }
   return result
}
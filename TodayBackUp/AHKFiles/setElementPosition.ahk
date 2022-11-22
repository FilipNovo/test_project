#SingleInstance, force
#Include %A_ScriptDir%\IncludesFiles.ahk

class ElementPosition{

__New(windowName, elementName, coordinateX, coordinateY, index, consoleFile, console)
    {
		this.coordinateX := coordinateX
		this.coordinateY := coordinateY
		this.elementName := elementName
		this.index := index
		referenceImage =  %A_ScriptDir%\ReferenceAppImages\%elementName%.png
		this.referenceImage := referenceImage
		pToken := Gdip_Startup()
		imgRef := Gdip_CreateBitmapFromFile( this.referenceImage )                 
		this.width:= Gdip_GetImageWidth( imgRef )
		this.height:= Gdip_GetImageHeight( imgRef )
		this.windowName := windowName
		this.cancel := false
		this.change := false
		Gdip_ShutDown(pToken)
		this.consoleFile := consoleFile
		this.console := console
    }
	
	getPositionChanged(){
		return this.change
	}
	
	getCancel(){
		return this.cancel
	}
	
	getIndex(){
		return this.index
	}
	
	getCordinates(){
		return [this.coordinateX, this.coordinateY]
	}
	
	getSize(){
		return [this.width, this.height]
	}

	getElementInPosition(){
		IfWinNotExist, % this.windowName 
		{
			MsgBox, Cygnet is not active to check icons
			return
		} else IfWinNotActive, % this.windowName 
		{
			WinExistActivate(this.windowName)
			WinWaitActive, windowName,, 2
		}
		pToken := Gdip_Startup()
		WinGetActiveStats, Title, Width, Height, X, Y
		if(this.change){
			xStart := this.coordinateX - ( this.width/2)
			yStart := this.coordinateY - (this.height/2)
			screenShot := Gdip_BitmapFromScreen(xStart "|" yStart "|" this.width "|" this.height)
			this.coordinateX -= X
			this.coordinateY -= Y
		} else {
			xStart := this.coordinateX - ( this.width/2)
			yStart := this.coordinateY - (this.height/2)
			screenShot := Gdip_BitmapFromScreen(xStart+X "|" yStart+Y "|" this.width "|" this.height)
		}
		Gdip_SaveBitmapToFile(screenShot, "Element.png")
		this.MsgB("Element.png")
		Gdip_ShutDown(pToken)
		return
	}

	setElementInPosition(manually){
		IfWinNotExist, % this.windowName
		{
			MsgBox, Cygnet is not active to check icons
			return
		} else IfWinNotActive, % this.windowName 
		{
			WinExistActivate(this.windowName)
			WinWaitActive, windowName,, 1
		}
		GetRange(x,y,w,h,manually) 
		    if(manually){
				this.coordinateX := x + (w/2) 
				this.coordinateY := y + (h/2)
			} else {
				this.coordinateX := x
				this.coordinateY := y
			}
		this.change := true
		this.getElementInPosition()
		return
	}

	MsgB(img)
	{
		global Index := this.index
	    element := this.elementName
		Gui, 2:font, s10 w700 cblack, Segoe UI
		Gui, 2:-MinimizeBox +AlwaysOnTop +HWNDmsg
	    Gui, 2:Add,Text, section , Detected
		Gui, 2:Add, Picture,  xs, %img%
		Gui, 2:Add,Text, section ys, Reference
		Gui, 2:Add, Picture, xs , % this.referenceImage
		Gui, 2:Add,Text, section ys, Element index
		Gui, 2:Add, Edit, xs w100 vIndex, % this.index
		Gui, 2:Add,Text,  xm , Correct Position for %element% element?
		Gui, 2:Add,Button,  Default Section w50 Center gyes,Yes
		Gui, 2:Add, Button, ys w50 gno, No 
		Gui, 2:Add, Button, ys w200 gSetWH, Set position Manually
		Gui, 2:Add, Button, ys w200 gCancel, Cancel
		Gui, 2:Add,Text, section  xm  , 1. If its the right position click yes
		Gui, 2:Add,Text, section  xm  , 2. If its not click No and click in the middle of the right position
		Gui, 2:Add,Text, section  xm  , 3. If its not and you want to set position Manually click Set position Manually
		Gui, 2:Add,Text, section  xm  ,click Mouse Left Button and drag around the element and release, you can move the produced rectangel using left button after you are done click right click
		Gui, 2:Add,Text, section  xm  , 4. Click Cancel if you want to stop checking elements positions
		Gui, 2:Show, Center AutoSize
		Gui, 2:+Border
		WinWaitClose,ahk_id %msg%
		return
		
		yes:
		{
			GuiControlGet, newIndex,, Index
			if(newIndex and newIndex != this.index){
				this.index := newIndex
			}
			Gui,2: Destroy
			FileDelete, Element.png
			mess =  % this.elementName " position is right"
			addTextFile(this.consoleFile, mess)
			this.console.minmize()
			this.console.restore()
			return
		}	

		no: 
		{	
			Gui,2: Destroy
			mess =  % this.elementName " position needs to be re-set"
			addTextFile(this.consoleFile, mess)
			this.console.minmize()
			this.console.restore()
			this.setElementInPosition(false)
			return
		}
		
		SetWH:
		{
			Gui,2: Destroy
			addTextFile(this.consoleFile, "Re-Set element position mode")
			this.console.minmize()
			this.console.restore()
			this.setElementInPosition(true)
			return
		}
		
		Cancel:
		{
			Gui,2: Destroy
			this.cancel := true
			addTextFile(this.consoleFile, "Element postion gui is closed")
			this.console.minmize()
			this.console.restore()
			return
		}
	}

}

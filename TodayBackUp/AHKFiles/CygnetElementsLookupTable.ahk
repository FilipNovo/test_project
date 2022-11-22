#Include %A_ScriptDir%\IncludesFiles.ahk

class CygnetElementsLookupTable {

	  __New(ExcelSheet)
    {
		this.ExcelSheet := ExcelSheet
		this.ExcelHandler := new ExcelHandler(ExcelSheet)
		this.ExcelArray := this.ExcelHandler.ExcelArray

    }
	
	GetNameID(index){
		return this.ExcelArray[index].NameID
	}
	
	GetWidth(index){
		return this.ExcelArray[index].Width
	}
	
	GetHeight(index){
		return this.ExcelArray[index].Height
	}
	
	GetClassNN(index){
		return this.ExcelArray[index].ClassNN
	}
	
	
	GetXCoordinates(index){
		return this.ExcelArray[index].XCordinate
	}
	
	GetYCoordinates(index){
		return this.ExcelArray[index].YCordinate
	}
	
	
	SetWidth(index, width){
		this.ExcelHandler.WriteToExcelSheetWithLabels(width, index, 3, 1)
	}
	
	SetHeight(index, height){
		this.ExcelHandler.WriteToExcelSheetWithLabels(height, index, 4, 1)
	}

	
	SetXCoordinates(index, xcordinate){
		this.ExcelHandler.WriteToExcelSheetWithLabels(xcordinate, index, 5, 1)
	}
	
	SetYCoordinates(index, ycordinate){
		this.ExcelHandler.WriteToExcelSheetWithLabels(ycordinate, index, 6, 1)
	}
	
	GetRow(index){
		return this.ExcelArray[index]
	}
	
	GetListOfElements(index){
	    elmentNamesList := []
		Loop, % this.ExcelArray.MaxIndex() {
			if(index-1 >= A_index){
				elementChecked := this.ExcelArray[A_Index].NameID
				elementChecked .= "  OK"
				elmentNamesList.push(elementChecked)
			} else {
				elmentNamesList.push(this.ExcelArray[A_Index].NameID)
			}
		}
		return elmentNamesList
	}
	
	GetCalssNNByNameID(nameID){
		Loop, % this.ExcelArray.MaxIndex() {		
			if(this.ExcelArray[A_Index].NameID = nameID){
				return this.ExcelArray[A_Index].ClassNN
			}
		}
		return false
	}

    GetWidthByNameID(nameID){
		Loop, % this.ExcelArray.MaxIndex() {		
			if(this.ExcelArray[A_Index].NameID = nameID){
				return this.ExcelArray[A_Index].Width
			}
		}
		return false
	}
	
	  GetHeightByNameID(nameID){
		Loop, % this.ExcelArray.MaxIndex() {		
			if(this.ExcelArray[A_Index].NameID = nameID){
				return this.ExcelArray[A_Index].Height
			}
		}
		return false
	}

    GetXCordinateByNameID(nameID){
		Loop, % this.ExcelArray.MaxIndex() {		
			if(this.ExcelArray[A_Index].NameID = nameID){
				return this.ExcelArray[A_Index].XCordinate
			}
		}
		return false
	}
	
	  GetYCordinateByNameID(nameID){
		Loop, % this.ExcelArray.MaxIndex() {		
			if(this.ExcelArray[A_Index].NameID = nameID){
				return this.ExcelArray[A_Index].YCordinate
			}
		}
		return false
	}
	
}

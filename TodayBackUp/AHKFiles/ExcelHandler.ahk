#Include %A_ScriptDir%\IncludesFiles.ahk

#NoEnv

class ExcelHandler
{
	__New(FileName) {
		this.sheetNo := 0
		this.ExcelFile := FileName
		this.ExcelArray := this.ExcelToArray()
	}

	
	ExcelToArray()
	{
		arr := this.GetArr()
		ret := this.ArrWithLabels_To_AHKKeyValueArr(arr)
		return ret
	}
	
	GetArr()
	{
		fPath := this.GetFullPath(this.ExcelFile)
		if !wb {
			wb:= ComObjGet(fPath)
		}
		this.sheetNo :=  wb.Sheets.Count				
		sheet := wb.Sheets(this.sheetNo)
		arr := sheet.UsedRange.FormulaR1C1
		xlObj.Quit
		return arr
	}
	
	WriteToExcelSheetWithLabels(text, row, column, sheetNumber){

		WorkBookPath := this.GetFullPath(this.ExcelFile)
		objExcel := ComObjCreate("Excel.Application") 
		objWorkBook := objExcel.Workbooks.Open( WorkBookPath )
		objExcel.Visible := false

		objSheet := objWorkBook.Worksheets( sheetNumber )
		objSheet.Cells((row+1), column).Value := text

		objWorkBook.Save() 
		objWorkBook.Close()
		xlApp.DisplayAlerts := False
		xlApp.ActiveDocument.Close(False)  ; *** this is what fixed it for me in the past
		xlApp.Quit
		return
		
	}

	ArrWithLabels_To_AHKKeyValueArr(arr)
	{
		ret := []
		rowCount := arr.MaxIndex(1)
		colCount := arr.MaxIndex(2)
		arrLabels := []
		
		Loop, % colCount
			arrLabels.push(arr[1, A_Index])
						
		Loop, % rowCount - 1
		{
			row := A_Index + 1
			
			ahkArr := {}
			Loop, % colCount
				ahkArr.Insert(arrLabels[A_Index], arr[row, A_Index])	
	
			ret.push(ahkArr)
		}
			
		return ret
	
	}

	
	GetFullPath(file)
	{
		Loop, % file
		return A_LoopFileLongPath
	}
	
	
}


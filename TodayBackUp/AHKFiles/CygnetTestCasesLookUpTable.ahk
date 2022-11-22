#Include %A_ScriptDir%\IncludesFiles.ahk

class CygnetTestCasesLookUptable {

__New(ExcelSheet)
    {
		this.ExcelSheet := ExcelSheet
		this.ExcelHandler := new ExcelHandler(ExcelSheet)
		this.ExcelArray := this.ExcelHandler.ExcelArray
		this.OutputFile := new ExcelHandler(this.OutputFile())
    }
	
	GetTestName(index){
		return this.ExcelArray[index].TestName
	}
	
	GetResult(index){
		return this.ExcelArray[index].Result
	}
	
	GetParent(index){
		return this.ExcelArray[index].Parent
	}
	
	GetChildren(index){
		return this.ExcelArray[index].Children
	}
	
	GetTime(index){
		return this.ExcelArray[index].Time
	}
	
	GetRow(index){
		return this.ExcelArray[index]
	}
	
	GetResultByTestName(testName){
		Loop, % this.ExcelArray.MaxIndex() {		
			if(this.ExcelArray[A_Index].TestName = testName){
				return this.ExcelArray[A_Index].Result
			}
		}
		return false
	}

    GetTimeByTestName(testName){
		Loop, % this.ExcelArray.MaxIndex() {		
			if(this.ExcelArray[A_Index].TestName = testName){
				return this.ExcelArray[A_Index].Time
			}
		}
		return false
	}
	
	SetResultByTestName(testName, result){
		Loop, % this.ExcelArray.MaxIndex() {
			if(this.ExcelArray[A_Index].TestName = testName){
				this.OutputFile.WriteToExcelSheetWithLabels(result, A_Index, 2, 1)
				return true
			}
		}
		return false
	}
	
	SetTimeByTestName(testName){
		Loop, % this.ExcelArray.MaxIndex() {		
			if(this.ExcelArray[A_Index].TestName = testName){
				currentTime := GetCurrentTime()
				this.OutputFile.WriteToExcelSheetWithLabels(currentTime, A_Index, 3, 1)
				return true
			}
		}
		return false
	}
	
	SetResultByTestIndex(testIndex, result){
		this.OutputFile.WriteToExcelSheetWithLabels(result, A_Index, 2, 1)

	}
	
	SetTimeByTestIndex(testIndex){
		this.OutputFile.WriteToExcelSheetWithLabels(currentTime, testIndex, 3, 1)
	}
	
	OutputFile(){
		fileName := this.ExcelSheet
		SplitPath, fileName,, dir
		dir.="\"
		SplitPath, fileName, name
		FormatTime, TimeString, L0x0009, ddddMMMMyyyy-HHmmss
		StringReplace , time, time, %A_Space%,,All
        FileCopy, %fileName%, %dir%%TimeString%%name%
		fileName = %dir%%TimeString%%name%
	   return fileName
	}	
}

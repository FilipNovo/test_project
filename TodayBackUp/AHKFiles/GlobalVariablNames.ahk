objShell := ComObjCreate("WScript.Shell")
objEnv := objShell.Environment("Volatile")

CYGNETHD_WINDOW_NAME := "Cygnet  - HD"
CYGNETEEG_WINDOW_NAME := "Cygnet  - 1_channel_eeg"
CYGNETINFRA_WINDOW_NAME := "Cygnet  - 1_channel_EEG.ILF"
CYGNETPIR_WINDOW_NAME := "Cygnet  - pIRx3"

PROGRAM_NAME := "Cygnet"
EXECTUION_FILE := "C:\bee.systems\Cygnet\Cygnet_2.0.7\bioera.exe"

CygnetTestCasesDefaultFile := "C:\Users\filip\Desktop\TodayBackUp\TestCases.ods"
CygnetElementsLookupTableDefaultFile := "C:\Users\filip\Desktop\TodayBackUp\CygnetElementsNameID.ods"

PROGRAM_LOCATION := "C:\Program Files (x86)\Cygnet"
FEEDBACK_WINDOW_NAME := "Select Feedback"
CYGNET_CLASS_NAME := "SunAwtFrame"

objEnv.Item("CYGNETHD_WINDOW_NAME") := CYGNETHD_WINDOW_NAME
objEnv.Item("CYGNETEEG_WINDOW_NAME") := CYGNETEEG_WINDOW_NAME
objEnv.Item("CYGNETINFRA_WINDOW_NAME") := CYGNETINFRA_WINDOW_NAME
objEnv.Item("CYGNETPIR_WINDOW_NAME") := CYGNETPIR_WINDOW_NAME

objEnv.Item("PROGRAM_NAME") := PROGRAM_NAME
objEnv.Item("PROGRAM_LOCATION") := PROGRAM_LOCATION

objEnv.Item("CygnetTestCasesDefaultFile") := CygnetTestCasesDefaultFile
objEnv.Item("CygnetElementsLookupTableDefaultFile") := CygnetElementsLookupTableDefaultFile

objEnv.Item("EXECTUION_FILE") := EXECTUION_FILE
objEnv.Item("FEEDBACK_WINDOW_NAME") := FEEDBACK_WINDOW_NAME
objEnv.Item("CYGNET_CLASS_NAME") := CYGNET_CLASS_NAME
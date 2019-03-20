'Import an Excel Sheet(a single file) to DataTable(UFT)
'A DataTable is similar with Excel. There are two types of Datatables: 
'	- Action/Local DataTable - Each action has its own private datatable, also known as local datatable, which can be accessed across actions.
'	- Global DataTable - Each test has one global data sheet that is accessible across actions.

'The syntax to import an Excel sheet is:  DataTable.ImportSheet Path_FileName, SheetName, DataTable_Sheet


'Import "Sheet1" from Excel TestImportExcelSheet.xlsx  to local Datatable "Action1"
'!!! The Excel file must be closed
'!!! Introduce the RIGHT path
DataTable.ImportSheet "C:\Users\user-pc\Desktop\TestImportExcelSheet.xlsx", "Sheet1", "Action1"

'The top row of Excel sheet "Sheet1" becomes the column header in datatable "Action1" (See Picture.png)

'i - Number of rows
i = DataTable.GetSheet("Action1").GetRowCount
'print "Number of rows: "  & i

'j - Number of columns
j = DataTable.GetSheet("Action1").GetParameterCount
'print "Number of colums: "  & j

'Iterate to all rows and print actions+parameters
For LineIterator = 1 To i
	'DataTable.SetCurrentRow(RowNumber) - set current row, specified by RowNumber
	DataTable.SetCurrentRow(LineIterator)
	'Select element from Datatable_Sheet :  DataTable(Column Name, Datatable_SheetName)
	Select Case DataTable("Action", "Action1")
		Case "UpdateSearch"
			UpdateSearch
		Case "LaunchDiagBox"
			LaunchDiagBox
		Case "LoginUFT"
			Set Param_Login = new User
			Param_Login = LoginUFT
		Case "Authentication"
			UserAction = DataTable("Parameter1", "Action1")
			UserName = DataTable("Parameter2", "Action1")
			UserPassword = DataTable("Parameter3", "Action1")
			BrandName = DataTable("Parameter4", "Action1")
			
			If UserAction = "CANCEL" Then
				Authentification UserAction, "", "", ""
			Else
				If UserName = "" and UserPassword = "" Then
					If Param_Login.name = "" and Param_Login.password = "" Then
						Desktop.CaptureBitmap path_images + "Failure_Authentication.png", True
						Reporter.ReportEvent micFail, "Failure Authentification(" + UserAction + ", " + UserName + ", " + UserPassword + ", " + BrandName + ")", "LoginUFT wasn't called before", path_images + "Failure_Authentication.png"
						f.WriteLine "FAILED Authentification(" + UserAction + ", " + UserName + ", " + UserPassword + ", " + BrandName + ")  " + "LoginUFT wasn't called before" + ", " + path_images + "Failure_Authentication.png"
						ArchiveFolder path + "\" + "APP_log.zip", "C:\APP\ddc\log"
						'ArchiveFolder path + "\" + "AWRoot_trace.zip", "C:\AWRoot\dtwr\trace"
						'ArchiveFolder path + "\" + "AWRoot_log.zip", "C:\AWRoot\dtwr\stcapi\log"
						ExitTest
					Else
						UserName = Param_Login.name
						UserPassword = Param_Login.password
					End If 
				End If
				Authentification UserAction, UserName, UserPassword, BrandName
			End If
			
		Case "SelectBrand"
			BrandName = DataTable("Parameter1", "Action1")

			SelectBrand BrandName
		Case "ModelSelect"
			DetectionType = DataTable("Parameter1", "Action1")
			
			ModelSelect DetectionType
		Case "WiFiButton"
			Name = DataTable("Parameter1", "Action1")
			
			WiFiButton Name
		Case "LaunchApplication"
			Name = DataTable("Parameter1", "Action1")
			
			LaunchApplication Name 
		Case "SelectButton"
			Action = DataTable("Parameter1", "Action1")
			
			SelectButton Action
		Case "SelectTab"
			TabName = DataTable("Parameter1", "Action1")
			
			SelectTab TabName
		Case "SelectECU"
			Family = DataTable("Parameter1", "Action1")
			SubFamily = DataTable("Parameter2", "Action1")
			
			SelectECU Family, SubFamily
		Case "SelectMenu"
			Name = DataTable("Parameter1", "Action1")
			
			SelectMenu Name
		Case "SelectSideMenu"
			Name = DataTable("Parameter1", "Action1")
			
			SelectSideMenu Name
		Case "CaptureScreen"
			Report = DataTable("Parameter1", "Action1")
			
			CaptureScreen Report
		Case "SeeTG"
			SeeTG
		Case "Impression"
			Impression
		Case "TestIdent"
			ParamName = DataTable("Parameter1", "Action1")
			Format = DataTable("Parameter2", "Action1")				
			DataType = DataTable("Parameter3", "Action1")	
			
			TestIdent ParamName, Format, DataType 
		Case "TestDTC"
			TestDTC
		Case "EFFDTC"
			EFFDTC
		Case "TestMP"
			ParamName = DataTable("Parameter1", "Action1")
			DataType = DataTable("Parameter2", "Action1")
			Format = DataTable("Parameter3", "Action1")
			Unit = DataTable("Parameter3", "Action1")
			Help = DataTable("Parameter3", "Action1")
										
			TestMP ParamName, DataType, Format, Unit, Help
		Case "Delay"
			DelayTime = DataTable("Parameter1", "Action1")
			If vartime = "" and DelayTime = "" Then
				Desktop.CaptureBitmap path_images + "FailureDelay" + CStr(DelayTime) + ".png", True
				Reporter.ReportEvent micFail, "Failure Delay(" + CStr(DelayTime) + ")", "No Delay(DelayTime) was called before" , path_images + "FailureDelay" + CStr(DelayTime) + ".png"
				f.WriteLine "FAILED Delay(" + CStr(DelayTime) + ")  " + "No Delay(DelayTime) was called before" + ", " +  path_images + "FailureDelay" + CStr(DelayTime) + ".png"	
				ArchiveFolder path + "\" + "APP_log.zip", "C:\APP\ddc\log"
				'ArchiveFolder path + "\" + "AWRoot_trace.zip", "C:\AWRoot\dtwr\trace"
				'ArchiveFolder path + "\" + "AWRoot_log.zip", "C:\AWRoot\dtwr\stcapi\log"
				ExitTest
			ElseIf DelayTime = "" Then
				Delay(vartime)
			Else
				vartime = Delay(DelayTime)
			End If
	End Select
Next

ArchiveFolder path + "\" + "APP_log.zip", "C:\APP\ddc\log"
'ArchiveFolder path + "\" + "AWRoot_trace.zip", "C:\AWRoot\dtwr\trace"
'ArchiveFolder path + "\" + "AWRoot_log.zip", "C:\AWRoot\dtwr\stcapi\log" @@ script infofile_;_ZIP::ssf77.xml_;_

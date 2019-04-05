'DataTable.ImportSheet "C:\UFT\TestImportExcelSheet.xlsx", "Sheet1", "Action1"
'
'i = DataTable.GetSheet("Action1").GetRowCount
'
'For LineIterator = 1 To i
'	DataTable.SetCurrentRow(LineIterator)
'	Select Case DataTable("Action", "Action1")
'		Case "UpdateSearch"
'			UpdateSearch
'		Case "LaunchDiagBox"
'			LaunchDiagBox
'		Case "LoginUFT"
'			LoginUFT
'		Case "Authentification"
'			UserAction = DataTable("Parameter1", "Action1")
'			UserName = DataTable("Parameter2", "Action1")
'			UserPassword = DataTable("Parameter3", "Action1")
'			BrandName = DataTable("Parameter4", "Action1")
'			
'			If UserAction = "CANCEL" Then
'				Authentification UserAction, "", "", ""
'			Else
'				If UserName = "" and UserPassword = "" Then
'					If LoginParamUft.name = "" and LoginParamUft.password = "" Then
'						counter_pic = counter_pic + 1
'						path_photo = TakeAScreenshot ("Authentification", counter_pic)
'						Reporter.ReportEvent micFail, "Failure Authentification(" + UserAction + ", " + UserName + ", *PASSWORD*, " + BrandName + ")", "LoginUFT wasn't called before or password/name is empty", path_photo
'						f.WriteLine "Authentification,FAILED," + path_photo + ",LoginUFT wasn't called before or password/name is empty"
'						ArchiveFolder path + "\" + "APP_log.zip", "C:\APP\ddc\log"
'						ArchiveFolder path + "\" + "AWRoot_trace.zip", "C:\AWRoot\dtwr\trace"
'						ArchiveFolder path + "\" + "AWRoot_log.zip", "C:\AWRoot\dtwr\stcapi\log"
'						ExitTest
'					Else
'						UserName = LoginParamUft.name
'						UserPassword = LoginParamUft.password
'						Authentification UserAction, UserName, UserPassword, BrandName
'					End If 
'				Else
'					LoginParamUft.name = UserName
'					LoginParamUft.password = UserPassword
'					Authentification UserAction, UserName, UserPassword, BrandName
'				End If
'			End If
'			
'		Case "SelectBrand"
'			BrandName = DataTable("Parameter1", "Action1")
'
'			SelectBrand BrandName
'		Case "ModelSelect"
'			DetectionType = DataTable("Parameter1", "Action1")
'			
'			ModelSelect DetectionType
'		Case "WiFiButton"
'			Name = DataTable("Parameter1", "Action1")
'			
'			WiFiButton Name
'		Case "LaunchApplication"
'			Name = DataTable("Parameter1", "Action1")
'			
'			LaunchApplication Name 
'		Case "SelectButton"
'			Action = DataTable("Parameter1", "Action1")
'			
'			SelectButton Action
'		Case "SelectTab"
'			TabName = DataTable("Parameter1", "Action1")
'			
'			SelectTab TabName
'		Case "SelectECU"
'			Family = DataTable("Parameter1", "Action1")
'			SubFamily = DataTable("Parameter2", "Action1")
'			
'			SelectECU Family, SubFamily
'		Case "SelectMenu"
'			Name = DataTable("Parameter1", "Action1")
'			
'			SelectMenu Name
'		Case "SelectSideMenu"
'			Name = DataTable("Parameter1", "Action1")
'			
'			SelectSideMenu Name
'		Case "TakeAScreenshot"
'			Report = DataTable("Parameter1", "Action1")
'			
'			TakeAScreenshot Report
'		Case "SeeTG"
'			SeeTG
'		Case "Impression"
'			Impression
'		Case "TestIdent"
'			ParamName = DataTable("Parameter1", "Action1")
'			Format = DataTable("Parameter2", "Action1")				
'			DataType = DataTable("Parameter3", "Action1")	
'			
'			TestIdent ParamName, Format, DataType 
'		Case "TestDTC"
'			TestDTC
'		Case "EFFDTC"
'			EFFDTC
'		Case "TestMP"
'			ParamName = DataTable("Parameter1", "Action1")
'			DataType = DataTable("Parameter2", "Action1")
'			Format = DataTable("Parameter3", "Action1")
'			Unit = DataTable("Parameter3", "Action1")
'			Help = DataTable("Parameter3", "Action1")
'										
'			TestMP ParamName, DataType, Format, Unit, Help
'		Case "Delay"
'			DelayTime = DataTable("Parameter1", "Action1")
'			Delay(DelayTime)
'	End Select
'Next
'
'ArchiveFolder path + "\" + "APP_log.zip", "C:\APP\ddc\log"
''ArchiveFolder path + "\" + "AWRoot_trace.zip", "C:\AWRoot\dtwr\trace"
''ArchiveFolder path + "\" + "AWRoot_log.zip", "C:\AWRoot\dtwr\stcapi\log" @@ script infofile_;_ZIP::ssf77.xml_;_
'

'LaunchDiagBox
'SelectBrand "PEUGEOT"
'wait 0, 4000
'ModelSelect "AUTO"
'Delay 10000
'Authentification "OK", "E518720", "Cst67677", "PEUGEOT"
'SelectButton "OK"
'SelectButton "OK"
'LaunchApplication "Fault finding"
'SelectTab "Expert"
'SelectECU "Built-in systems interface", "BSI2010"
'SelectButton "OK"
'SelectButton "OK"
'SelectMenu "FAULT READING"
'Delay 5000
'SelectSideMenu "Identification"

'Varin = TimeValue
'print Varin
'GlobalDelay = 5000
'SelectBrand "PEUGEOT"

'print Varin
'
'UserPassword = "ASFSAFS"
'BrandName = "FAFASFAs"


'SelectBrand "PEUGEOT"

'ModelSelect "406gsgd"

'LoginUft
'LaunchDiagBox
'LoginUFT
'time1 = timer
'print cstr(time1)
'path_photo = TakeAScreenshot ("LoginUFT", counter_pic)
'time1 = timer
'print cstr(time1)

'LaunchDiagBox
'SelectBrand "PEUGEOT"
'ModelSelect "AUTO"
'
'LaunchApplication "Fault finding"
'SelectECU "Engine management ECU", "CMM1_MD1CS003"

'SelectMenu "Identification"

'SelectSideMenu "Identification"


'TestIdent "Name of the Supplier","Date","CONTINENTAL"
TestMP "Status of the power train","","String","","Information broadcast on the Low Speed CAN networks"


 @@ script infofile_;_ZIP::ssf88.xml_;_

'Call for Authentication function
'UserAction = "OK"
'UserName = "username"
'UserPassword = "password"
'BrandName = "PEUGEOT"	
'
'If UserAction = "CANCEL" Then
'	Authentification UserAction, "", "", ""
'Else
'	If UserName = "" and UserPassword = "" Then
'		If LoginParamUft.name = "" and LoginParamUft.password = "" Then
'			path_photo = TakeAScreenshot ("Authentification", counter_pic)
'			Reporter.ReportEvent micFail, "Failure Authentification(" + UserAction + ", " + UserName + ", *PASSWORD*, " + BrandName + ")", "LoginUFT wasn't called before or password/name is empty", path_photo
'			f.WriteLine "Authentification,FAILED," + path_photo + ",LoginUFT wasn't called before or password/name is empty"
'			ArchiveFolder path + "\" + "APP_log.zip", "C:\APP\ddc\log"
'			'ArchiveFolder path + "\" + "AWRoot_trace.zip", "C:\AWRoot\dtwr\trace"
'			'ArchiveFolder path + "\" + "AWRoot_log.zip", "C:\AWRoot\dtwr\stcapi\log"
'			ExitTest
'		Else
'			UserName = LoginParamUft.name
'			UserPassword = LoginParamUft.password
'			Authentification UserAction, UserName, UserPassword, BrandName
'		End If 
'	Else
'		LoginParamUft.name = UserName
'		LoginParamUft.password = UserPassword
'		Authentification UserAction, UserName, UserPassword, BrandName
'	End If
'End If


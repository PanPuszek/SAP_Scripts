If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

Dim objExcel
Dim objSheet, i, shl
Set fso = CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject( "WScript.Shell" )
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
fso.CopyFile "L:\akartasi\SAP_Scripts\PRS_PlatesCosting.xlsm", "L:\akartasi\SAP_Scripts\PRS_PlatesCosting" & "_" & strUserName & ".xlsm"
Set shl = CreateObject("WScript.Shell")
shl.Run """L:\akartasi\SAP_Scripts\PRS_PlatesCosting" & "_" & strUserName & ".xlsm"""
INPUT = InputBox("Podaj numer BOM", "BOM")
Set objExcel = GetObject( ,"Excel.Application") 
WorkbookName = "PRS_PlatesCosting" & "_" & strUserName & ".xlsm"
Set objSheet = objExcel.Workbooks(WorkbookName).Worksheets("Main")




If INPUT <> "" Then
	
	session.findById("wnd[0]").resizeWorkingPane 82,17,false
	session.findById("wnd[0]/tbar[0]/okcd").text = "CS11"
	session.findById("wnd[0]/tbar[0]/btn[0]").press
	session.findById("wnd[0]/usr/ctxtRC29L-MATNR").text = INPUT
	session.findById("wnd[0]/usr/ctxtRC29L-CAPID").text = "PP01"
	session.findById("wnd[0]/usr/ctxtRC29L-CAPID").setFocus
	session.findById("wnd[0]/usr/ctxtRC29L-CAPID").caretPosition = 4
	session.findById("wnd[0]/tbar[1]/btn[5]").press
	session.findById("wnd[0]/usr/chkRC29L-VALST").setFocus
	session.findById("wnd[0]/usr/chkRC29L-VALST").selected = false
	session.findById("wnd[0]/tbar[1]/btn[8]").press
	
	If fso.FileExists("C:\Users\" & strUserName & "\Desktop\tag.vbs") Then
		
		session.findById("wnd[0]/tbar[1]/btn[32]").press
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").currentCellRow = 4
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").selectedRows = "4"
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_FL_SING").press
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "0"
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 2
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "2"
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 49
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 49
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "49"
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 71
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 71
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "71"
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 117
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 117
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "117"
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 119
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 119
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "119"
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").currentCellRow = 4
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").selectedRows = "4"
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").pressToolbarButton "DTC_UPPOS1"
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").currentCellRow = 6
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").selectedRows = "6"
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").pressToolbarButton "DTC_UP"
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").pressToolbarButton "DTC_UP"
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").pressToolbarButton "DTC_UP"
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").currentCellRow = 5
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").selectedRows = "5"
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").pressToolbarButton "DTC_UP"
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").currentCellRow = 6
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").selectedRows = "6"
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").pressToolbarButton "DTC_UP"
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").currentCellRow = 7
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").selectedRows = "7"
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").pressToolbarButton "DTC_UP"
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").currentCellRow = 9
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").selectedRows = "9"
		session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").pressToolbarButton "DTC_UP"
		session.findById("wnd[1]/tbar[0]/btn[5]").press
		session.findById("wnd[2]/usr/tabsG50_TABSTRIP/tabpTAB_D0501/ssubD0505_SUBSCREEN:SAPLSLVC_DIALOG:0501/txtLTDX-VARIANT").text = "AAA1"
		session.findById("wnd[2]/usr/tabsG50_TABSTRIP/tabpTAB_D0501/ssubD0505_SUBSCREEN:SAPLSLVC_DIALOG:0501/txtLTDXT-TEXT").text = "PRS_PlatesCosting"
		session.findById("wnd[2]/usr/tabsG50_TABSTRIP/tabpTAB_D0501/ssubD0505_SUBSCREEN:SAPLSLVC_DIALOG:0501/txtLTDXT-TEXT").setFocus
		session.findById("wnd[2]/usr/tabsG50_TABSTRIP/tabpTAB_D0501/ssubD0505_SUBSCREEN:SAPLSLVC_DIALOG:0501/txtLTDXT-TEXT").caretPosition = 11
		session.findById("wnd[2]/tbar[0]/btn[0]").press
		session.findById("wnd[1]/tbar[0]/btn[0]").press
		
		fso.DeleteFile("C:\Users\" & strUserName & "\Desktop\tag.vbs")
	End If
	
	session.findById("wnd[0]/tbar[1]/btn[33]").press
	session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = -1
	session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectColumn "VARIANT"
	session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").pressColumnHeader "VARIANT"
	session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
	session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
	session.findById("wnd[0]/tbar[1]/btn[43]").press
	session.findById("wnd[1]/usr/radRB_OTHERS").setFocus
	session.findById("wnd[1]/usr/radRB_OTHERS").select
	session.findById("wnd[1]/usr/cmbG_LISTBOX").setFocus
	session.findById("wnd[1]/usr/cmbG_LISTBOX").key = "10"
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & "\Desktop\"
	session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "temp.xlsx"
	session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	session.findById("wnd[0]/tbar[1]/btn[2]").press
	
	BOM_REVISION_NBR = session.findById("wnd[0]/usr/subSTLKOPF:SAPLCSDI:0802/txtRC29K-REVLV").text
	objSheet.Cells(1,16382).Value = BOM_REVISION_NBR
	BOM_DESCRIPTION = session.findById("wnd[0]/usr/subSTLKOPF:SAPLCSDI:0802/txtRC29K-OBKTX").text
	objSheet.Cells(1,16384) = BOM_DESCRIPTION
	objSheet.Cells(1,16383) = INPUT
	
	session.findById("wnd[0]/tbar[0]/btn[15]").press
	
	session.findById("wnd[0]/tbar[0]/btn[15]").press
	session.findById("wnd[0]/tbar[0]/btn[15]").press
	
	shl.AppActivate WorkbookName
	objExcel.Run WorkbookName & "!CopyFileData"
	
	
		
		FileName = CStr(objSheet.Cells(1, 2).Value) & "_" & CStr(objSheet.Cells(1, 5).Value)
		FileName = Replace(FileName, "Numer BOM: ", "")
		FileName = Replace(FileName, "Opis: ", "")
		FileName = Replace(FileName, ",", "_")
		FileName = Filename + "_PLATES"
		objExcel.Run "SaveAsPlates"
		InfoBox = MsgBox ("Arkusz zosta³ zapisany na pulpicie pod nazw¹: " + FileName + ".xlsx", vbOKOnly, "Zapis ukoñczony")
	
	
End if

If fso.FileExists("L:\akartasi\SAP_Scripts\" & WorkbookName) Then
	
	fso.DeleteFile("L:\akartasi\SAP_Scripts\" & WorkbookName)
	
End If
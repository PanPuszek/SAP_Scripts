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
'
Dim objExcel
Dim objSheet, objSheetBWS, i, shl
Set fso = CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject( "WScript.Shell" )
Set shl = CreateObject("WScript.Shell")
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
fso.CopyFile "L:\akartasi\SAP_Scripts\DRUMS.xlsm", "L:\akartasi\SAP_Scripts\DRUMS" & "_" & strUserName & ".xlsm"
shl.Run """L:\akartasi\SAP_Scripts\DRUMS" & "_" & strUserName & ".xlsm"""
INPUT = InputBox("Podaj numer BOM", "BOM")
Set objExcel = GetObject( ,"Excel.Application") 
WorkbookName = "DRUMS" & "_" & strUserName & ".xlsm"
Set objSheet = objExcel.Workbooks(WorkbookName).Worksheets("Main")
Set objSheetBWS = objExcel.Workbooks(WorkbookName).Worksheets("BWS")



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
		session.findById("wnd[2]/usr/tabsG50_TABSTRIP/tabpTAB_D0501/ssubD0505_SUBSCREEN:SAPLSLVC_DIALOG:0501/txtLTDXT-TEXT").text = "BOM_Explode"
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
	
	On Error Resume next
	session.findById("wnd[1]/usr/lblDY_PATH").setFocus
	
	If err.number <> 0 Then
	
		session.findById("wnd[1]/usr/radRB_OTHERS").setFocus
		session.findById("wnd[1]/usr/radRB_OTHERS").select
		session.findById("wnd[1]/usr/cmbG_LISTBOX").setFocus
		session.findById("wnd[1]/usr/cmbG_LISTBOX").key = "10"
		session.findById("wnd[1]/tbar[0]/btn[0]").press
		
	End If
	
	On Error GoTo 0
	
	session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & "\Desktop\"
	session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "temp.xlsx"
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
	'msgbox ("OK")
	objExcel.Run WorkbookName & "!CopyFileData"
	
	'########################
	'##########MM03##########
	'########################
	
	askSearch =  msgbox("Czy chcesz aby program wyszukal BWS automatycznie?", vbYesNo, "Opcje wyszukiwania")
	askPlates = vbNo
	
	If askSearch = vbYes Then
		
		askPlates = msgbox("Czy chcesz aby wyszukiwane byly rowniez BWS dla blach?" & vbCrLf & "(Wyszukiwanie BWS blach znaczaco wydluza dzialanie programu)", vbYesNo, "Wyszukiwanie blach")
	
	End If
	
	session.findById("wnd[0]/tbar[0]/okcd").text = "MM03"
	session.findById("wnd[0]/tbar[0]/btn[0]").press
	
	k = 2
	j = 3
	
	If askSearch = vbYes Then
		
		Do Until objSheet.Cells(j, 4).Value = "" 
			
			On Error Resume Next
			
			If askPlates = vbYes OR (Not InStr(objSheet.Cells(j, 9), "PLATE,DETAIL") <> 0 AND Not InStr(objSheet.Cells(j, 9), "PLATE,STEEL") <> 0) Then
				
				session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = objSheet.Cells(j, 4)
				session.findById("wnd[0]/tbar[1]/btn[5]").press
				
				session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).selected = true
				session.findById("wnd[1]/tbar[0]/btn[0]").press
				
				session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP27").select
				
				If Err.Number = 0 Then
					
					session.findById("wnd[1]").sendVKey 4
					session.findById("wnd[2]/usr/cntlGRID1/shellcont/shell").currentCellColumn = ""
					session.findById("wnd[2]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
					session.findById("wnd[2]/tbar[0]/btn[2]").press
					session.findById("wnd[1]/tbar[0]/btn[0]").press
					
					l = 1
					w = 4
					
					first = True
					
					Do Until DoStop
						
						objSheetBWS.Cells (3, k).Value = objSheet.Cells(j, 4).Value
						objSheetBWS.Cells (2, k).Value = objSheet.Cells(j, 9).Value
						
						session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP27").select
						
						objSheetBWS.Cells(w, k).Value = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP27/ssubTABFRA1:SAPLMGMM:2000/subSUB1:SAPLMGD1:1009/txtRMMG1_BEZ-WERKS_BEZ").text
						objSheetBWS.Cells(w + 1, k).Value = Cstr(session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP27/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2902/txtMBEW-STPRS").text) + " " + Cstr(session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP27/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2902/txtT001-WAERS").text)
						
						session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14").select
						
						If session.findById("wnd[2]/usr/txtMESSTXT1").text <> "MRP 3 not active for the organizational level" Then
							
							If first Then
								
								session.findById("wnd[1]/tbar[0]/btn[0]").press
								
							End If
							
							objSheetBWS.Cells(w + 2, k).Value = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2493/txtMARC-WZEIT").text
							
							session.findById("wnd[0]/tbar[1]/btn[13]").press
							session.findById("wnd[1]").sendVKey 4
							session.findById("wnd[2]/usr/cntlGRID1/shellcont/shell").setCurrentCell l,""
							session.findById("wnd[2]/usr/cntlGRID1/shellcont/shell").selectedRows = Cstr(l)
							session.findById("wnd[2]/tbar[0]/btn[2]").press
							
							If session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = "999" Then
								
								DoStop = True
								session.findById("wnd[1]/tbar[0]/btn[12]").press
								
							Else
								
								session.findById("wnd[1]/tbar[0]/btn[0]").press
								
						
									
							End If
							
							first = False
							
							l = l + 1
							
						End If
							
							w = w + 3
							
					Loop
					
					DoStop = false
					
					k = k + 2
					
				End If
				
				session.findById("wnd[0]/tbar[0]/btn[3]").press
	
				End If
				
			j = j + 1
		
		Loop
		
		ElseIf AskSearch = vbNo Then
			
			objExcel.Run WorkbookName & "!ChooseParts"
			
			Do Until askChoice = vbYes
				
				shl.AppActivate "R3P(1)/050 SAP Easy Access"
				shl.AppActivate "R3P(1)/050 Display Material (Initial Screen)"
				shl.AppActivate "Display Material (Initial Screen)"
				
				msgbox "W odpowiednich wierszach wpisz 'x' aby wyszukac dana czesc", vbApplicationModal
				askChoice = msgbox ("Czy jestes pewien?" & vbCrLf & vbCrLf & "Wybranie opcji 'Tak' uniemozliwi dalsza modyfikacje i zatrzymanie programu nie bedzie mozliwe az do jego kompletnego wykonania.", vbYesNo, "Ostrzezenie")
			
			Loop
			
			j = 3
			k = 2
			
			Do Until objSheet.Cells(j, 4).Value = "" 
				
				On Error Resume Next
				
			If objSheet.Cells(j, 12).Value = "x" OR objSheet.Cells(j, 12).Value = "X" Then
				
				session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = objSheet.Cells(j, 4)
				session.findById("wnd[0]/tbar[1]/btn[5]").press
				
				session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).selected = true
				session.findById("wnd[1]/tbar[0]/btn[0]").press
				
				session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP27").select
				
				If Err.Number = 0 Then
					
					session.findById("wnd[1]").sendVKey 4
					session.findById("wnd[2]/usr/cntlGRID1/shellcont/shell").currentCellColumn = ""
					session.findById("wnd[2]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
					session.findById("wnd[2]/tbar[0]/btn[2]").press
					session.findById("wnd[1]/tbar[0]/btn[0]").press
					
					l = 1
					w = 4
					
					first = True
					
					Do Until DoStop
					
						objSheetBWS.Cells (3, k).Value = objSheet.Cells(j, 4).Value
						objSheetBWS.Cells (2, k).Value = objSheet.Cells(j, 9).Value
						
						session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP27").select
						
						objSheetBWS.Cells(w, k).Value = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP27/ssubTABFRA1:SAPLMGMM:2000/subSUB1:SAPLMGD1:1009/txtRMMG1_BEZ-WERKS_BEZ").text
						objSheetBWS.Cells(w + 1, k).Value = Cstr(session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP27/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2902/txtMBEW-STPRS").text) + " " + Cstr(session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP27/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2902/txtT001-WAERS").text)
						
						session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14").select
						
						If session.findById("wnd[2]/usr/txtMESSTXT1").text <> "MRP 3 not active for the organizational level" Then
							
							If first Then
								
								session.findById("wnd[1]/tbar[0]/btn[0]").press
								
							End If
							
							objSheetBWS.Cells(w + 2, k).Value = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2493/txtMARC-WZEIT").text
							
							session.findById("wnd[0]/tbar[1]/btn[13]").press
							session.findById("wnd[1]").sendVKey 4
							session.findById("wnd[2]/usr/cntlGRID1/shellcont/shell").setCurrentCell l,""
							session.findById("wnd[2]/usr/cntlGRID1/shellcont/shell").selectedRows = Cstr(l)
							session.findById("wnd[2]/tbar[0]/btn[2]").press
							
							If session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = "999" Then
								
								DoStop = True
								session.findById("wnd[1]/tbar[0]/btn[12]").press
								
							Else
								
								session.findById("wnd[1]/tbar[0]/btn[0]").press
								
							
								
							End If
							
							first = False
							
							l = l + 1
						
						End If						
						
						w = w + 3
						
					Loop
					
					DoStop = false
					
					k = k + 2
					
				End If
				
				session.findById("wnd[0]/tbar[0]/btn[3]").press
	
				End If
				
			j = j + 1
		
		Loop
	
	End If
	
	
	session.findById("wnd[0]/tbar[0]/btn[15]").press
	
	shl.AppActivate WorkbookName
	objExcel.Run WorkbookName & "!SortingBWS"

End if

If fso.FileExists("L:\akartasi\SAP_Scripts\" & WorkbookName) Then
	
	fso.DeleteFile("L:\akartasi\SAP_Scripts\" & WorkbookName)
	
End If
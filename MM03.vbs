'Kod dodany przez SAP
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


REM ***//POCZ¥TEK KODU DODANEGO PRZEZ EXCEL//***

'Stworzenie zmiennych o nazwie objExcel, objSheet
'Nastêpnie nastêpuje nadanie im w³aœciwoœci obiektu i przypisanie objSheet arkusza Excela o nazwie
'MM03.xlsm oraz zak³adki o nazwie Main
Dim objExcel
Dim objSheet, i, shl, wshShell, fso
Set fso = CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject("WScript.Shell")
Set shl = CreateObject("WScript.Shell")
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
fso.CopyFile "L:\akartasi\SAP_Scripts\MM03.xlsm", "L:\akartasi\SAP_Scripts\MM03" & "_" & strUserName & ".xlsm"
shl.Run """L:\akartasi\SAP_Scripts\MM03" & "_" & strUserName & ".xlsm"""
INPUT = InputBox("Podaj numer czêœci", "Wyszukiwanie czêœci")
Set objExcel = GetObject(,"Excel.Application") 
WorkbookName = "MM03" & "_" & strUserName & ".xlsm"
Set objSheet = objExcel.Workbooks(WorkbookName).Worksheets("Main")



If INPUT <> "" Then
	
	session.findById("wnd[0]/tbar[0]/okcd").text = "MM03"
	session.findById("wnd[0]/tbar[0]/btn[0]").press
	
	session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = INPUT
	session.findById("wnd[0]").sendVKey 0
	
	session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).selected = true
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	
	session.findById("wnd[0]/tbar[1]/btn[30]").press
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU04").select
	
	i = 1
	Do while  session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU04/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:3400/subDOCU:SAPLCV140:0204/subDOC_ALV:SAPLCV140:0207/tblSAPLCV140SUB_DOC/ctxtDRAW-DOKAR[0,0]").text <> ""
		
		ObjSheet.Cells(i+1, 2).Value = session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU04/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:3400/subDOCU:SAPLCV140:0204/subDOC_ALV:SAPLCV140:0207/tblSAPLCV140SUB_DOC/ctxtDRAW-DOKAR[0,0]").text
		ObjSheet.Cells(i+1, 3).Value = session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU04/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:3400/subDOCU:SAPLCV140:0204/subDOC_ALV:SAPLCV140:0207/tblSAPLCV140SUB_DOC/ctxtDRAW-DOKNR[1,0]").text
		ObjSheet.Cells(i+1, 4).Value = session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU04/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:3400/subDOCU:SAPLCV140:0204/subDOC_ALV:SAPLCV140:0207/tblSAPLCV140SUB_DOC/ctxtDRAW-DOKTL[2,0]").text
		ObjSheet.Cells(i+1, 5).Value = session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU04/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:3400/subDOCU:SAPLCV140:0204/subDOC_ALV:SAPLCV140:0207/tblSAPLCV140SUB_DOC/ctxtDRAW-DOKVR[3,0]").text
		ObjSheet.Cells(i+1, 6).Value = session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU04/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:3400/subDOCU:SAPLCV140:0204/subDOC_ALV:SAPLCV140:0207/tblSAPLCV140SUB_DOC/txtDRAT-DKTXT[4,0]").text
		
		session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU04/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:3400/subDOCU:SAPLCV140:0204/subDOC_ALV:SAPLCV140:0207/tblSAPLCV140SUB_DOC").verticalScrollbar.position = i
		
		i = i + 1
		
	Loop
	
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU05").select
	
	If session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU05/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:2031/tblSAPLMGD1TC_LONGTEXT/txtLANG_TC_TAB_TC-SPTXT[1,0]").text <> "" Then
		
		ObjSheet.Cells(4, 8).Value = session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU05/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:2031/cntlLONGTEXT_GRUNDD/shellcont/shell").text
		
	End If
	
	session.findById("wnd[0]/tbar[1]/btn[27]").press
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP15").select
	session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = "999"
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	ObjSheet.Cells(4, 10).Value = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP15/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2504/ctxtMARC-NFMAT").text
	
	
	session.findById("wnd[0]/tbar[0]/btn[15]").press
	
	objExcel.Run WorkbookName & "!Labels"
	
End if

If fso.FileExists("L:\akartasi\SAP_Scripts\" & WorkbookName) Then
	
	fso.DeleteFile("L:\akartasi\SAP_Scripts\" & WorkbookName)
	
End If
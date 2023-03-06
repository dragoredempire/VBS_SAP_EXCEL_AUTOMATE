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
Dim objSheet, intRow, i
Set objExcel = GetObject(,"Excel.Application")
Set objSheet = objExcel.ActiveWorkbook.Worksheets("LISTAS")
Set objSheetfast = objExcel.ActiveWorkbook.Worksheets("Fast Entry 2")
Set objSheetLESER = objExcel.ActiveWorkbook.Worksheets("C_PL_LESER_Ltda")
Set objSheetComercial = objExcel.ActiveWorkbook.Worksheets("Comercial")
max = Trim(CStr(objSheetfast.Cells(1, 65).Value))*1' máximo
'lista de tags adicionais
Set myRange = objSheetfast.Range("BO2:CM"&max)
Set myList = CreateObject("System.Collections.ArrayList")
For Each cell In myRange
	if cell.value <> "" then
		myList.Add cell.Value
	end if
Next
contador_erros = 0
'fim da lista



proposta = objSheetLESER.range("S9").Text

INCONTERM = objSheetComercial.range("P1").text



objExcel.Calculation  = EnableCalculation
objSheet.Cells(1, 64) = "ERF2_" ' QUAL ERF
objSheet.calculate


'session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nva22"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = proposta

session.findById("wnd[0]").sendVKey 0

'######## copiar informações do cliente
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press 'entra no cabeçalho
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07").select 'entra na aba partners
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,1]").setFocus
nome = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/txtGVS_TC_DATA-REC-NAME1[2,1]").text
session.findById("wnd[0]").sendVKey 2 'entra no cliente
empresa = session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME1").text
estado = session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-REGION").text 'ESTADO
email = session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtSZA1_D0100-SMTP_ADDR").text 'email
telefone = session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtSZA1_D0100-TEL_NUMBER").text 'telefone
cidade = session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-CITY1").text
session.findById("wnd[1]").close 'fechar janela 1
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12").select 'aba additional datab
ticket = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZTICKET").text
session.findById("wnd[0]/tbar[0]/btn[3]").press 'voltar
po_number = session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").text
'############# final
objSheetLESER.range("C7") = empresa
objSheetLESER.range("C8") = nome
objSheetLESER.range("J9") = email
objSheetLESER.range("D9") = telefone
objSheetLESER.range("C11") = "[ Ticket: " & ticket & " ]" & po_number
objSheetComercial.range("N1") = estado
ESTADO =  objSheetComercial.range("N1").text

IF INCONTERM = "FCA" then
	ESTADO = "RIO DE JANEIRO - RJ"
end if

IF INCONTERM = "CIF" then
	ESTADO = cidade &" - "& ESTADO
end if

session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").select
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/ctxtVBKD-INCO1").text = INCONTERM
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/txtVBKD-INCO2").text = ESTADO
session.findById("wnd[0]").sendVKey 0
on error resume next
session.findById("wnd[1]/tbar[0]/btn[0]").press
on error goto 0

space1=Trim(CStr(objSheet.Cells(2, 66).Value)) ' ERF
space2=Trim(CStr(objSheet.Cells(3, 66).Value)) ' ERF
space3=Trim(CStr(objSheet.Cells(4, 66).Value)) ' ERF
space4=Trim(CStr(objSheet.Cells(5, 66).Value)) ' ERF
space5=Trim(CStr(objSheet.Cells(6, 66).Value)) ' ERF ' option code
space6=Trim(CStr(objSheet.Cells(7, 66).Value)) ' ERF
space7=Trim(CStr(objSheet.Cells(8, 66).Value)) ' ERF
space8=Trim(CStr(objSheet.Cells(9, 66).Value)) ' ERF
space9=Trim(CStr(objSheet.Cells(10, 66).Value)) ' ERF
space10=Trim(CStr(objSheet.Cells(11, 66).Value)) ' ERF




'session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\07").select
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4434/cmbRV45A-MUEBS").key = " "
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4434/btnFCODE_MAUS").press
session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,0]").text = space1
session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,1]").text = space2
session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,2]").text = space3
session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,3]").text = space4
session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,4]").text = space5
session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,5]").text = space6
session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,6]").text = space7
session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,7]").text = space8
session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,8]").text = space9
session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,9]").text = space10
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08").select

For x=0 To 42 ' quantas linhas cabem no sap
	On Error Resume next
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/txtVBAP-POSNR[0," & x &"]").text = " "

	If err.number = 0 then
	maxsap=maxsap+1
	err.clear
	End if

Next
On error GoTo 0
max_scrollbar = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG").verticalScrollbar.Maximum 
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG").verticalScrollbar.Position = max_scrollbar
y=0


For i=2 To max

	COL2    = Trim(CStr(objSheetfast.Cells(i,  1).Value)) * 1 'item sap
	COL3    = Trim(CStr(objSheetfast.Cells(i,  3).Value))  ' MATERIAL
	COL4    = Trim(CStr(objSheetfast.Cells(i,  4).Value))  ' QUANTIDADE
	COL5    = Trim(CStr(objSheetfast.Cells(i,  6).Value))  'INPUT 3
	COL6    = Trim(CStr(objSheetfast.Cells(i,  7).Value))  'INPUT 4
	COL7    = Trim(CStr(objSheetfast.Cells(i,  8).Value))  'INPUT 5
	COL8    = Trim(CStr(objSheetfast.Cells(i,  9).Value))  'INPUT 6
	COL9    = Left(Trim(CStr(objSheetfast.Cells(i,  10).Value)),30)  'INPUT 7 ' option code 'até 30 caracteres
	MORE 	= MID(Trim(CStr(objSheetfast.Cells(i,  10).Value)),31)
	
	COL10   = Trim(CStr(objSheetfast.Cells(i, 11).Value))  'INPUT 8
	COL11   = Trim(CStr(objSheetfast.Cells(i, 12).Value))  'INPUT 9
	COL12   = Trim(CStr(objSheetfast.Cells(i, 13).Value))  'INPUT 10
	COL13   = Trim(CStr(objSheetfast.Cells(i, 14).Value))  'INPUT 11
	COL14   = Trim(CStr(objSheetfast.Cells(i, 15).Value))  'INPUT 12
	COL15   = Trim(CStr(objSheetfast.Cells(i, 18).Value))  'ITEMCATEGORY
	max_scrollbar = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG").verticalScrollbar.Maximum 
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG").verticalScrollbar.Position = max_scrollbar
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/txtVBAP-POSNR[0," & y &"]").text = COL2
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MABNR[2," & y &"]").text = COL3
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/txtRV45A-KWMENG[3," & y &"]").text = COL4
	z=5
	If space1  <> "" Then
	z=z+1
	end if
	If COL5 <> "" then
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT01["& z &"," & y &"]").text = COL5
	end if
	If space2  <> "" Then
	z=z+1
	end if
	if COL6 <> "" then
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT02["& z &"," & y &"]").text = COL6
	end if
	If space3  <> "" Then
	z=z+1
	end if
	if COL7 <> "" then
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT03["& z &"," & y &"]").text = COL7
	end if
	If space4  <> "" Then
	z=z+1
	end if
	if COL8 <> "" then
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT04["& z &"," & y &"]").text = COL8
	end if
	If space5  <> "" Then
	z=z+1
	end if
	if COL9 <> "" then
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT05["& z &"," & y &"]").text = COL9
	end if
	If space6  <> "" Then
	z=z+1
	end if
	if COL10 <> "" then
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT06["& z &"," & y &"]").text = COL10
	end if
	If space7  <> "" Then
	z=z+1
	end if
	if COL11 <> "" then
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT07["& z &"," & y &"]").text = COL11
	end if
	If space8  <> "" Then
	z=z+1
	end if
	if COL12 <> "" then
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT08["& z &"," & y &"]").text = COL12
	end if
	If space9  <> "" Then
	z=z+1
	end if
	if COL13 <> "" then
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT09["& z &"," & y &"]").text = COL13
	end if
	If space10  <> "" Then
	z=z+1
	end if
	if COL14 <> "" then
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT10["& z &"," & y &"]").text = COL14
	end if

	On Error Resume next
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtVBAP-PSTYV[19," & y &"]").text = COL15

	If err.number <> 0 then
		session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtVBAP-PSTYV[19, " & y &"]").text = COL15
		err.Clear
		On error goTo 0
	End If

	If y => maxsap Then
		session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG").verticalScrollbar.position = i-1
		
		On Error Resume next 
		objSheet.Cells(i, 14) = " "
		session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT01[6," & y-1 &"]").text = " "

		Do while err.number <> 0 
			session.findById("wnd[1]/usr/tblSAPLCEI0VALUE_S/txtRCTMS-ATWRT[1,1]").text = myList(contador_erros)
			if session.findById("wnd[1]/usr/tblSAPLCEI0VALUE_S/txtRCTMS-ATWRT[1,2]").text = "M84" then
				session.findById("wnd[1]/usr/tblSAPLCEI0VALUE_S/radRCTMS-SEL01[0,2]").selected = true
			end if
			session.findById("wnd[1]/tbar[0]/btn[8]").press
			'session.findById(""wnd[1]/usr/tblSAPLCEI0VALUE_S/txtRCTMS-ATWRT[1,2]"").caretPosition = 1
			'session.findById("wnd[1]").sendVKey 2
			session.findById("wnd[1]/usr/btnBUTTON_1").press
			session.findById("wnd[2]/usr/btnBUTTON_1").press
			err.Clear
			session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT01[6," & y-1 &"]").text = " "
			contador_erros = contador_erros + 1
		loop

		On Error GoTo 0

		'session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG").verticalScrollbar.position = i-3
		y=-1
		i=i-1
	End if

	y=y+1

	If i = max Then

		session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG").verticalScrollbar.position = i-1

		On Error Resume next 
		session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT01[6," & y-1 &"]").text = COL5

		Do while err.number <> 0
			session.findById("wnd[1]/usr/tblSAPLCEI0VALUE_S/txtRCTMS-ATWRT[1,1]").text = myList(contador_erros)
			if session.findById("wnd[1]/usr/tblSAPLCEI0VALUE_S/txtRCTMS-ATWRT[1,2]").text = "M84" then
				session.findById("wnd[1]/usr/tblSAPLCEI0VALUE_S/radRCTMS-SEL01[0,2]").selected = true
			end if
			session.findById("wnd[1]/tbar[0]/btn[8]").press
			'session.findById(""wnd[1]/usr/tblSAPLCEI0VALUE_S/txtRCTMS-ATWRT[1,2]"").caretPosition = 1
			'session.findById("wnd[1]").sendVKey 2
			session.findById("wnd[1]/usr/btnBUTTON_1").press
			session.findById("wnd[2]/usr/btnBUTTON_1").press
			err.Clear
			session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT01[6," & y-1 &"]").text = COL5
			contador_erros = contador_erros + 1
		loop

		On Error GoTo 0

	End If 

Next

objSheet.Cells(1, 64) = "ERF12_" ' QUAL ERF
objSheet.calculate

space1=Trim(CStr(objSheet.Cells(2, 66).Value)) ' ERF
space2=Trim(CStr(objSheet.Cells(3, 66).Value)) ' ERF
space3=Trim(CStr(objSheet.Cells(4, 66).Value)) ' ERF
space4=Trim(CStr(objSheet.Cells(5, 66).Value)) ' ERF
space5=Trim(CStr(objSheet.Cells(6, 66).Value)) ' ERF
space6=Trim(CStr(objSheet.Cells(7, 66).Value)) ' ERF
space7=Trim(CStr(objSheet.Cells(8, 66).Value)) ' ERF
space8=Trim(CStr(objSheet.Cells(9, 66).Value)) ' ERF
space9=Trim(CStr(objSheet.Cells(10, 66).Value)) ' ERF
space10=Trim(CStr(objSheet.Cells(11, 66).Value)) ' ERF

session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\07").select

session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4434/btnFCODE_MAUS").press
session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,0]").text = space1
session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,1]").text = space2
session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,2]").text = space3
session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,3]").text = space4
session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,4]").text = space5
session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,5]").text = space6
session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,6]").text = space7
session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,7]").text = space8
session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,8]").text = space9
session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,9]").text = space10
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]").sendVKey 0


session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4434/subSUBSCREEN_TC:SAPMV45A:4923/tblSAPMV45ATCTRL_U_ERF_KONFIG").verticalScrollbar.position = 0
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08").select

y=0
For i=2 To max 'ERF12
COL5    = Trim(CStr(objSheetfast.Cells(i,  27).Value))  'DESENVOLVIDA
COL6    = Trim(CStr(objSheetfast.Cells(i,  20).Value))  'CERTIFICADOS
COL7    = Trim(CStr(objSheetfast.Cells(i,  21).Value))  'CERTIFICADOS PETROBRAS
COL8    = Left(Trim(CSTr(objSheetfast.Cells(i,  22).value)),30)  'testes PETROBRAS até 30 caracteres
COL9 	= MID(Trim(CStr(objSheetfast.Cells(i,  22).Value)),31) & MORE ' o restante
COL10 =""
COL11   = Trim(CStr(objSheetfast.Cells(i, 66).Value))  'TAG1
COL12   = Trim(CStr(objSheetfast.Cells(i, 67).Value))  'TAG2
COL13   = Trim(CStr(objSheetfast.Cells(i, 68).Value))  'TAG3
COL14   = Trim(CStr(objSheetfast.Cells(i, 69).Value))  'TAG4
z=5
If space1  <> "" Then
z=z+1
end if
'MSGBOX Y & "-" & maxsap
If COL5 <> "" then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT01["& z &"," & y &"]").text = COL5
end if
If space2  <> "" Then
z=z+1
end if
if COL6 <> "" then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT02["& z &"," & y &"]").text = COL6
end if
If space3  <> "" Then
z=z+1
end if
if COL7 <> "" then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT03["& z &"," & y &"]").text = COL7
end if
If space4  <> "" Then
z=z+1
end if
if COL8 <> "" then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT04["& z &"," & y &"]").text = COL8
end if
If space5  <> "" Then
z=z+1
end if
if COL9 <> "" then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT05["& z &"," & y &"]").text = COL9
end if
If space6  <> "" Then
z=z+1
end if
if COL10 <> "" then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT06["& z &"," & y &"]").text = COL10
end if
If space7  <> "" Then
z=z+1
end if
if COL11 <> "" then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT07["& z &"," & y &"]").text = COL11 'tag1
end if
If space8  <> "" Then
z=z+1
end if
if COL12 <> "" then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT08["& z &"," & y &"]").text = COL12 'tag2
end if
If space9  <> "" Then
z=z+1
end if
if COL13 <> "" then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT09["& z &"," & y &"]").text = COL13 'tag3
end if
If space10  <> "" Then
z=z+1
end if
if COL14 <> "" then
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT10["& z &"," & y &"]").text = COL14 'tag4
end if

If y => maxsap-1 Then

	session.findById("wnd[0]").sendVKey 82 'pagedown

	On Error Resume next 

	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT01[6," & y-1 &"]").text = " "


	Do while err.number <> 0 
		session.findById("wnd[1]/usr/tblSAPLCEI0VALUE_S/txtRCTMS-ATWRT[1,1]").text = "YOU NEED TO FILL"
		if session.findById("wnd[1]/usr/tblSAPLCEI0VALUE_S/txtRCTMS-ATWRT[1,2]").text = "M84" then
			session.findById("wnd[1]/usr/tblSAPLCEI0VALUE_S/radRCTMS-SEL01[0,2]").selected = true
		end if
		session.findById("wnd[1]/tbar[0]/btn[8]").press
		'session.findById(""wnd[1]/usr/tblSAPLCEI0VALUE_S/txtRCTMS-ATWRT[1,2]"").caretPosition = 1
		session.findById("wnd[1]").sendVKey 2
		session.findById("wnd[1]/usr/btnBUTTON_1").press
		session.findById("wnd[2]/usr/btnBUTTON_1").press
		err.Clear
		session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT01[6," & y-1 &"]").text = " "
	loop

	On Error GoTo 0

	'session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG").verticalScrollbar.position = i-1
	y=-1
	

End if

y=y+1

If i = max Then

	session.findById("wnd[0]").sendVKey 0
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG").verticalScrollbar.position = i-1

	On Error Resume next 
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT01[6," & y-1 &"]").text = COL5

	Do while err.number <> 0
		session.findById("wnd[1]/usr/tblSAPLCEI0VALUE_S/txtRCTMS-ATWRT[1,1]").text = "YOU NEED TO FILL"
		if session.findById("wnd[1]/usr/tblSAPLCEI0VALUE_S/txtRCTMS-ATWRT[1,2]").text = "M84" then
			session.findById("wnd[1]/usr/tblSAPLCEI0VALUE_S/radRCTMS-SEL01[0,2]").selected = true
		end if
		session.findById("wnd[1]/tbar[0]/btn[8]").press
		session.findById("wnd[1]/usr/btnBUTTON_1").press
		session.findById("wnd[2]/usr/btnBUTTON_1").press
		err.Clear
		session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT01[6," & y-1 &"]").text = COL5
	loop

	On Error GoTo 0

End If 

Next
  '######################## contrapressão

if objExcel.Application.WorksheetFunction.XLOOKUP("variable", objSheetfast.range("W2:W" & max),objSheetfast.range("W2:W" & max),"aa") ="variable" or objExcel.Application.WorksheetFunction.XLOOKUP("constant", objSheetfast.range("W2:W" & max),objSheetfast.range("W2:W" & max),"aa") ="constant"   then

	objSheet.Cells(1, 64) = "ERF_BACKPRESSURE" ' QUAL ERF
	objSheet.calculate

	space1=""
	space2=""
	space3=""
	space4=""
	space5=""
	space6=Trim(CStr(objSheet.Cells(7, 66).Value)) ' ERF
	space7=Trim(CStr(objSheet.Cells(8, 66).Value)) ' ERF
	space8=Trim(CStr(objSheet.Cells(9, 66).Value)) ' ERF
	space9=Trim(CStr(objSheet.Cells(10, 66).Value)) ' ERF
	space10=Trim(CStr(objSheet.Cells(11, 66).Value)) ' ERF

	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\07").select

	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4434/btnFCODE_MAUS").press
	session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,0]").text = space1
	session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,1]").text = space2
	session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,2]").text = space3
	session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,3]").text = space4
	session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,4]").text = space5
	session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,5]").text = space6
	session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,6]").text = space7
	session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,7]").text = space8
	session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,8]").text = space9
	session.findById("wnd[1]/usr/tblSAPMV45ATCTRL_CHAR_VA_SEL/ctxtCHAR_VA_SEL-ATNAM[0,9]").text = space10
	session.findById("wnd[1]/tbar[0]/btn[8]").press
	session.findById("wnd[0]").sendVKey 0
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4434/subSUBSCREEN_TC:SAPMV45A:4923/tblSAPMV45ATCTRL_U_ERF_KONFIG").verticalScrollbar.position = 0
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08").select

	y=0
	For i=2 To max

	COL5=""
	COL6=""
	COL7=""
	COL8=""
	COL9=""
	COL10 	= Trim(CStr(objSheetfast.Cells(i, 23).Value)) 'tipo de backp
	COL11   = Trim(CStr(objSheetfast.Cells(i, 24).Value))  'min
	COL12   = Trim(CStr(objSheetfast.Cells(i, 25).Value))  'max
	COL13   = Trim(CStr(objSheetfast.Cells(i, 26).Value))  'const
	COL14   = Trim(CStr(objSheetfast.Cells(i, 27).Value))  'desenvolvida
	z=5
	If space1  <> "" Then
	z=z+1
	end if
	z=6
	if COL10 <> "" then
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT06["& z &"," & y &"]").text = COL10 ' tipo de contrapressão
	end if
	If space7  <> "" Then
	z=7
	end if
	if COL13 <> "" and COL13 <> 0 then
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT07["& z &"," & y &"]").text = COL13 'constant
	end if
	If space8  <> "" Then
	z=z+1
	end if
	if COL11 <> "" and COL12 <> "" and COL12 <> 0 then
	'msgbox z & " " & y
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT08["& z &"," & y &"]").text = COL11 'min
	end if
	If space9  <> "" Then
	z=z+1
	end if
	if COL12 <> "" and COL12 <> 0 then
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT09["& z &"," & y &"]").text = COL12 'maximo
	end if
	If space10  <> "" Then
	z=z+1
	end if
	if COL14 <> "" and COL14 <> 0 then
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT10["& z &"," & y &"]").text = COL14 'desenvolvida
	end if

	If y => maxsap-1 Then

		session.findById("wnd[0]").sendVKey 82

		On Error Resume next 

		

		On Error GoTo 0

		'session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG").verticalScrollbar.position = i-3
		y=-1
		
	End if

	y=y+1

	If i = max Then

		session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG").verticalScrollbar.position = i-3

	End If 

	Next
		
		
		
end if


y=2
max = Trim(CStr(objSheetfast.Cells(1, 65).Value))*1+1 ' máximo
 

do
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").select
	err.clear


	for x = 0 to 23

	 	on error resume next


		item = objSheetfast.range("A"& y)
		semanas = objSheetfast.range("BM"& y)
		err.clear
		y=y+1
		teste= session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0," & x &"]").text*1
		if err.number <> 0 Then
			session.findById("wnd[0]").sendVKey 82
			Err.Clear
			x=0
		end if
		
		
		if session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0," & x &"]").text*1 = item*1 Then
				session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/cmbVBKD-DELCO[37," & x &"]").key = semanas
		Else
			
			session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POPO").press
			if err.number <> 0 Then
				err.clear
				session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POPO").press
				ERR.CLEAR
			end if
			session.findById("wnd[1]/usr/txtRV45A-POSNR").text = item
			session.findById("wnd[1]").sendVKey 0
			x=-1
			y=y-1
			err.clear
		
		end if
		

		if y = max Then
			msgbox "END"
			Exit do
		end if

	next
loop




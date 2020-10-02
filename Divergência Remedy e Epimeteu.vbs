'-----------------------------------------------------------
'|  SCRIPT PARA PREENCHER EXCEL - TRELLO EXPORT    			|
'|															|
'|  DESENVOLVEDOR: JOSE ARTEIRO TEIXEIRA (JOSE.CAVALCANTI)	|
'------------------------------------------------------------

Sub Criar_colunas(Col_Name, Col_Orig, Col_Number)
	Set objRange = objExcel.Range(Col_Orig).EntireColumn
	objRange.Insert(xlShiftToRight)
	objExcel.Cells(1, Col_Number).Value = Col_Name
End Sub

Function Ajust_Cols(Range_cell_border, Num_Cols)
	objExcel.Range(Range_cell_border).Select
	objExcel.Cells.EntireColumn.AutoFit
	objExcel.Range(Range_cell_border).Borders.colorindex = 1
	
	For Aux_loop = 1 to Num_Cols
		objExcel.Cells(1, Aux_loop).Interior.Color = RGB(141, 180, 226)
	Next
End Function

Function Ajust_Cols_Red(Range_cell_border)
	objExcel.Range(Range_cell_border).Select
	objExcel.Cells.EntireColumn.AutoFit
	objExcel.Range(Range_cell_border).Borders.colorindex = 1
	
	For Aux_loop = 16 to 19
		objExcel.Cells(1, Aux_loop).Interior.Color = RGB(248, 203, 173)
	Next
End Function

Function Contar_Linhas(Col_range, Sheet_Name)
	Cont_rows = 0
	Count_White = 0
	objExcel.Cells(1, 1).Select
	For Each Cell In objWorkbook.Worksheets(Sheet_Name).Range(Col_range).Cells
		'If Cell.Value = "" Then Exit For
		If Cell.Value = "" Then 
			Count_White = Count_White + 1
		Else
			Cont_rows = Cont_rows + 1
			Count_White = 0
		End If
		If Count_White = 10 Then Exit For
	Next
	Contar_Linhas = Cont_rows
End Function

Function Create_Sheet(Sheet_Name)
	objExcel.ActiveWorkbook.Sheets.Add 
	objExcel.ActiveSheet.name = Sheet_Name
End Function

Function Copy_Sheet(Orign_Sheet, Dest_Sheet)
	objExcel.Worksheets(Orign_Sheet).Copy , objExcel.Worksheets(Orign_Sheet)
	objExcel.ActiveSheet.name = Dest_Sheet
End Function

Function WhatEver(num)
    If(Len(num)=1) Then
        WhatEver="0"&num
    Else
        WhatEver=num
    End If
End Function

Function myDateFormat(myDate)
    d = WhatEver(Day(myDate)-1)
    m = WhatEver(Month(myDate))    
    y = Year(myDate)
	If d < 1 Then
		d = 30
		m = m - 1
	End If
	If m < 1 Then
		m = 12
		y = y - 1
	End If
    myDateFormat= d & "/" & m & "/" & y
End Function

Sub Delete_Column(P_Val, P_Column)
	If objExcel.Cells(1, P_Column).Value = P_Val Then
		objExcel.Range(P_Column & ":" & P_Column).Delete
	Else
		Msgbox "Não encontrado Coluna '" & P_Val &"' em " & P_Column
		objExcel.ActiveWorkbook.Close
		objExcel.Application.Quit
		WScript.Quit
	End If
End Sub

Sub Move_Column(P_Val, P_Column, T_Column)
	If objExcel.Cells(1, P_Column).Value = P_Val Then
		
		objExcel.Range(T_Column & ":" & T_Column).Value = objExcel.Range(P_Column & ":" & P_Column).Value
		objExcel.Range(P_Column & ":" & P_Column).Delete
	Else
		Msgbox "Não encontrado Coluna '" & P_Val &"' em " & P_Column
		objExcel.ActiveWorkbook.Close
		objExcel.Application.Quit
		WScript.Quit
	End If
End Sub

Function WorksheetExists(wsName, objWorkbook)
    
    ret = False
    wsName = UCase(wsName)
    For Each ws In objWorkbook.Sheets
        If UCase(ws.Name) = wsName Then
            ret = True
            Exit For
        End If
    Next
    WorksheetExists = ret
End Function

Sub Remover_Sheet(STR_VAR)
	'Stopping Application Alerts
	objExcel.DisplayAlerts=FALSE
	
	objExcel.Worksheets(STR_VAR).delete 
	'objWorkbook.sheets(STR_VAR).delete
	
	'Enabling Application alerts once we are done with our task
	objExcel.DisplayAlerts=TRUE
End Sub

Function Validar_TrelloExport(Aux)
	Dim Column_Name(30)
	Column_Name(1) = "SISTEMA"
	Column_Name(2) = "NUM_ERRO"
	Column_Name(3) = "ABERTURA_PKE"
	Column_Name(4) = "24x7 dias Fila Acc -> CRQ"
	Column_Name(5) = "NUM_CRQ"
	Column_Name(6) = "DAT_ASSOCIACAO"
	Column_Name(7) = "DAT_FIM_PLAN"
	Column_Name(8) = "DAT_ALVO"
	Column_Name(9) = "DAT_RELEASE_ATUAL"
	Column_Name(10) = "EQUIPE_REGRA"
	Column_Name(11) = "TEC_RESPONSAVEL"
	Column_Name(12) = "STATUS"
	Column_Name(13) = "DET_STATUS"
	Column_Name(14) = "AGENTE_SOLUCIONADOR"
	Column_Name(15) = "AGS_ABERTURA"
	Column_Name(16) = "OFENSOR"
	Column_Name(17) = "MUD_ASSOCIADA"
	Column_Name(18) = "DAT_INICIAL"
	Column_Name(19) = "DAT_FINAL"
	Column_Name(20) = "DAT_RELEASE_ORIGINAL"
	Column_Name(21) = "REABERTURA"
	Column_Name(22) = "PRIORIDADE"
	Column_Name(23) = "DAT_ENVIO_PDC"
	Column_Name(24) = "CANDIDATURA"
	Column_Name(25) = "TAKEOVER"
	Column_Name(26) = "MISSED"
	Column_Name(27) = "BUILD"
	Column_Name(28) = "DESC_ERRO"
	Column_Name(29) = "JUSTIFICATIVA_PERDA_SLA"
	Column_Name(30) = "HISTORICO"
	
	If Aux = "Colum_num" Then
		Validar_TrelloExport = uBound(Column_Name)
	
	Else		
		STR_CSV = Column_Name(1)
		For Column = 2 to uBound(Column_Name)
			STR_CSV = STR_CSV & ";" & Column_Name(Column)
		Next
		
		Validar_TrelloExport = STR_CSV
	End If
End Function


'------ VOID MAIN ------
VAR_TRELLO_EMANAGER = "-= Divergência Remedy e Epimeteu V1.0=-"

MsgBox(VAR_TRELLO_EMANAGER)


Dim DEBUG

DEBUG = MsgBox("Debug?", "36", VAR_TRELLO_EMANAGER)

'For Each Count_Month in Var_Count_Month
'	Count_Month = 0
'Next

Set FSO = WScript.CreateObject("Scripting.FileSystemObject")
	Set objExcel = CreateObject("Excel.Application")
		vFileName = objExcel.GetOpenFilename ("CSV Files (*.csv), *.csv")
		
		If (FSO.FileExists(vFileName)) Then
			Set OTF = FSO.OpenTextFile(vFileName, 1)
				
				TextLine = OTF.ReadLine
					
				If (Validar_TrelloExport("ColumnTitle") = TextLine) Then
					contents = OTF.ReadAll
					OTF.Close
					contents = Replace(contents, vbCr, "")
					contents = Replace(contents, vbLf, "")				
				End If
				
			Set OTF = Nothing
			
			If Len(contents) > 0 Then
				
				Array_Column = Split(contents, ";")
							
				Count_Array = 0
				Total_Array = Validar_TrelloExport("Colum_num") - 1
				
				Test = ""
				For ForAux = 0 to uBound(Array_Column)
					Count_Array = Count_Array + 1
					Test = Test & ";" & Array_Column(ForAux)
					If Count_Array = Total_Array Then
						WScript.Echo Test
						Count_Array = 0
						Test = ""
					End IF
				Next
			Else
				WScript.Echo "Colunas Incorretas."
			End If
			
		Else
			WScript.Echo "Arquivo " & vFileName & " Inexistente"
		End If
	Set objExcel = Nothing
Set FSO = Nothing



WScript.Quit



				'While (OTF.AtEndOfStream <> True) 
				'	TextLine = OTF.ReadLine
				'	WScript.Echo TextLine
				'Wend
				'OTF.Close

				'Set OTF = FSO.OpenTextFile("teste.txt", 8, True)
				'	OTF.Write contents
				'	OTF.Close
				'Set OTF = Nothing
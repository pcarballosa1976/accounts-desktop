'====================================================================================
'
'  DetailedBrief.bas - Formulario de resumen detallado sobre los gastos e ingresos
'  por período.
'  
'  El formulario permite ver de un modo detallado y para un período determinado (diario,
'  mensual, anual) los distintos gastos e ingresos introducidos, así como el saldo
'  obtenido y los totales para dicho período.
'
'  Copyright (C) 2025 Cuban Solutions
'
'  Este programa es software libre: puedes redistribuirlo y/o modificarlo
'  bajo los términos de la Licencia Pública General de GNU publicada por
'  la Free Software Foundation, ya sea la versión 3 de la Licencia, o
'  (a tu elección) cualquier versión posterior.
'
'  Este programa se distribuye con la esperanza de que sea útil,
'  pero SIN GARANTÍA; sin incluso la garantía implícita de COMERCIALIZACIÓN
'  o ADECUACIÓN PARA UN PROPÓSITO PARTICULAR. Consulta la Licencia Pública
'  General de GNU para más detalles.
'
'  Deberías haber recibido una copia de la Licencia Pública General de GNU
'  junto a este programa. Si no, consulta <https://www.gnu.org/licenses/>.
'
'  También puedes visitar el blog del autor en: https://cubansolutions.blogspot.com
'
'====================================================================================

DECLARE SUB DetailedBriefMainFormClose (Action AS INTEGER)
DECLARE SUB DetailedBriefParametersQueryButtonClick (Sender AS QBUTTON)
DECLARE SUB DetailedBriefDetailStringGridDrawCell (Col As Integer, Row As Integer, State As Integer, Rect As QRect, Sender As QStringGrid)
DECLARE SUB DetailedBriefDetailStringGridKeyDown (Key AS Word, Shift AS INTEGER)

CREATE DetailedBriefMainForm AS QFORM
    Caption = "Resumen detallado"
    Width = 610
    Height = 458
	BorderStyle = bsDialog
    Center
	OnClose = DetailedBriefMainFormClose
    CREATE DetailedBriefTopPanel AS QPANEL
        Left = 0
        Top = 0
        Height = 60
        Align = alTop
        TabOrder = 0
        CREATE DetailedBriefTopCommentTitleLabel AS QLABEL
            Caption = "Resumen detallado de gastos e ingresos"
            Left = 15
            Top = 13
            Width = 200
			Font.AddStyles(fsBold)
        END CREATE
        CREATE DetailedBriefTopCommentTextLabel AS QLABEL
            Caption = "La vista de los gastos e ingresos por concepto introducidos, y de sus totales y saldo según período."
            Left = 15
            Top = 30
            Width = 550
            Height = 62
            Wordwrap = 1
        END CREATE
    END CREATE
    CREATE DetailedBriefParametersPanel AS QPANEL
        Left = 0
        Top = 49
        Height = 44
        Align = 1
        TabOrder = 1
        CREATE DetailedBriefParametersDateLabel AS QLABEL
            Caption = "Fecha:"
            Left = 10
            Top = 13
            Transparent = 1
        END CREATE
        CREATE DetailedBriefParametersDayEdit AS QEDIT
            Text = ""
            Left = 46
            Top = 11
            Width = 20
			InputMask = "00!"
            MaxLength = 2
            TabOrder = 0
        END CREATE
        CREATE DetailedBriefParametersMonthEdit AS QEDIT
            Text = ""
            Left = 74
            Top = 11
            Width = 20
			InputMask = "00!"
            MaxLength = 2
            TabOrder = 1
        END CREATE
        CREATE DetailedBriefParametersYearEdit AS QEDIT
            Text = ""
            Left = 102
            Top = 11
            Width = 32
			InputMask = "0000!"
            MaxLength = 4
            TabOrder = 2
        END CREATE
		CREATE DetailedBriefParametersPeriodLabel AS QLABEL
            Caption = "Período:"
            Left = 151
            Top = 13
            Width = 43
            Transparent = 1
        END CREATE
        CREATE DetailedBriefParametersPeriodComboBox AS QCOMBOBOX
            AddItems "Diario", "Mensual", "Anual"
            Left = 196
            Top = 11
            Width = 97
			Style = csDropDownList 
            TabOrder = 3
        END CREATE
        CREATE DetailedBriefParametersQueryButton AS QBUTTON
            Caption = "&Consultar"
            Left = 522
            Top = 10
            TabOrder = 4
            Default = 1
			OnClick = DetailedBriefParametersQueryButtonClick
        END CREATE
    END CREATE
    CREATE DetailedBriefDetailPanel AS QPANEL
        Left = 0
        Top = 93
        Align = 5
        TabOrder = 2
        CREATE DetailedBriefDetailStringGrid AS QSTRINGGRID
            AddOptions(goRowSelect)
			Left = 6
            Top = 12
			DefaultRowHeight = 18
			DefaultColWidth = 80
			FixedRows = 1
			RowCount = 1
			FixedCols = 0
            ColCount = 7
			Height = 234
			Width = 592
			ColWidths(0) = 0
			ColWidths(1) = 65
			ColWidths(2) = 158
			ColWidths(3) = 120
			ColWidths(4) = 70
			ColWidths(5) = 70
			Cell(1,0) = " Fecha"
			Cell(2,0) = " Concepto"
			Cell(3,0) = " Detalle"
			Cell(4,0) = " Ingreso"
			Cell(5,0) = " Gasto"
			Cell(6,0) = " Saldo"
			OnDrawCell = DetailedBriefDetailStringGridDrawCell
			OnKeyDown = DetailedBriefDetailStringGridKeyDown
        END CREATE
    END CREATE
    CREATE DetailedBriefSubtotalsPanel AS QPANEL
        Left = 0
        Top = 350
        Height = 70
        Align = 2
        TabOrder = 3
		CREATE DetailedBriefHistoricTotalsLabel AS QLABEL
            Caption = "Histórico:"
            Left = 6
            Top = 17
			Transparent = 1
			Font.AddStyles(fsBold)
        END CREATE
		CREATE DetailedBriefHistoricTotalsBalanceRichEdit AS QRICHEDIT
			Alignment = 2
            Text = ""
            Left = 64
            Top = 15
			Height = 21
            Width = 100
            TabOrder = 0
			PlainText = 1
			ReadOnly = True
			Font.AddStyles(fsBold)
			Font.Color = QBColor(2)
        END CREATE
		CREATE DetailedBriefHistoricTotalsBalanceLabel AS QLABEL
            Caption = "Saldo"
            Left = 98
            Top = 40
			Transparent = 1
			Font.AddStyles(fsBold)
			Font.Color = QBColor(2)
        END CREATE
        CREATE DetailedBriefSubtotalsTotalsLabel AS QLABEL
            Caption = "Totales:"
            Left = 258
            Top = 17
			Transparent = 1
			Font.AddStyles(fsBold)
        END CREATE
		CREATE DetailedBriefSubtotalsTotalsRevenuesRichEdit AS QRICHEDIT
			Alignment = 2
            Text = ""
            Left = 308
            Top = 15
			Height = 21
            Width = 90
			TabOrder = 1
			PlainText = 1
			ReadOnly = True
			Font.AddStyles(fsBold)
			Font.Color = QBColor(9)
        END CREATE
		CREATE DetailedBriefSubtotalsTotalsExpensesRichEdit AS QRICHEDIT
			Alignment = 2
            Text = ""
            Left = 402
            Top = 15
			Height = 21
            Width = 90
            TabOrder = 2
			PlainText = 1
			ReadOnly = True
			Font.AddStyles(fsBold)
			Font.Color = QBColor(12)
        END CREATE
		CREATE DetailedBriefSubtotalsTotalsBalanceRichEdit AS QRICHEDIT
			Alignment = 2
            Text = ""
            Left = 496
            Top = 15
			Height = 21
            Width = 100
            TabOrder = 3
			PlainText = 1
			ReadOnly = True
			Font.AddStyles(fsBold)
			Font.Color = QBColor(2)
        END CREATE
		CREATE DetailedBriefSubtotalsTotalsRevenuesLabel AS QLABEL
            Caption = "Ingresos"
            Left = 328
            Top = 40
			Transparent = 1
			Font.AddStyles(fsBold)
			Font.Color = QBColor(9)
        END CREATE
        CREATE DetailedBriefSubtotalsTotalsExpensesLabel AS QLABEL
            Caption = "Gastos"
            Left = 430
            Top = 40
            Transparent = 1
			Font.AddStyles(fsBold)
			Font.Color = QBColor(12)
        END CREATE
        CREATE DetailedBriefSubtotalsTotalsBalanceLabel AS QLABEL
            Caption = "Saldo"
            Left = 531
            Top = 40
			Transparent = 1
			Font.AddStyles(fsBold)
			Font.Color = QBColor(2)
        END CREATE
    END CREATE
END CREATE

'--------- Subroutines ---------

SUB DetailedBriefMainFormClose (Action AS INTEGER)
    MainDBFTable.Close
END SUB

SUB DetailedBriefParametersQueryButtonClick (Sender AS QBUTTON)
	If Val(DetailedBriefParametersMonthEdit.Text) <= 0 Or Val(DetailedBriefParametersMonthEdit.Text) > 12 Then
		MessageDlg("El mes debe ser un número positivo entre 01 y 12, introduzca un valor correcto y vuelva a intentarlo.", mtInformation, mbOk, 0)
		Exit Sub
	End If
	
	If Val(DetailedBriefParametersYearEdit.Text) < 1980 Or Val(DetailedBriefParametersYearEdit.Text) > Val(Right$(Date$, 4)) Then
		MessageDlg("El año debe ser un número positivo entre 1980 y el año en curso, introduzca un valor correcto y vuelva a intentarlo.", mtInformation, mbOk, 0)
		Exit Sub
	End If
	
	Select Case Val(DetailedBriefParametersMonthEdit.Text)
		Case 4, 6, 9, 11                  
			Select Case Val(DetailedBriefParametersDayEdit.Text)
				Case 1 TO 30
				Case Else
					MessageDlg("El día debe ser un número positivo entre 01 y 30, introduzca un valor correcto y vuelva a intentarlo.", mtInformation, mbOk, 0)
					Exit Sub
			End Select
		Case 1, 3, 5, 7, 8, 10, 12
			Select Case Val(DetailedBriefParametersDayEdit.Text)
				Case 1 TO 31
				Case Else
					MessageDlg("El día debe ser un número positivo entre 01 y 31, introduzca un valor correcto y vuelva a intentarlo.", mtInformation, mbOk, 0)
					Exit Sub
			End Select
		Case 2                            
			Select Case Val(DetailedBriefParametersYearEdit.Text) MOD 4      
				Case 0  
					Select Case Val(DetailedBriefParametersDayEdit.Text)
						Case 1 TO 29
						Case Else
							MessageDlg("El día debe ser un número positivo entre 01 y 29, introduzca un valor correcto y vuelva a intentarlo.", mtInformation, mbOk, 0)
							Exit Sub		
					End Select
				Case Else
					Select Case Val(DetailedBriefParametersDayEdit.Text)
						Case 1 TO 28 
						Case Else
							MessageDlg("El día debe ser un número positivo entre 01 y 28, introduzca un valor correcto y vuelva a intentarlo.", mtInformation, mbOk, 0)
							Exit Sub
					End Select
			End Select                  
		Case Else
			MessageDlg("El mes debe ser un número positivo entre 01 y 12, introduzca un valor correcto y vuelva a intentarlo.", mtInformation, mbOk, 0)
			Exit Sub		
	End Select
	
	If DetailedBriefParametersPeriodComboBox.ItemIndex = -1 Then
		MessageDlg("Es necesario seleccionar un período en la lista desplegable.", mtInformation, mbOk, 0)
		Exit Sub
	End If
	
	Dim Day As String, Month As String, Year As String
	
	Day = IIF(Val(DetailedBriefParametersDayEdit.Text) < 10, "0" + Str$(Val(DetailedBriefParametersDayEdit.Text)), Str$(Val(DetailedBriefParametersDayEdit.Text)))
	Month = IIF(Val(DetailedBriefParametersMonthEdit.Text) < 10, "0" + Str$(Val(DetailedBriefParametersMonthEdit.Text)), Str$(Val(DetailedBriefParametersMonthEdit.Text)))
	Year = DetailedBriefParametersYearEdit.Text
	
	Dim StartBalance As Double
	Dim HistoricRevenues As Double, HistoricExpenses As Double, HistoricBalance As Double
	Dim PeriodRevenues As Double, PeriodExpenses As Double
	Dim Row As DWord
	Dim ConceptID As String, MovementDate As String
	Dim MovementData As MovmentsDataType
	Dim Movments As QMemoryStream
	
	StartBalance = 0
	HistoricRevenues = 0
	HistoricExpenses = 0
	HistoricBalance = 0
	PeriodRevenues = 0
	PeriodExpenses = 0
	Row = 0
	
	Dim StartDate As String, EndDate As String
	
	Select Case DetailedBriefParametersPeriodComboBox.ItemIndex
		Case 0 'Diario
			StartDate = Year + Month + Day
		Case 1 'Mensual
			StartDate = Year + Month + "01"
		Case 2 'Anual
			StartDate = Year + "01" + "01"
		Case Else
			MessageDlg("Es necesario seleccionar un período en la lista desplegable... no se continuará con la operación en curso.", mtInformation, mbOk, 0)
			Exit Sub
	End Select
	
	EndDate = Year + Month + Day
	
	If AuxDBFTable.Connect(CurDir$ + "\data\" + "BALANCES.DBF") = False Then
		MessageDlg("La tabla 'BALANCES.DBF' no es válida o no se encuentra en la carpeta 'data' del programa.", mtError, mbOk, 0)
		Exit Sub
	End If
	
	If AuxDBFTable.Locate("YEAR", Str$(Val(Year) - 1), False) Then
		StartBalance = Val(AuxDBFTable.FieldByName("AMOUNT"))
	End If
	
	AuxDBFTable.Close
	
	If AuxDBFTable.Connect(CurDir$ + "\data\" + "CONCEPTS.DBF") = False Then
		MessageDlg("La tabla 'CONCEPTS.DBF' no es válida o no se encuentra en la carpeta 'data' del programa.", mtError, mbOk, 0)
		Exit Sub
	End If

	With MainDBFTable
		.GoTop
		While .Fetch
			MovementDate = .FieldByName("DATE")
			If Val(Left$(MovementDate, 4)) = Val(Year) Then
				ConceptID = .FieldByName("CONCEPT")
				AuxDBFTable.GoTop
				If AuxDBFTable.Locate("ID", ConceptID, False) Then
					If MovementDate <= EndDate Then
						If MovementDate >= StartDate Then
							MovementData.Id = .FieldByName("ID")
							MovementData.Date = .FieldByName("DATE")
							MovementData.ConceptDesc = AuxDBFTable.FieldByName("DESC")
							MovementData.Detail = .FieldByName("DETAIL")
							MovementData.Kind = AuxDBFTable.FieldByName("KIND")
							MovementData.Amount = .FieldByName("AMOUNT")
							Select Case Val(MovementData.Kind)
								Case 1
									PeriodRevenues += Val(MovementData.Amount)
								Case 2
									PeriodExpenses += Val(MovementData.Amount)
							End Select
							Movments.WriteUDT(MovementData)
							Row += 1
						Else
							Select Case Val(AuxDBFTable.FieldByName("KIND"))
								Case 1
									HistoricRevenues += Val(.FieldByName("AMOUNT"))
								Case 2
									HistoricExpenses += Val(.FieldByName("AMOUNT"))
							End Select
						End If
					End If
				End If
			End If
			DoEvents
		Wend
	End With
	
	AuxDBFTable.Close
	
	HistoricBalance = StartBalance + HistoricRevenues - HistoricExpenses
	
	If HistoricBalance < 0 Then
		DetailedBriefHistoricTotalsBalanceRichEdit.Font.Color = QBColor(12)
	Else
		DetailedBriefHistoricTotalsBalanceRichEdit.Font.Color = QBColor(2)
	End If
	DetailedBriefHistoricTotalsBalanceRichEdit.Text = Format$("%.2f", HistoricBalance)
	
	Dim i As DWord
	
	DetailedBriefDetailStringGrid.RowCount = 1
	Movments.Position = 0
	For i = 1 To Row
		DetailedBriefDetailStringGrid.InsertRow(i)
		Movments.ReadUDT(MovementData)
		With DetailedBriefDetailStringGrid
			.Cell(0, i) = MovementData.Id
			.Cell(1, i) = Right$(MovementData.Date, 2) + "/" + Mid$(MovementData.Date, 5, 2) + "/" + Left$(MovementData.Date, 4)
			.Cell(2, i) = MovementData.ConceptDesc
			.Cell(3, i) = MovementData.Detail
			Select Case Val(MovementData.Kind)
				Case 1
					HistoricRevenues += Val(MovementData.Amount)
					HistoricBalance += Val(MovementData.Amount)
					.Cell(4, i) = Format$("%.2f", Val(MovementData.Amount))
				Case 2
					HistoricExpenses += Val(MovementData.Amount)
					HistoricBalance -= Val(MovementData.Amount)
					.Cell(5, i) = Format$("%.2f", Val(MovementData.Amount))
			End Select
			.Cell(6, i) = Format$("%.2f", HistoricBalance)
		End With
		DoEvents
	Next i
	Movments.Close
	
	If DetailedBriefDetailStringGrid.RowCount > 1 Then
		DetailedBriefDetailStringGrid.FixedRows = 1	
	End If
	
	DetailedBriefSubtotalsTotalsRevenuesRichEdit.Text = Format$("%.2f", PeriodRevenues)
	DetailedBriefSubtotalsTotalsExpensesRichEdit.Text = Format$("%.2f", PeriodExpenses)
	If HistoricBalance < 0 Then
		DetailedBriefSubtotalsTotalsBalanceRichEdit.Font.Color = QBColor(12)
	Else
		DetailedBriefSubtotalsTotalsBalanceRichEdit.Font.Color = QBColor(2)
	End If
	DetailedBriefSubtotalsTotalsBalanceRichEdit.Text = Format$("%.2f", HistoricBalance)

END SUB

SUB DetailedBriefDetailStringGridDrawCell (Col As Integer, Row As Integer, State As Integer, Rect As QRect, Sender As QStringGrid)
	If Row > 0 Then
		Sender.FillRect(Rect.Left, Rect.Top, Rect.Right, Rect.Bottom, QBColor(15))
		Select Case Col
			Case 4
				Sender.TextOut(Rect.Left + ((Rect.Right - Rect.Left) - Sender.TextWidth(Sender.Cell(Col, Row)) - 2), Rect.Top + 2, Sender.Cell(Col, Row), QBColor(9), -1)
			Case 5
				Sender.TextOut(Rect.Left + ((Rect.Right - Rect.Left) - Sender.TextWidth(Sender.Cell(Col, Row)) - 2), Rect.Top + 2, Sender.Cell(Col, Row), QBColor(12), -1)
			Case 6
				If Val(Sender.Cell(Col, Row)) < 0 Then
					Sender.TextOut(Rect.Left + ((Rect.Right - Rect.Left) - Sender.TextWidth(Sender.Cell(Col, Row)) - 2), Rect.Top + 2, Sender.Cell(Col, Row), QBColor(12), -1)
				Else
					Sender.TextOut(Rect.Left + ((Rect.Right - Rect.Left) - Sender.TextWidth(Sender.Cell(Col, Row)) - 2), Rect.Top + 2, Sender.Cell(Col, Row), QBColor(2), -1)
				End If
			Case Else
				Sender.TextOut(Rect.Left + 2, Rect.Top + 2, Sender.Cell(Col, Row), 0, -1)
		End Select
	Else
		'Sender.TextOut(Rect.Left + 2, Rect.Top + 2, Sender.Cell(Col, Row), 0, -1)
	End If
END SUB

SUB DetailedBriefDetailStringGridKeyDown(Key AS Word, Shift AS INTEGER)
	If DetailedBriefDetailStringGrid.Row <> 0 And Key = 46 Then
		Dim MovmentsID As String
		If MessageDlg("¿Está seguro de eliminar el movimiento seleccionado?" + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10) + "La vista deberá ser actualizada con el botón Consultar para mostrar correctamente los saldos.", mtConfirmation, mbYes Or mbNo, 0) = mrYes Then
			With MainDBFTable
				MovmentsID = DetailedBriefDetailStringGrid.cell(0, DetailedBriefDetailStringGrid.Row)
				.GoTop
				If .Locate("ID", MovmentsID, False) = True Then
					If .Delete = True Then
						DetailedBriefDetailStringGrid.DeleteRow(DetailedBriefDetailStringGrid.Row)
						MessageDlg("El movimiento se eliminó correctamente del sistema, recuerde utilizar el botón consultar para actualizar la vista y recalcular los saldos.", mtInformation, mbOk, 0)
					Else
						.LastError(MainDBFErr)
						MessageDlg(MainDBFErr.Description, mtError, mbOk, 0)
					End If
				End If
			End With
		End If   
	End If
END SUB

'====================================================================================
'
'  ConceptsBrief.bas - Formulario de resumen consolidado por conceptos de gastos
'  e ingresos.
'  
'  El formulario permite ver de un modo general y para un período determinado (diario,
'  mensual, anual) los totales de gastos e ingresos totalizados por conceptos.
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

DECLARE SUB ConceptsBriefMainFormClose (Action AS INTEGER)
DECLARE SUB ConceptsBriefParametersQueryButtonClick (Sender AS QBUTTON)
DECLARE SUB ConceptsBriefDetailStringGridDrawCell (Col As Integer, Row As Integer, State As Integer, Rect As QRect, Sender As QStringGrid)

CREATE ConceptsBriefMainForm AS QFORM
    Caption = "Resumen por conceptos"
    Width = 506
    Height = 458
	BorderStyle = bsDialog
    Center
	OnClose = ConceptsBriefMainFormClose
    CREATE ConceptsBriefTopPanel AS QPANEL
        Left = 0
        Top = 0
        Height = 60
        Align = alTop
        TabOrder = 0
        CREATE ConceptsBriefTopCommentTitleLabel AS QLABEL
            Caption = "Resumen por conceptos de gastos e ingresos"
            Left = 15
            Top = 13
            Width = 200
			Font.AddStyles(fsBold)
        END CREATE
        CREATE ConceptsBriefTopCommentTextLabel AS QLABEL
            Caption = "La vista de los conceptos introducidos, y de sus totales de gastos o ingresos según período."
            Left = 15
            Top = 30
            Width = 550
            Height = 62
            Wordwrap = 1
        END CREATE
    END CREATE
    CREATE ConceptsBriefParametersPanel AS QPANEL
        Left = 0
        Top = 49
        Height = 44
        Align = 1
        TabOrder = 1
        CREATE ConceptsBriefParametersDateLabel AS QLABEL
            Caption = "Fecha:"
            Left = 10
            Top = 13
            Transparent = 1
        END CREATE
        CREATE ConceptsBriefParametersDayEdit AS QEDIT
            Text = ""
            Left = 46
            Top = 11
            Width = 20
			InputMask = "00!"
            MaxLength = 2
            TabOrder = 0
        END CREATE
        CREATE ConceptsBriefParametersMonthEdit AS QEDIT
            Text = ""
            Left = 74
            Top = 11
            Width = 20
			InputMask = "00!"
            MaxLength = 2
            TabOrder = 1
        END CREATE
        CREATE ConceptsBriefParametersYearEdit AS QEDIT
            Text = ""
            Left = 102
            Top = 11
            Width = 32
			InputMask = "0000!"
            MaxLength = 4
            TabOrder = 2
        END CREATE
		CREATE ConceptsBriefParametersPeriodLabel AS QLABEL
            Caption = "Período:"
            Left = 151
            Top = 13
            Width = 43
            Transparent = 1
        END CREATE
        CREATE ConceptsBriefParametersPeriodComboBox AS QCOMBOBOX
            AddItems "Diario", "Mensual", "Anual"
            Left = 196
            Top = 11
            Width = 97
			Style = csDropDownList 
            TabOrder = 3
        END CREATE
        CREATE ConceptsBriefParametersQueryButton AS QBUTTON
            Caption = "&Consultar"
            Left = 422
            Top = 10
			Width = 70
            TabOrder = 4
            Default = 1
			OnClick = ConceptsBriefParametersQueryButtonClick
        END CREATE
    END CREATE
    CREATE ConceptsBriefDetailPanel AS QPANEL
        Left = 0
        Top = 93
        Align = 5
        TabOrder = 2
        CREATE ConceptsBriefDetailStringGrid AS QSTRINGGRID
            AddOptions(goRowSelect)
			Left = 6
            Top = 12
			DefaultRowHeight = 18
			DefaultColWidth = 100
			FixedRows = 1
			RowCount = 1
			FixedCols = 0
            ColCount = 3
			Height = 234
			Width = 486
			ColWidths(0) = 260
			Cell(0,0) = " Concepto"
			Cell(1,0) = " Ingreso"
			Cell(2,0) = " Gasto"
			OnDrawCell = ConceptsBriefDetailStringGridDrawCell
        END CREATE
    END CREATE
    CREATE ConceptsBriefSubtotalsPanel AS QPANEL
        Left = 0
        Top = 350
        Height = 70
        Align = 2
        TabOrder = 3
        CREATE ConceptsBriefSubtotalsTotalsLabel AS QLABEL
            Caption = "Totales:"
            Left = 238
            Top = 17
			Transparent = 1
			Font.AddStyles(fsBold)
        END CREATE
		CREATE ConceptsBriefSubtotalsTotalsRevenuesRichEdit AS QRICHEDIT
			Alignment = 2
            Text = ""
            Left = 288
            Top = 15
			Height = 21
            Width = 100
			TabOrder = 0
			PlainText = 1
			ReadOnly = True
			Font.AddStyles(fsBold)
			Font.Color = QBColor(9)
        END CREATE
		CREATE ConceptsBriefSubtotalsTotalsExpensesRichEdit AS QRICHEDIT
			Alignment = 2
            Text = ""
            Left = 392 '412
            Top = 15
			Height = 21
            Width = 100
            TabOrder = 1
			PlainText = 1
			ReadOnly = True
			Font.AddStyles(fsBold)
			Font.Color = QBColor(12)
        END CREATE
		CREATE ConceptsBriefSubtotalsTotalsRevenuesLabel AS QLABEL
            Caption = "Ingresos"
            Left = 313
            Top = 40
			Transparent = 1
			Font.AddStyles(fsBold)
			Font.Color = QBColor(9)
        END CREATE
        CREATE ConceptsBriefSubtotalsTotalsExpensesLabel AS QLABEL
            Caption = "Gastos"
            Left = 425
            Top = 40
            Transparent = 1
			Font.AddStyles(fsBold)
			Font.Color = QBColor(12)
        END CREATE
    END CREATE
END CREATE

'--------- Subroutines ---------

SUB ConceptsBriefMainFormClose (Action AS INTEGER)
    MainDBFTable.Close
END SUB

SUB ConceptsBriefParametersQueryButtonClick (Sender AS QBUTTON)
	If Val(ConceptsBriefParametersMonthEdit.Text) <= 0 Or Val(ConceptsBriefParametersMonthEdit.Text) > 12 Then
		MessageDlg("El mes debe ser un número positivo entre 01 y 12, introduzca un valor correcto y vuelva a intentarlo.", mtInformation, mbOk, 0)
		Exit Sub
	End If
	
	If Val(ConceptsBriefParametersYearEdit.Text) < 1980 Or Val(ConceptsBriefParametersYearEdit.Text) > Val(Right$(Date$, 4)) Then
		MessageDlg("El año debe ser un número positivo entre 1980 y el año en curso, introduzca un valor correcto y vuelva a intentarlo.", mtInformation, mbOk, 0)
		Exit Sub
	End If
	
	Select Case Val(ConceptsBriefParametersMonthEdit.Text)
		Case 4, 6, 9, 11                  
			Select Case Val(ConceptsBriefParametersDayEdit.Text)
				Case 1 TO 30
				Case Else
					MessageDlg("El día debe ser un número positivo entre 01 y 30, introduzca un valor correcto y vuelva a intentarlo.", mtInformation, mbOk, 0)
					Exit Sub
			End Select
		Case 1, 3, 5, 7, 8, 10, 12
			Select Case Val(ConceptsBriefParametersDayEdit.Text)
				Case 1 TO 31
				Case Else
					MessageDlg("El día debe ser un número positivo entre 01 y 31, introduzca un valor correcto y vuelva a intentarlo.", mtInformation, mbOk, 0)
					Exit Sub
			End Select
		Case 2                            
			Select Case Val(ConceptsBriefParametersYearEdit.Text) MOD 4      
				Case 0  
					Select Case Val(ConceptsBriefParametersDayEdit.Text)
						Case 1 TO 29
						Case Else
							MessageDlg("El día debe ser un número positivo entre 01 y 29, introduzca un valor correcto y vuelva a intentarlo.", mtInformation, mbOk, 0)
							Exit Sub		
					End Select
				Case Else
					Select Case Val(ConceptsBriefParametersDayEdit.Text)
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
	
	If ConceptsBriefParametersPeriodComboBox.ItemIndex = -1 Then
		MessageDlg("Es necesario seleccionar un período en la lista desplegable.", mtInformation, mbOk, 0)
		Exit Sub
	End If
	
	Dim Day As String, Month As String, Year As String
	
	Day = IIF(Val(ConceptsBriefParametersDayEdit.Text) < 10, "0" + Str$(Val(ConceptsBriefParametersDayEdit.Text)), Str$(Val(ConceptsBriefParametersDayEdit.Text)))
	Month = IIF(Val(ConceptsBriefParametersMonthEdit.Text) < 10, "0" + Str$(Val(ConceptsBriefParametersMonthEdit.Text)), Str$(Val(ConceptsBriefParametersMonthEdit.Text)))
	Year = ConceptsBriefParametersYearEdit.Text
	
	Dim PeriodRevenues As Double, PeriodExpenses As Double
	Dim TotalPeriodRevenues As Double, TotalPeriodExpenses As Double
	Dim ConceptID As Word, ConceptKind As Word
	Dim MovementDate As String
	Dim Row As Word
	
	TotalPeriodRevenues = 0
	TotalPeriodExpenses = 0
	
	Row = 0
	
	Dim StartDate As String, EndDate As String
	
	Select Case ConceptsBriefParametersPeriodComboBox.ItemIndex
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
	
	If AuxDBFTable.Connect(CurDir$ + "\data\" + "MOVMENTS.DBF") = False Then
		MessageDlg("La tabla 'MOVMENTS.DBF' no es válida o no se encuentra en la carpeta 'data' del programa.", mtError, mbOk, 0)
		Exit Sub
	End If
	
	ConceptsBriefDetailStringGrid.RowCount = 1
	With MainDBFTable
		.GoTop
		While .Fetch
			ConceptID = Val(.FieldByName("ID"))
			ConceptKind = Val(.FieldByName("KIND"))
			PeriodRevenues = 0
			PeriodExpenses = 0
			AuxDBFTable.GoTop
			While AuxDBFTable.Fetch
				MovementDate = AuxDBFTable.FieldByName("DATE")
				If (MovementDate <= EndDate) And (MovementDate >= StartDate) And (ConceptID = Val(AuxDBFTable.FieldByName("CONCEPT"))) Then
					Select Case ConceptKind 
						Case 1
							PeriodRevenues += Val(AuxDBFTable.FieldByName("AMOUNT"))
						Case 2
							PeriodExpenses += Val(AuxDBFTable.FieldByName("AMOUNT"))
					End Select	
				End If
				AuxDBFTable.Skip(1)
			Wend
			Row += 1
			ConceptsBriefDetailStringGrid.InsertRow(Row)
			ConceptsBriefDetailStringGrid.Cell(0, Row) = .FieldByName("DESC")
			Select Case ConceptKind
				Case 1
					TotalPeriodRevenues += PeriodRevenues
					ConceptsBriefDetailStringGrid.Cell(1, Row) = Format$("%.2f", PeriodRevenues)
				Case 2
					TotalPeriodExpenses += PeriodExpenses
					ConceptsBriefDetailStringGrid.Cell(2, Row) = Format$("%.2f", PeriodExpenses)
			End Select
			.Skip(1)
			DoEvents
		Wend
	End With
	
	AuxDBFTable.Close
	
	If ConceptsBriefDetailStringGrid.RowCount > 1 Then
		ConceptsBriefDetailStringGrid.FixedRows = 1	
	End If
	
	ConceptsBriefSubtotalsTotalsRevenuesRichEdit.Text = Format$("%.2f", TotalPeriodRevenues)
	ConceptsBriefSubtotalsTotalsExpensesRichEdit.Text = Format$("%.2f", TotalPeriodExpenses)
	
END SUB

SUB ConceptsBriefDetailStringGridDrawCell (Col As Integer, Row As Integer, State As Integer, Rect As QRect, Sender As QStringGrid)
	If Row > 0 Then
		Sender.FillRect(Rect.Left, Rect.Top, Rect.Right, Rect.Bottom, QBColor(15))
		Select Case Col
			Case 1
				Sender.TextOut(Rect.Left + ((Rect.Right - Rect.Left) - Sender.TextWidth(Sender.Cell(Col, Row)) - 2), Rect.Top + 2, Sender.Cell(Col, Row), QBColor(9), -1)
			Case 2
				Sender.TextOut(Rect.Left + ((Rect.Right - Rect.Left) - Sender.TextWidth(Sender.Cell(Col, Row)) - 2), Rect.Top + 2, Sender.Cell(Col, Row), QBColor(12), -1)
			Case Else
				Sender.TextOut(Rect.Left + 2, Rect.Top + 2, Sender.Cell(Col, Row), 0, -1)
		End Select
	Else
		'Sender.TextOut(Rect.Left + 2, Rect.Top + 2, Sender.Cell(Col, Row), 0, -1)
	End If
END SUB


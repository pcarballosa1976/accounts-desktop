'=====================================================================================
'
'  Revenues.bas - Formulario para la entrada de los ingresos.
'  
'  El formulario permite introducir los ingresos por día teniendo en cuenta los
'  distintos conceptos.
'
'  Copyright (C) 2022 Cuban Solutions
'
'  Este programa es gratuito y está publicado en https://cubansolutions.blogspot.com
'  para su descarga. 
'
'=====================================================================================

DECLARE SUB RevenuesMainFormClose (Action AS INTEGER)
DECLARE SUB RevenuesAcceptButtonClick (Sender AS QBUTTON)
DECLARE SUB RevenuesCloseButtonClick (Sender AS QBUTTON)

CREATE RevenuesMainForm AS QFORM
    Caption = "Ingresos"
    Width = 455
    Height = 294
    BorderStyle = bsDialog
    Center
    OnClose = RevenuesMainFormClose
    CREATE RevenuesTopPanel AS QPANEL
        Left = 0
        Top = 0
        Height = 60
        Align = alTop
        TabOrder = 0
        CREATE RenevuesTopCommentTitleLabel AS QLABEL
            Caption = "Ingresos de efectivo"
            Left = 15
            Top = 13
            Width = 200
			Font.AddStyles(fsBold)
        END CREATE
        CREATE RevenuesTopCommentTextLabel AS QLABEL
            Caption = "La entrada de los ingresos diarios según los distintos conceptos predefinidos."
            Left = 15
            Top = 30
            Width = 435
            Height = 62
            Wordwrap = 1
        END CREATE
    END CREATE
    CREATE RevenuesMainPanel AS QPANEL
        Left = 0
        Top = 82
        Align = alClient
        TabOrder = 1
        CREATE RevenuesDateLabel AS QLABEL
            Caption = "Fecha:"
            Left = 136
            Top = 20
        END CREATE
        CREATE RevenuesDayEdit AS QEDIT
            Text = ""
            Left = 171
            Top = 18
            Width = 20
			InputMask = "00!"
            MaxLength = 2
            TabOrder = 0
        END CREATE
        CREATE RevenuesMonthEdit AS QEDIT
            Text = ""
            Left = 197
            Top = 18
            Width = 20
			InputMask = "00!"
            MaxLength = 2
            TabOrder = 1
        END CREATE
		CREATE RevenuesYearEdit AS QEDIT
            Text = ""
            Left = 223
            Top = 18
            Width = 32
			InputMask = "0000!"
            MaxLength = 4
            TabOrder = 2
        END CREATE
		CREATE RevenuesConceptLabel AS QLABEL
            Caption = "Concepto:"
            Left = 120
            Top = 56
        END CREATE
        CREATE RevenuesConceptComboBox AS QCOMBOBOX
            Left = 171
            Top = 54
            Width = 200
			Sorted = True
			Style = csDropDownList 
            TabOrder = 3
        END CREATE
		CREATE RevenuesDetailLabel AS QLABEL
            Caption = "Detalle:"
            Left = 133
            Top = 92
        END CREATE
        CREATE RevenuesDetailEdit AS QEDIT
            Text = ""
			Left = 171
            Top = 90
            Width = 200
			MaxLength = 100
            TabOrder = 4
        END CREATE
		CREATE RevenuesAmountLabel AS QLABEL
            Caption = "Importe:"
            Left = 131
            Top = 128
        END CREATE
        CREATE RevenuesAmountEdit AS QEDIT
            Text = ""
			Left = 171
            Top = 126
            Width = 64
			InputMask = "#######.##!"
			MaxLength = 10
            TabOrder = 5
        END CREATE
    END CREATE
	CREATE RevenuesToolsPanel AS QPANEL
        Left = 0
        Top = 282
        Height = 40
        Align = alBottom
        TabOrder = 2
        CREATE RevenuesAcceptButton AS QBUTTON
            Caption = "&Aceptar"
            Left = 300
            Top = 8
            Width = 60
            Height = 24
			Default = True
            ShowHint = 1
			TabOrder = 0
            Hint = "Aceptar entrada"
            OnClick = RevenuesAcceptButtonClick
        END CREATE
		CREATE RevenuesCloseButton AS QBUTTON
            Caption = "&Cerrar"
            Left = 370
            Top = 8
            Width = 60
            Height = 24
			Cancel = True
            ShowHint = 1
			TabOrder = 1
            Hint = "Cerrar ventana"
            OnClick = RevenuesCloseButtonClick
        END CREATE
    END CREATE
END CREATE

'--------- Subroutines ---------

SUB RevenuesMainFormClose (Action AS INTEGER)
    MainDBFTable.Close
END SUB

SUB RevenuesAcceptButtonClick (Sender AS QBUTTON)
	
	If Val(RevenuesMonthEdit.Text) <= 0 Or Val(RevenuesMonthEdit.Text) > 12 Then
		MessageDlg("El mes debe ser un número positivo entre 01 y 12, introduzca un valor correcto y vuelva a intentarlo.", mtInformation, mbOk, 0)
		Exit Sub
	End If
	
	If Val(RevenuesYearEdit.Text) < 1980 Or Val(RevenuesYearEdit.Text) > Val(Right$(Date$, 4)) Then
		MessageDlg("El año debe ser un número positivo entre 1980 y el año en curso, introduzca un valor correcto y vuelva a intentarlo.", mtInformation, mbOk, 0)
		Exit Sub
	End If
	
	Select Case Val(RevenuesMonthEdit.Text)
		Case 4, 6, 9, 11                  
			Select Case Val(RevenuesDayEdit.Text)
				Case 1 TO 30
				Case Else
					MessageDlg("El día debe ser un número positivo entre 01 y 30, introduzca un valor correcto y vuelva a intentarlo.", mtInformation, mbOk, 0)
					Exit Sub
			End Select
		Case 1, 3, 5, 7, 8, 10, 12
			Select Case Val(RevenuesDayEdit.Text)
				Case 1 TO 31
				Case Else
					MessageDlg("El día debe ser un número positivo entre 01 y 31, introduzca un valor correcto y vuelva a intentarlo.", mtInformation, mbOk, 0)
					Exit Sub
			End Select
		Case 2                            
			Select Case Val(RevenuesYearEdit.Text) MOD 4      
				Case 0  
					Select Case Val(RevenuesDayEdit.Text)
						Case 1 TO 29
						Case Else
							MessageDlg("El día debe ser un número positivo entre 01 y 29, introduzca un valor correcto y vuelva a intentarlo.", mtInformation, mbOk, 0)
							Exit Sub		
					End Select
				Case Else
					Select Case Val(RevenuesDayEdit.Text)
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
		
	If RevenuesConceptComboBox.ItemIndex = -1 Then
		MessageDlg("Es necesario seleccionar un concepto en la lista desplegable, en caso de no existir ninguno introduzca uno usando el menú correspondiente del programa.", mtInformation, mbOk, 0)
		Exit Sub
	End If
	
	If Val(RevenuesAmountEdit.Text) <= 0 Then
		MessageDlg("El importe debe ser un número positivo, introduzca un valor correcto y vuelva a intentarlo.", mtInformation, mbOk, 0)
		Exit Sub
	End If
	
	Dim Day As String, Month As String, Year As String
	
	Day = IIF(Val(RevenuesDayEdit.Text) < 10, "0" + Str$(Val(RevenuesDayEdit.Text)), Str$(Val(RevenuesDayEdit.Text)))
	Month = IIF(Val(RevenuesMonthEdit.Text) < 10, "0" + Str$(Val(RevenuesMonthEdit.Text)), Str$(Val(RevenuesMonthEdit.Text)))
	Year = RevenuesYearEdit.Text
	
	Dim ConceptID As String
	
	With AuxDBFTable 
		If .Connect(CurDir$ + "\data\" + "CONCEPTS.DBF") = False Then
			MessageDlg("La tabla 'CONCEPTS.DBF' no es válida o no se encuentra en la carpeta 'data' del programa.", mtError, mbOk, 0)
			Exit Sub
		End If
		.GoTop
		If .Locate("DESC", RevenuesConceptComboBox.Text, False) = False Then
			MessageDlg("El identificador de concepto no pudo ser localizado, seleccione un concepto en la lista desplegable y vuelva a intentarlo.", mtInformation, mbOk, 0)
			Exit Sub
		End If
		ConceptID = .FieldByName("ID")
		.Close
	End With
	
	Dim Consecutive As Word
	
    With MainDBFTable
		If .Append = False Then
			MessageDlg("La tabla 'MOVMENTS.DBF' no puedo ponerse en modo de inserción... no se continuará con la operación en curso." , mtError, mbOk, 0)
			Exit Sub
		End If
		If AuxDBFTable.Connect(CurDir$ + "\data\" + "CONSECUT.DBF") = False Then
			MessageDlg("La tabla 'CONSECUT.DBF' no es válida o no se encuentra en la carpeta 'data' del programa.", mtError, mbOk, 0)
			.Cancel
			Exit Sub
		End If
		Consecutive = Val(AuxDBFTable.FieldByName("MOVMENTS"))
        .SetField(1, Str$(Consecutive))
        If AuxDBFTable.Edit = False Then
            MessageDlg("La tabla 'CONSECUT.DBF' no pudo ponerse en modo de edición para actualizar el consecutivo... no se continuará con la operación en curso.", mtError, mbOk, 0)
			AuxDBFTable.Close
			.Cancel
            Exit Sub    
        End If
        AuxDBFTable.SetField(2, Str$(Consecutive + 1))
        If AuxDBFTable.Save = False Then
            MessageDlg("El consecutivo no pudo guardarse en la tabla 'CONSECUT.DBF'... no se continuará con la operación en curso.", mtError, mbOk, 0)
			AuxDBFTable.Close
			.Cancel
            Exit Sub    
        End If
        AuxDBFTable.Close
		.SetField(2, Year + Month + Day)
		.SetField(3, ConceptID)
		.SetField(4, RevenuesDetailEdit.Text)
		.SetField(5, RevenuesAmountEdit.Text)
		If .Save = False Then
			MessageDlg("El ingreso no pudo ser guardado con éxito... no se continuará con la operación en curso.", mtError, mbOk, 0)
			.Cancel
            Exit Sub    
		End If
		MessageDlg("El ingreso fue guardado con éxito.", mtInformation, mbOk, 0)
    End With
END SUB

SUB RevenuesCloseButtonClick (Sender AS QBUTTON)
	RevenuesMainForm.Close
END SUB




'====================================================================================
'
'  Concepts.bas - Formulario para la manipulación de los conceptos de gastos
'  e ingresos.
'  
'  El formulario permite introducir, editar, y eliminar los diferentes conceptos de
'  gastos e ingresos a tener en cuenta en los movimientos de efectivo; en caso de
'  haber usado un concepto previamente definido en uno de los movimientos de gastos
'  o ingresos luego no podrá eliminarlo. 
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

DECLARE SUB ConceptsMainFormClose (Action AS INTEGER)
DECLARE SUB ConceptsFirstCoolBtnClick (Sender AS QCOOLBTN)
DECLARE SUB ConceptsPreviousCoolBtnClick (Sender AS QCOOLBTN)
DECLARE SUB ConceptsNextCoolBtnClick (Sender AS QCOOLBTN)
DECLARE SUB ConceptsLastCoolBtnClick (Sender AS QCOOLBTN)
DECLARE SUB ConceptsInsertCoolBtnClick (Sender AS QCOOLBTN)
DECLARE SUB ConceptsDeleteCoolBtnClick (Sender AS QCOOLBTN)
DECLARE SUB ConceptsEditCoolBtnClick (Sender AS QCOOLBTN)
DECLARE SUB ConceptsSaveCoolBtnClick (Sender AS QCOOLBTN)
DECLARE SUB ConceptsCancelCoolBtnClick (Sender AS QCOOLBTN)
DECLARE SUB ConceptsFindCoolBtnClick (Sender AS QCOOLBTN)
DECLARE SUB ConceptsRefreshCoolBtnClick (Sender AS QCOOLBTN)

CREATE ConceptsMainForm AS QFORM
    Caption = "Conceptos"
    Width = 455
    Height = 266
    BorderStyle = bsDialog
    Center
    OnClose = ConceptsMainFormClose
    CREATE ConceptsTopPanel AS QPANEL
        Left = 0
        Top = 0
        Height = 60
        Align = alTop
        TabOrder = 0
        CREATE ConceptsTopCommentTitleLabel AS QLABEL
            Caption = "Conceptos de gastos e ingresos"
            Left = 15
            Top = 13
            Width = 200
			Font.AddStyles(fsBold)
        END CREATE
        CREATE ConceptsTopCommentTextLabel AS QLABEL
            Caption = "La definición de los conceptos de gastos e ingresos para ser usados por los movimientos."
            Left = 15
            Top = 30
            Width = 440
            Height = 62
            Wordwrap = 1
        END CREATE
    END CREATE
    CREATE ConceptsMainPanel AS QPANEL
        Left = 0
        Top = 82
        Align = alClient
        TabOrder = 1
        CREATE ConceptsDescriptionLabel AS QLABEL
            Caption = "Descripción:"
            Left = 110
            Top = 20
        END CREATE
        CREATE ConceptsDescriptionEdit AS QEDIT
            Text = ""
            Left = 171
            Top = 18
            Width = 193
            MaxLength = 100
            TabOrder = 0
            ReadOnly = True
        END CREATE
		CREATE ConceptsCharacterGroupBox AS QGROUPBOX
			Caption = "Tipo"
			Left = 171
			Top = 48
			Width = 121
			Height = 73
			TabOrder = 1
			Enabled = False
			CREATE ConceptsRevenuesRadioButton AS QRADIOBUTTON
				Caption = "Ingresos"
				Left = 14
				Top = 21
				Width = 100
				Checked = 1
				TabOrder = 0
			END CREATE
			CREATE ConceptsExpensesRadioButton AS QRADIOBUTTON
				Caption = "Gastos"
				Left = 14
				Top = 44
				Width = 100
				TabOrder = 1
			END CREATE
		END CREATE
    END CREATE
	CREATE ConceptsToolsPanel AS QPANEL
        Left = 0
        Top = 282
        Height = 40
        Align = alBottom
        TabOrder = 2
        CREATE ConceptsFirstCoolBtn AS QCOOLBTN
            BMPHandle = NAVTOOLS_FIRST
            Left = 150
            Top = 10
            Width = 24
            Height = 22
            ShowHint = 1
            Hint = "Primero"
            NumBMPs = 2
            OnClick = ConceptsFirstCoolBtnClick
        END CREATE
        CREATE ConceptsPreviousCoolBtn AS QCOOLBTN
            BMPHandle = NAVTOOLS_PRIOR
            Left = 176
            Top = 10
            Width = 24
            Height = 22
            ShowHint = 1
            Hint = "Anterior"
            NumBMPs = 2
            OnClick = ConceptsPreviousCoolBtnClick
        END CREATE
        CREATE ConceptsNextCoolBtn AS QCOOLBTN
            BMPHandle = NAVTOOLS_NEXT
            Left = 202
            Top = 10
            Width = 24
            Height = 22
            ShowHint = 1
            Hint = "Siguiente"
            NumBMPs = 2
            OnClick = ConceptsNextCoolBtnClick
        END CREATE
        CREATE ConceptsLastCoolBtn AS QCOOLBTN
            BMPHandle = NAVTOOLS_LAST
            Left = 228
            Top = 10
            Width = 24
            Height = 22
            ShowHint = 1
            Hint = "Último"
            NumBMPs = 2
            OnClick = ConceptsLastCoolBtnClick
        END CREATE
        CREATE ConceptsInsertCoolBtn AS QCOOLBTN
            BMPHandle = NAVTOOLS_ADD
            Left = 254
            Top = 10
            Width = 24
            Height = 22
            ShowHint = 1
            Hint = "Insertar"
            NumBMPs = 2
            OnClick = ConceptsInsertCoolBtnClick
        END CREATE
        CREATE ConceptsDeleteCoolBtn AS QCOOLBTN
            BMPHandle = NAVTOOLS_DELETE
            Left = 280
            Top = 10
            Width = 24
            Height = 22
            ShowHint = 1
            Hint = "Eliminar"
            NumBMPs = 2
            OnClick = ConceptsDeleteCoolBtnClick
        END CREATE
        CREATE ConceptsEditCoolBtn AS QCOOLBTN
            BMPHandle = NAVTOOLS_EDIT
            Left = 306
            Top = 10
            Width = 24
            Height = 22
            ShowHint = 1
            Hint = "Editar"
            NumBMPs = 2
            OnClick = ConceptsEditCoolBtnClick
        END CREATE
        CREATE ConceptsSaveCoolBtn AS QCOOLBTN
            BMPHandle = NAVTOOLS_SAVE
            Left = 332
            Top = 10
            Width = 24
            Height = 22
            ShowHint = 1
            Hint = "Salvar"
            NumBMPs = 2
            OnClick = ConceptsSaveCoolBtnClick
        END CREATE
        CREATE ConceptsCancelCoolBtn AS QCOOLBTN
            BMPHandle = NAVTOOLS_CANCEL
            Left = 358
            Top = 10
            Width = 24
            Height = 22
            ShowHint = 1
            Hint = "Cancelar"
            NumBMPs = 2
            OnClick = ConceptsCancelCoolBtnClick
        END CREATE
        CREATE ConceptsFindCoolBtn AS QCOOLBTN
            BMPHandle = NAVTOOLS_FIND
            Left = 384
            Top = 10
            Width = 24
            Height = 22
            ShowHint = 1
            Hint = "Buscar"
            NumBMPs = 2
            OnClick = ConceptsFindCoolBtnClick
        END CREATE
        CREATE ConceptsRefreshCoolBtn AS QCOOLBTN
            BMPHandle = NAVTOOLS_REFRESH
            Left = 410
            Top = 10
            Width = 24
            Height = 22
            ShowHint = 1
            Hint = "Actualizar"
            NumBMPs = 2
            OnClick = ConceptsRefreshCoolBtnClick
        END CREATE
    END CREATE
END CREATE

'--------- Subroutines ---------

SUB ConceptsMainFormClose (Action AS INTEGER)
    MainDBFTable.Close
END SUB

SUB ConceptsFirstCoolBtnClick (Sender AS QCOOLBTN)
    With MainDBFTable
        If .First = True Then
            ConceptsDescriptionEdit.Text = .FieldByName("DESC")
			Select Case Val(.FieldByName("KIND"))
				Case 1
					ConceptsRevenuesRadioButton.Checked = True
				Case 2
					ConceptsExpensesRadioButton.Checked = True
			End Select
        Else
            .LastError(MainDBFErr)
            If MainDBFErr.Id = QDBF_ERR_NONE Then
				ConceptsDescriptionEdit.Text = ""
            Else
                MessageDlg(MainDBFErr.Description, mtError, mbOk, 0)
            End If
        End If
    End With
END SUB

SUB ConceptsPreviousCoolBtnClick (Sender AS QCOOLBTN)
    With MainDBFTable
        If .Prior = True Then
            ConceptsDescriptionEdit.Text = .FieldByName("DESC")
			Select Case Val(.FieldByName("KIND"))
				Case 1
					ConceptsRevenuesRadioButton.Checked = True
				Case 2
					ConceptsExpensesRadioButton.Checked = True
			End Select
        Else
            .LastError(MainDBFErr)
            If MainDBFErr.Id = QDBF_ERR_NONE Then
				ConceptsDescriptionEdit.Text = ""
            Else
                MessageDlg(MainDBFErr.Description, mtError, mbOk, 0)
            End If
        End If
    End With
END SUB

SUB ConceptsNextCoolBtnClick (Sender AS QCOOLBTN)
    With MainDBFTable
        If .Next = True Then
            ConceptsDescriptionEdit.Text = .FieldByName("DESC")
			Select Case Val(.FieldByName("KIND"))
				Case 1
					ConceptsRevenuesRadioButton.Checked = True
				Case 2
					ConceptsExpensesRadioButton.Checked = True
			End Select
        Else
            .LastError(MainDBFErr)
            If MainDBFErr.Id = QDBF_ERR_NONE Then
				ConceptsDescriptionEdit.Text = ""
            Else
                MessageDlg(MainDBFErr.Description, mtError, mbOk, 0)
            End If
        End If
    End With

END SUB

SUB ConceptsLastCoolBtnClick (Sender AS QCOOLBTN)
    With MainDBFTable
        If .Last = True Then
            ConceptsDescriptionEdit.Text = .FieldByName("DESC")
			Select Case Val(.FieldByName("KIND"))
				Case 1
					ConceptsRevenuesRadioButton.Checked = True
				Case 2
					ConceptsExpensesRadioButton.Checked = True
			End Select
        Else
            .LastError(MainDBFErr)
            If MainDBFErr.Id = QDBF_ERR_NONE Then
				ConceptsDescriptionEdit.Text = ""
            Else
                MessageDlg(MainDBFErr.Description, mtError, mbOk, 0)
            End If
        End If
    End With
END SUB

SUB ConceptsInsertCoolBtnClick (Sender AS QCOOLBTN)
    With MainDBFTable
        If .Append = True Then
            ConceptsFirstCoolBtn.Enabled = False
            ConceptsPreviousCoolBtn.Enabled = False
            ConceptsNextCoolBtn.Enabled = False
            ConceptsLastCoolBtn.Enabled = False
            ConceptsInsertCoolBtn.Enabled = False
            ConceptsDeleteCoolBtn.Enabled = False
            ConceptsEditCoolBtn.Enabled = False
            ConceptsSaveCoolBtn.Enabled = True
            ConceptsCancelCoolBtn.Enabled = True
            ConceptsFindCoolBtn.Enabled = False
            ConceptsRefreshCoolBtn.Enabled = False
            
            ConceptsDescriptionEdit.Text = ""
            ConceptsDescriptionEdit.ReadOnly = False
			ConceptsCharacterGroupBox.Enabled = True
        Else
            .LastError(MainDBFErr)
            MessageDlg(MainDBFErr.Description, mtError, mbOk, 0)
        End If
    End With
END SUB

SUB ConceptsDeleteCoolBtnClick (Sender AS QCOOLBTN)
	Dim ConceptID As String
    If MessageDlg("¿Está seguro de eliminar este concepto?" + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10) + "En caso de existir movimientos usando este concepto la eliminación no será posible." , mtConfirmation, mbYes Or mbNo, 0) = mrYes Then
        With AuxDBFTable
			If .Connect(CurDir$ + "\data\" + "MOVMENTS.DBF") = False Then
				MessageDlg("La tabla 'MOVMENTS.DBF' no es válida o no se encuentra en la carpeta 'data' del programa.", mtError, mbOk, 0)
				Exit Sub
			End If
			ConceptID = MainDBFTable.FieldByName("ID")
			.GoTop
			If .Locate("CONCEPT", ConceptID, False) = True Then
				MessageDlg("El concepto está siendo utilizado en los movimientos y por esto no puede ser eliminado... no se continuará con la operación en curso.", mtInformation, mbOk, 0)
				.Close
				Exit Sub
			End If
			.Close
		End With
		
        With MainDBFTable
            If .Delete = True Then
                MessageDlg("El concepto se ha eliminado con éxito.", mtInformation, mbOk, 0)    
                If .Next = True Then
                    ConceptsDescriptionEdit.Text = .FieldByName("DESC")
					Select Case Val(.FieldByName("KIND"))
						Case 1
							ConceptsRevenuesRadioButton.Checked = True
						Case 2
							ConceptsExpensesRadioButton.Checked = True
					End Select
				Else
					.LastError(MainDBFErr)
					If MainDBFErr.Id = QDBF_ERR_NONE Then
						ConceptsDescriptionEdit.Text = ""
					Else
						MessageDlg(MainDBFErr.Description, mtError, mbOk, 0)
					End If
					ConceptsFirstCoolBtn.Enabled = False
					ConceptsPreviousCoolBtn.Enabled = False
					ConceptsNextCoolBtn.Enabled = False
					ConceptsLastCoolBtn.Enabled = False
					ConceptsInsertCoolBtn.Enabled = True
					ConceptsDeleteCoolBtn.Enabled = False
					ConceptsEditCoolBtn.Enabled = False
					ConceptsSaveCoolBtn.Enabled = False
					ConceptsCancelCoolBtn.Enabled = False
					ConceptsFindCoolBtn.Enabled = False
					ConceptsRefreshCoolBtn.Enabled = False
					ConceptsCharacterGroupBox.Enabled = False
				End If
            Else
                .LastError(MainDBFErr)
                MessageDlg(MainDBFErr.Description, mtError, mbOk, 0)
            End If
        End With
    End If   
END SUB

SUB ConceptsEditCoolBtnClick (Sender AS QCOOLBTN)
    With MainDBFTable
        If .Edit = True Then
            ConceptsFirstCoolBtn.Enabled = False
            ConceptsPreviousCoolBtn.Enabled = False
            ConceptsNextCoolBtn.Enabled = False
            ConceptsLastCoolBtn.Enabled = False
            ConceptsInsertCoolBtn.Enabled = False
            ConceptsDeleteCoolBtn.Enabled = False
            ConceptsEditCoolBtn.Enabled = False
            ConceptsSaveCoolBtn.Enabled = True
            ConceptsCancelCoolBtn.Enabled = True
            ConceptsFindCoolBtn.Enabled = False
            ConceptsRefreshCoolBtn.Enabled = False
            
            ConceptsDescriptionEdit.ReadOnly = False
			ConceptsCharacterGroupBox.Enabled = True
        Else
            .LastError(MainDBFErr)
            MessageDlg(MainDBFErr.Description, mtError, mbOk, 0)
        End If
    End With
END SUB

SUB ConceptsSaveCoolBtnClick (Sender AS QCOOLBTN)
    If ConceptsDescriptionEdit.Text = "" Then
        MessageDlg("La descripción no puede estar vacía.", mtInformation, mbOk, 0)
        Exit Sub
    End If
	
	Dim ConceptDESC As String
	With AuxDBFTable
		If .Connect(CurDir$ + "\data\" + "CONCEPTS.DBF") = False Then
			MessageDlg("La tabla 'CONCEPTS.DBF' no es válida o no se encuentra en la carpeta 'data' del programa.", mtError, mbOk, 0)
			Exit Sub
		End If
		ConceptDESC = ConceptsDescriptionEdit.Text
		.GoTop
		If .Locate("DESC", ConceptDESC, False) = True Then
			Select Case MainDBFTable.Mode
				Case QDBF_APPEND_MODE
					MessageDlg("La descripción especificada está siendo usada... no se continuará con la operación en curso.", mtInformation, mbOk, 0)
					.Close
					Exit Sub
				Case QDBF_EDIT_MODE
					If MainDBFTable.RecordNo <> .RecordNo Then
						MessageDlg("La descripción especificada está siendo usada... no se continuará con la operación en curso.", mtInformation, mbOk, 0)
						.Close
						Exit Sub
					End If
				Case Else
					MessageDlg("Es necesario estar en modo inserción o en modo de edición... no se continuará con la operación en curso.", mtInformation, mbOk, 0)
					.Close
					Exit Sub
			End Select
		End If
		.Close
	End With

    Dim Consecutive As Word

    With MainDBFTable
        If .Mode = QDBF_APPEND_MODE Then
            If AuxDBFTable.Connect(CurDir$ + "\data\" + "CONSECUT.DBF") = False Then
                MessageDlg("La tabla 'CONSECUT.DBF' no es válida o no se encuentra en la carpeta 'data' del programa... no se continuará con la operación en curso.", mtError, mbOk, 0)
                Exit Sub
            End If
            Consecutive = Val(AuxDBFTable.FieldByName("Concepts"))
            .SetField(1, Str$(Consecutive))
            If AuxDBFTable.Edit = False Then
                MessageDlg("La tabla 'CONSECUT.DBF' no pudo ponerse en modo de edición para actualizar el consecutivo... no se continuará con la operación en curso.", mtError, mbOk, 0)
				AuxDBFTable.Close
                Exit Sub    
            End If
            AuxDBFTable.SetField(1, Str$(Consecutive + 1))
            If AuxDBFTable.Save = False Then
                MessageDlg("El consecutivo no pudo guardarse en la tabla 'CONSECUT.DBF'... no se continuará con la operación en curso.", mtError, mbOk, 0)
				AuxDBFTable.Close
                Exit Sub    
            End If
            AuxDBFTable.Close
        End If
        .SetField(2, ConceptsDescriptionEdit.Text)
		.SetField(3, IIF(ConceptsRevenuesRadioButton.Checked = True, "1", "2"))
        
        If .Save = True Then
            ConceptsFirstCoolBtn.Enabled = True
            ConceptsPreviousCoolBtn.Enabled = True
            ConceptsNextCoolBtn.Enabled = True
            ConceptsLastCoolBtn.Enabled = True
            ConceptsInsertCoolBtn.Enabled = True
            ConceptsDeleteCoolBtn.Enabled = True
            ConceptsEditCoolBtn.Enabled = True
            ConceptsSaveCoolBtn.Enabled = False
            ConceptsCancelCoolBtn.Enabled = False
            ConceptsFindCoolBtn.Enabled = True
            ConceptsRefreshCoolBtn.Enabled = True
            
            ConceptsDescriptionEdit.ReadOnly = True
			ConceptsCharacterGroupBox.Enabled = False
        Else
            .LastError(MainDBFErr)
            MessageDlg(MainDBFErr.Description, mtError, mbOk, 0)
        End If
    End With

END SUB

SUB ConceptsCancelCoolBtnClick (Sender AS QCOOLBTN)
    With MainDBFTable
        If .Cancel = True Then      
            ConceptsDescriptionEdit.Text = .FieldByName("DESC")
                        
            ConceptsFirstCoolBtn.Enabled = True
            ConceptsPreviousCoolBtn.Enabled = True
            ConceptsNextCoolBtn.Enabled = True
            ConceptsLastCoolBtn.Enabled = True
            ConceptsInsertCoolBtn.Enabled = True
            ConceptsDeleteCoolBtn.Enabled = True
            ConceptsEditCoolBtn.Enabled = True
            ConceptsSaveCoolBtn.Enabled = False
            ConceptsCancelCoolBtn.Enabled = False
            ConceptsFindCoolBtn.Enabled = True
            ConceptsRefreshCoolBtn.Enabled = True
            
            ConceptsDescriptionEdit.ReadOnly = True
			ConceptsCharacterGroupBox.Enabled = False
        Else
            .LastError(MainDBFErr)
            MessageDlg(MainDBFErr.Description, mtError, mbOk, 0)
        End If
    End With
END SUB

SUB ConceptsFindCoolBtnClick (Sender AS QCOOLBTN)
    MessageDlg("Esta funcionalidad no se ha implementado en esta versión del sistema.", mtInformation, mbOk, 0)
END SUB

SUB ConceptsRefreshCoolBtnClick (Sender AS QCOOLBTN)
    With MainDBFTable
        If .BOF = False Or .EOF = False Then
			ConceptsDescriptionEdit.Text = .FieldByName("DESC")
			Select Case Val(.FieldByName("KIND"))
				Case 1
					ConceptsRevenuesRadioButton.Checked = True
				Case 2
					ConceptsExpensesRadioButton.Checked = True
			End Select
        Else
            ConceptsDescriptionEdit.Text = ""
        End If
    End With
END SUB


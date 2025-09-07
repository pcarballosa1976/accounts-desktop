'==================================================================================
'
'  Balances.bas - Formulario para la manipulación de los saldos finales de cada año.
'  
'  El formulario permite ver, y también introducir, editar y eliminar los saldos finales de cada año,
'  aun cuando estos dos últimos casos de uso no son recomendables.
'
'  La utilidad real de la opción consiste en permitir la introducción manual del saldo final del año
'  anterior antes de estar usando este programa, con lo cual podemos tener en cuenta el estado
'  de resultados previo a su utilización, puesto así dicho saldo final podría ser utilizado durante la
'  realización de los cálculos del saldo de los movimientos posteriores.
'
'  Copyright (C) 2022 Cuban Solutions
'
'  Este programa es gratuito y está publicado en https://cubansolutions.blogspot.com para
'  su descarga. 
'
'==================================================================================

DECLARE SUB BalancesMainFormClose (Action AS INTEGER)
DECLARE SUB BalancesFirstCoolBtnClick (Sender AS QCOOLBTN)
DECLARE SUB BalancesPreviousCoolBtnClick (Sender AS QCOOLBTN)
DECLARE SUB BalancesNextCoolBtnClick (Sender AS QCOOLBTN)
DECLARE SUB BalancesLastCoolBtnClick (Sender AS QCOOLBTN)
DECLARE SUB BalancesInsertCoolBtnClick (Sender AS QCOOLBTN)
DECLARE SUB BalancesDeleteCoolBtnClick (Sender AS QCOOLBTN)
DECLARE SUB BalancesEditCoolBtnClick (Sender AS QCOOLBTN)
DECLARE SUB BalancesSaveCoolBtnClick (Sender AS QCOOLBTN)
DECLARE SUB BalancesCancelCoolBtnClick (Sender AS QCOOLBTN)
DECLARE SUB BalancesFindCoolBtnClick (Sender AS QCOOLBTN)
DECLARE SUB BalancesRefreshCoolBtnClick (Sender AS QCOOLBTN)

CREATE BalancesMainForm AS QFORM
    Caption = "Saldos"
    Width = 418
    Height = 220
    BorderStyle = bsDialog
    Center
    OnClose = BalancesMainFormClose
    CREATE BalancesTopPanel AS QPANEL
        Left = 0
        Top = 0
        Height = 60
        Align = alTop
        TabOrder = 0
        CREATE BalancesTopCommentTitleLabel AS QLABEL
            Caption = "Saldos finales por año"
            Left = 15
            Top = 13
            Width = 200
			Font.AddStyles(fsBold)
        END CREATE
        CREATE BalancesTopCommentTextLabel AS QLABEL
            Caption = "Los saldos finales por cada año debidos a los movimientos de gastos e ingresos."
            Left = 15
            Top = 30
            Width = 400
            Height = 62
            Wordwrap = 1
        END CREATE
    END CREATE
    CREATE BalancesMainPanel AS QPANEL
        Left = 0
        Top = 82
        Align = alClient
        TabOrder = 1
        CREATE BalancesYearLabel AS QLABEL
            Caption = "Año:"
            Left = 147
            Top = 20
        END CREATE
        CREATE BalancesYearEdit AS QEDIT
            Text = ""
            Left = 171
            Top = 18
            Width = 32
			InputMask = "0000!"
            MaxLength = 4
            TabOrder = 0
            ReadOnly = True
        END CREATE
		CREATE BalancesAmountLabel AS QLABEL
            Caption = "Saldo:"
            Left = 139
            Top = 56
        END CREATE
        CREATE BalancesAmountEdit AS QEDIT
            Text = ""
            Left = 171
            Top = 54
            Width = 88
			InputMask = "###########.##!"
            MaxLength = 4
            TabOrder = 1
            ReadOnly = True
        END CREATE
    END CREATE
	CREATE BalancesToolsPanel AS QPANEL
        Left = 0
        Top = 282
        Height = 40
        Align = alBottom
        TabOrder = 2
        CREATE BalancesFirstCoolBtn AS QCOOLBTN
            BMPHandle = NAVTOOLS_FIRST
            Left = 113
            Top = 10
            Width = 24
            Height = 22
            ShowHint = 1
            Hint = "Primero"
            NumBMPs = 2
            OnClick = BalancesFirstCoolBtnClick
        END CREATE
        CREATE BalancesPreviousCoolBtn AS QCOOLBTN
            BMPHandle = NAVTOOLS_PRIOR
            Left = 139
            Top = 10
            Width = 24
            Height = 22
            ShowHint = 1
            Hint = "Anterior"
            NumBMPs = 2
            OnClick = BalancesPreviousCoolBtnClick
        END CREATE
        CREATE BalancesNextCoolBtn AS QCOOLBTN
            BMPHandle = NAVTOOLS_NEXT
            Left = 165
            Top = 10
            Width = 24
            Height = 22
            ShowHint = 1
            Hint = "Siguiente"
            NumBMPs = 2
            OnClick = BalancesNextCoolBtnClick
        END CREATE
        CREATE BalancesLastCoolBtn AS QCOOLBTN
            BMPHandle = NAVTOOLS_LAST
            Left = 191
            Top = 10
            Width = 24
            Height = 22
            ShowHint = 1
            Hint = "Último"
            NumBMPs = 2
            OnClick = BalancesLastCoolBtnClick
        END CREATE
        CREATE BalancesInsertCoolBtn AS QCOOLBTN
            BMPHandle = NAVTOOLS_ADD
            Left = 217
            Top = 10
            Width = 24
            Height = 22
            ShowHint = 1
            Hint = "Insertar"
            NumBMPs = 2
            OnClick = BalancesInsertCoolBtnClick
        END CREATE
        CREATE BalancesDeleteCoolBtn AS QCOOLBTN
            BMPHandle = NAVTOOLS_DELETE
            Left = 243
            Top = 10
            Width = 24
            Height = 22
            ShowHint = 1
            Hint = "Eliminar"
            NumBMPs = 2
            OnClick = BalancesDeleteCoolBtnClick
        END CREATE
        CREATE BalancesEditCoolBtn AS QCOOLBTN
            BMPHandle = NAVTOOLS_EDIT
            Left = 269
            Top = 10
            Width = 24
            Height = 22
            ShowHint = 1
            Hint = "Editar"
            NumBMPs = 2
            OnClick = BalancesEditCoolBtnClick
        END CREATE
        CREATE BalancesSaveCoolBtn AS QCOOLBTN
            BMPHandle = NAVTOOLS_SAVE
            Left = 295
            Top = 10
            Width = 24
            Height = 22
            ShowHint = 1
            Hint = "Salvar"
            NumBMPs = 2
            OnClick = BalancesSaveCoolBtnClick
        END CREATE
        CREATE BalancesCancelCoolBtn AS QCOOLBTN
            BMPHandle = NAVTOOLS_CANCEL
            Left = 321
            Top = 10
            Width = 24
            Height = 22
            ShowHint = 1
            Hint = "Cancelar"
            NumBMPs = 2
            OnClick = BalancesCancelCoolBtnClick
        END CREATE
        CREATE BalancesFindCoolBtn AS QCOOLBTN
            BMPHandle = NAVTOOLS_FIND
            Left = 347
            Top = 10
            Width = 24
            Height = 22
            ShowHint = 1
            Hint = "Buscar"
            NumBMPs = 2
            OnClick = BalancesFindCoolBtnClick
        END CREATE
        CREATE BalancesRefreshCoolBtn AS QCOOLBTN
            BMPHandle = NAVTOOLS_REFRESH
            Left = 373
            Top = 10
            Width = 24
            Height = 22
            ShowHint = 1
            Hint = "Actualizar"
            NumBMPs = 2
            OnClick = BalancesRefreshCoolBtnClick
        END CREATE
    END CREATE
END CREATE

'--------- Subroutines ---------

SUB BalancesMainFormClose (Action AS INTEGER)
    MainDBFTable.Close
END SUB

SUB BalancesFirstCoolBtnClick (Sender AS QCOOLBTN)
    With MainDBFTable
        If .First = True Then
            BalancesYearEdit.Text = .FieldByName("YEAR")
			BalancesAmountEdit.Text = .FieldByName("AMOUNT")
        Else
            .LastError(MainDBFErr)
            If MainDBFErr.Id = QDBF_ERR_NONE Then
				BalancesYearEdit.Text = ""
				BalancesAmountEdit.Text = ""
            Else
                MessageDlg(MainDBFErr.Description, mtError, mbOk, 0)
            End If
        End If
    End With
END SUB

SUB BalancesPreviousCoolBtnClick (Sender AS QCOOLBTN)
    With MainDBFTable
        If .Prior = True Then
            BalancesYearEdit.Text = .FieldByName("YEAR")
			BalancesAmountEdit.Text = .FieldByName("AMOUNT")
        Else
            .LastError(MainDBFErr)
            If MainDBFErr.Id = QDBF_ERR_NONE Then
				BalancesYearEdit.Text = ""
				BalancesAmountEdit.Text = ""
            Else
                MessageDlg(MainDBFErr.Description, mtError, mbOk, 0)
            End If
        End If
    End With
END SUB

SUB BalancesNextCoolBtnClick (Sender AS QCOOLBTN)
    With MainDBFTable
        If .Next = True Then
            BalancesYearEdit.Text = .FieldByName("YEAR")
			BalancesAmountEdit.Text = .FieldByName("AMOUNT")
        Else
            .LastError(MainDBFErr)
            If MainDBFErr.Id = QDBF_ERR_NONE Then
				BalancesYearEdit.Text = ""
				BalancesAmountEdit.Text = ""
            Else
                MessageDlg(MainDBFErr.Description, mtError, mbOk, 0)
            End If
        End If
    End With

END SUB

SUB BalancesLastCoolBtnClick (Sender AS QCOOLBTN)
    With MainDBFTable
        If .Last = True Then
            BalancesYearEdit.Text = .FieldByName("YEAR")
			BalancesAmountEdit.Text = .FieldByName("AMOUNT")
        Else
            .LastError(MainDBFErr)
            If MainDBFErr.Id = QDBF_ERR_NONE Then
				BalancesYearEdit.Text = ""
				BalancesAmountEdit.Text = ""
            Else
                MessageDlg(MainDBFErr.Description, mtError, mbOk, 0)
            End If
        End If
    End With
END SUB

SUB BalancesInsertCoolBtnClick (Sender AS QCOOLBTN)
    With MainDBFTable
        If .Append = True Then
            BalancesFirstCoolBtn.Enabled = False
            BalancesPreviousCoolBtn.Enabled = False
            BalancesNextCoolBtn.Enabled = False
            BalancesLastCoolBtn.Enabled = False
            BalancesInsertCoolBtn.Enabled = False
            BalancesDeleteCoolBtn.Enabled = False
            BalancesEditCoolBtn.Enabled = False
            BalancesSaveCoolBtn.Enabled = True
            BalancesCancelCoolBtn.Enabled = True
            BalancesFindCoolBtn.Enabled = False
            BalancesRefreshCoolBtn.Enabled = False
            
            BalancesYearEdit.Text = ""
            BalancesYearEdit.ReadOnly = False
			BalancesAmountEdit.Text = ""
			BalancesAmountEdit.ReadOnly = False
        Else
            .LastError(MainDBFErr)
            MessageDlg(MainDBFErr.Description, mtError, mbOk, 0)
        End If
    End With
END SUB

SUB BalancesDeleteCoolBtnClick (Sender AS QCOOLBTN)
	Dim ConceptID As String
    If MessageDlg("¿Está seguro de eliminar este saldo?", mtConfirmation, mbYes Or mbNo, 0) = mrYes Then
        With MainDBFTable
            If .Delete = True Then
                MessageDlg("El saldo se ha eliminado con éxito.", mtInformation, mbOk, 0)    
                If .Next = True Then
                    BalancesYearEdit.Text = .FieldByName("YEAR")
					BalancesAmountEdit.Text = .FieldByName("AMOUNT")
				Else
					.LastError(MainDBFErr)
					If MainDBFErr.Id = QDBF_ERR_NONE Then
						BalancesYearEdit.Text = ""
						BalancesAmountEdit.Text = ""
					Else
						MessageDlg(MainDBFErr.Description, mtError, mbOk, 0)
					End If
					BalancesFirstCoolBtn.Enabled = False
					BalancesPreviousCoolBtn.Enabled = False
					BalancesNextCoolBtn.Enabled = False
					BalancesLastCoolBtn.Enabled = False
					BalancesInsertCoolBtn.Enabled = True
					BalancesDeleteCoolBtn.Enabled = False
					BalancesEditCoolBtn.Enabled = False
					BalancesSaveCoolBtn.Enabled = False
					BalancesCancelCoolBtn.Enabled = False
					BalancesFindCoolBtn.Enabled = False
					BalancesRefreshCoolBtn.Enabled = False
				End If
            Else
                .LastError(MainDBFErr)
                MessageDlg(MainDBFErr.Description, mtError, mbOk, 0)
            End If
        End With
    End If   
END SUB

SUB BalancesEditCoolBtnClick (Sender AS QCOOLBTN)
    With MainDBFTable
        If .Edit = True Then
            BalancesFirstCoolBtn.Enabled = False
            BalancesPreviousCoolBtn.Enabled = False
            BalancesNextCoolBtn.Enabled = False
            BalancesLastCoolBtn.Enabled = False
            BalancesInsertCoolBtn.Enabled = False
            BalancesDeleteCoolBtn.Enabled = False
            BalancesEditCoolBtn.Enabled = False
            BalancesSaveCoolBtn.Enabled = True
            BalancesCancelCoolBtn.Enabled = True
            BalancesFindCoolBtn.Enabled = False
            BalancesRefreshCoolBtn.Enabled = False
            
            BalancesYearEdit.ReadOnly = False
			BalancesAmountEdit.ReadOnly = False
        Else
            .LastError(MainDBFErr)
            MessageDlg(MainDBFErr.Description, mtError, mbOk, 0)
        End If
    End With
END SUB

SUB BalancesSaveCoolBtnClick (Sender AS QCOOLBTN)
    If Val(BalancesYearEdit.Text) < 1980 Or Val(BalancesYearEdit.Text) > Val(Right$(Date$, 4)) Then
		MessageDlg("El año debe ser un número positivo entre 1980 y el año en curso, introduzca un valor correcto y vuelva a intentarlo.", mtInformation, mbOk, 0)
		Exit Sub
	End If
	
	Dim Year As String
	With AuxDBFTable
		If .Connect(CurDir$ + "\data\" + "BALANCES.DBF") = False Then
			MessageDlg("La tabla 'BALANCES.DBF' no es válida o no se encuentra en la carpeta 'data' del programa.", mtError, mbOk, 0)
			Exit Sub
		End If
		Year = BalancesYearEdit.Text
		.GoTop
		If .Locate("YEAR", YEAR, False) = True Then
			Select Case MainDBFTable.Mode
				Case QDBF_APPEND_MODE
					MessageDlg("El año especificado está siendo usado... no se continuará con la operación en curso.", mtInformation, mbOk, 0)
					.Close
					Exit Sub
				Case QDBF_EDIT_MODE
					If MainDBFTable.RecordNo <> .RecordNo Then
						MessageDlg("El año especificado está siendo usado... no se continuará con la operación en curso.", mtInformation, mbOk, 0)
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

    With MainDBFTable
        .SetField(1, BalancesYearEdit.Text)
		.SetField(2, BalancesAmountEdit.Text)
        
        If .Save = True Then
            BalancesFirstCoolBtn.Enabled = True
            BalancesPreviousCoolBtn.Enabled = True
            BalancesNextCoolBtn.Enabled = True
            BalancesLastCoolBtn.Enabled = True
            BalancesInsertCoolBtn.Enabled = True
            BalancesDeleteCoolBtn.Enabled = True
            BalancesEditCoolBtn.Enabled = True
            BalancesSaveCoolBtn.Enabled = False
            BalancesCancelCoolBtn.Enabled = False
            BalancesFindCoolBtn.Enabled = True
            BalancesRefreshCoolBtn.Enabled = True
            
            BalancesYearEdit.ReadOnly = True
			BalancesAmountEdit.ReadOnly = True
        Else
            .LastError(MainDBFErr)
            MessageDlg(MainDBFErr.Description, mtError, mbOk, 0)
        End If
    End With

END SUB

SUB BalancesCancelCoolBtnClick (Sender AS QCOOLBTN)
    With MainDBFTable
        If .Cancel = True Then 
			BalancesYearEdit.Text = .FieldByName("YEAR")
			BalancesAmountEdit.Text = .FieldByName("AMOUNT")
			         
            BalancesFirstCoolBtn.Enabled = True
            BalancesPreviousCoolBtn.Enabled = True
            BalancesNextCoolBtn.Enabled = True
            BalancesLastCoolBtn.Enabled = True
            BalancesInsertCoolBtn.Enabled = True
            BalancesDeleteCoolBtn.Enabled = True
            BalancesEditCoolBtn.Enabled = True
            BalancesSaveCoolBtn.Enabled = False
            BalancesCancelCoolBtn.Enabled = False
            BalancesFindCoolBtn.Enabled = True
            BalancesRefreshCoolBtn.Enabled = True
            
            BalancesYearEdit.ReadOnly = True
			BalancesAmountEdit.ReadOnly = True
        Else
            .LastError(MainDBFErr)
            MessageDlg(MainDBFErr.Description, mtError, mbOk, 0)
        End If
    End With
END SUB

SUB BalancesFindCoolBtnClick (Sender AS QCOOLBTN)
    MessageDlg("Esta funcionalidad no se ha implementado en esta versión del sistema.", mtInformation, mbOk, 0)
END SUB

SUB BalancesRefreshCoolBtnClick (Sender AS QCOOLBTN)
    With MainDBFTable
        If .BOF = False Or .EOF = False Then
			BalancesYearEdit.Text = .FieldByName("YEAR")
			BalancesAmountEdit.Text = .FieldByName("AMOUNT")
        Else
            BalancesYearEdit.Text = ""
			BalancesAmountEdit.Text = ""
        End If
    End With
END SUB


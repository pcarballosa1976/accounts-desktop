'=====================================================================================
'
'  Accounts.bas - Programa básico para llevar las cuentas de gastos e ingresos
'  de la casa o un pequeño negocio.
'
'  Copyright (C) 2022 Cuban Solutions
'
'  Este programa es gratuito y está publicado en https://cubansolutions.blogspot.com
'  para su descarga. 
'
'=====================================================================================

$APPTYPE GUI
$TYPECHECK ON

$RESOURCE APP_ICON AS "icon/icon.ico"
$RESOURCE APP_LOGO AS "logo/logo.bmp"
$RESOURCE NAVTOOLS_FIRST AS "images/First.bmp"
$RESOURCE NAVTOOLS_PRIOR AS "images/Prior.bmp"
$RESOURCE NAVTOOLS_NEXT AS "images/Next.bmp"
$RESOURCE NAVTOOLS_LAST AS "images/Last.bmp"
$RESOURCE NAVTOOLS_ADD AS "images/Add.bmp"
$RESOURCE NAVTOOLS_DELETE AS "images/Delete.bmp"
$RESOURCE NAVTOOLS_EDIT AS "images/Edit.bmp"
$RESOURCE NAVTOOLS_SAVE AS "images/Save.bmp"
$RESOURCE NAVTOOLS_CANCEL AS "images/Cancel.bmp"
$RESOURCE NAVTOOLS_FIND AS "images/Find.bmp"
$RESOURCE NAVTOOLS_REFRESH AS "images/Refresh.bmp"

$INCLUDE "RAPIDQ.INC"

DECLARE FUNCTION FindWindow LIB "USER32" ALIAS "FindWindowA" (className AS STRING, windowName AS STRING) AS INTEGER
DECLARE SUB SetForegroundWindow LIB "USER32" ALIAS "SetForegroundWindow" (HWnd AS LONG)
DECLARE SUB ShowWindow LIB "USER32" ALIAS "ShowWindow" (HWnd AS LONG, nCmd AS LONG)

DEFINT hWnd = FindWindow("TForm", "Accounts [Cuentas]")

If hWnd Then
    MessageDlg("El programa está corriendo como una instancia previa y no puede cargarse dos veces.", mtInformation, mbOk, 0)
    SetForegroundWindow(hWnd)
    ShowWindow(hWnd, 3)
    End
End If

$INCLUDE "QDBFTable.bas"

Application.IcoHandle = APP_ICON

Type MovmentsDataType
	Id As String*6
	Date As String*8
	ConceptDesc As String*100
	Detail As String*100
	Kind As String*2
	Amount As String*10
End Type

Dim MainDBFTable As QDBFTable
Dim MainDBFErr As QDBFErrorDesc
Dim AuxDBFTable As QDBFTable

$INCLUDE "Concepts.bas"
$INCLUDE "Revenues.bas"
$INCLUDE "Expenses.bas"
$INCLUDE "Balances.bas"
$INCLUDE "ConceptsBrief.bas"
$INCLUDE "DetailedBrief.bas"
$INCLUDE "About.bas"

DECLARE SUB AppExitItemClick(Sender AS QMENUITEM)
DECLARE SUB AppConceptsItemClick(Sender AS QMENUITEM)
DECLARE SUB AppRevenuesItemClick(Sender AS QMENUITEM)
DECLARE SUB AppExpensesItemClick(Sender AS QMENUITEM)
DECLARE SUB AppBalancesItemClick(Sender AS QMENUITEM)
DECLARE SUB AppConceptsBriefItemClick(Sender AS QMENUITEM)
DECLARE SUB AppDetailedBriefItemClick(Sender AS QMENUITEM)
DECLARE SUB AppAboutItemClick(Sender AS QMENUITEM)

CREATE AppAccountsMainForm AS QForm
	Center
	Caption = "Accounts [Cuentas]"
	WindowState = wsMaximized
	CREATE AppMainMenu AS QMainMenu
		CREATE AppSystemMenu AS QMenuItem
			Caption = "&Sistema"
			CREATE AppExitItem AS QMenuItem
				Caption = "&Salir"
				OnClick = AppExitItemClick
			END CREATE
		END CREATE
		CREATE AppDefinitionsMenu AS QMenuItem
			Caption = "&Definiciones"
			CREATE AppConceptsItem AS QMenuItem
				Caption = "&Conceptos..."
				OnClick = AppConceptsItemClick
			END CREATE
		END CREATE
		CREATE AppMovementsMenu AS QMenuItem
			Caption = "&Movimientos"
			CREATE AppRevenuesItem AS QMenuItem
				Caption = "&Ingresos..."
				OnClick = AppRevenuesItemClick
			END CREATE
			CREATE AppExpensesItem AS QMenuItem
				Caption = "&Gastos..."
				OnClick = AppExpensesItemClick
			END CREATE
			CREATE AppMovementsSep1 AS QMenuItem
				Caption = "-"
			END CREATE
			CREATE AppBalanceItem AS QMenuItem
				Caption = "&Saldos..."
				OnClick = AppBalancesItemClick
			END CREATE
			CREATE AppMovementsSep2 AS QMenuItem
				Caption = "-"
			END CREATE
			CREATE AppBriefItem AS QMenuItem
				Caption = "&Resumen"
				CREATE AppConceptsBriefItem AS QMenuItem
					Caption = "Resumen por &conceptos..."
					OnClick = AppConceptsBriefItemClick
				END CREATE
				CREATE AppDetailedBriefItem AS QMenuItem
					Caption = "Resumen &detallado..."
					OnClick = AppDetailedBriefItemClick
				END CREATE
			END CREATE
		END CREATE
		CREATE AppAboutMenu AS QMenuItem
			Caption = "&Acerca de"
			CREATE AppAboutItem AS QMenuItem
				Caption = "&Acerca de Accounts..."
				OnClick = AppAboutItemClick
			END CREATE
		END CREATE
	END CREATE
END CREATE

AppAccountsMainForm.ShowModal

'--------- Subroutines ---------

SUB AppExitItemClick(Sender AS QMENUITEM)
    Application.Terminate
END SUB

SUB AppConceptsItemClick(Sender AS QMENUITEM)
    If MainDBFTable.Connect(CurDir$ + "\data\" + "CONCEPTS.DBF") = False Then
        MessageDlg("La tabla 'CONCEPTS.DBF' no es válida o no se encuentra en la carpeta 'data' del programa.", mtError, mbOk, 0)
        Exit Sub
    End If
    
    If MainDBFTable.First Then
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
        
        ConceptsDescriptionEdit.Text = MainDBFTable.FieldByName("DESC")
		Select Case Val(MainDBFTable.FieldByName("KIND"))
			Case 1
				ConceptsRevenuesRadioButton.Checked = True
			Case 2
				ConceptsExpensesRadioButton.Checked = True
		End Select
    Else
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
    End If
    
    ConceptsMainForm.Parent = AppAccountsMainForm 
    ConceptsMainForm.ShowModal

END SUB

SUB AppRevenuesItemClick(Sender AS QMENUITEM)
	If MainDBFTable.Connect(CurDir$ + "\data\" + "MOVMENTS.DBF") = False Then
        MessageDlg("La tabla 'MOVMENTS.DBF' no es válida o no se encuentra en la carpeta 'data' del programa.", mtInformation, mbOk, 0)
        Exit Sub
    End If
	
	RevenuesConceptComboBox.Clear
	
	With AuxDBFTable 
		If .Connect(CurDir$ + "\data\" + "CONCEPTS.DBF") = False Then
			MessageDlg("La tabla 'CONCEPTS.DBF' no es válida o no se encuentra en la carpeta 'data' del programa.", mtInformation, mbOk, 0)
			Exit Sub
		End If
		While .Fetch
			If Val(.FieldByName("KIND")) = 1 Then
				RevenuesConceptComboBox.AddItems(.FieldByName("DESC"))	
			End If
		Wend
		.Close
	End With
	
	Dim CurrentDate As String
	
	CurrentDate = Date$
	
	RevenuesDayEdit.Text = Mid$(CurrentDate, 4, 2)
	RevenuesMonthEdit.Text = Left$(CurrentDate, 2)
	RevenuesYearEdit.Text = Right$(CurrentDate, 4)
	
	RevenuesDetailEdit.Text = ""
	RevenuesAmountEdit.Text = ""
	
	RevenuesMainForm.Parent = AppAccountsMainForm 
    RevenuesMainForm.ShowModal
	
END SUB

SUB AppExpensesItemClick(Sender AS QMENUITEM)
	If MainDBFTable.Connect(CurDir$ + "\data\" + "MOVMENTS.DBF") = False Then
        MessageDlg("La tabla 'MOVMENTS.DBF' no es válida o no se encuentra en la carpeta 'data' del programa.", mtInformation, mbOk, 0)
        Exit Sub
    End If
	
	ExpensesConceptComboBox.Clear
	
	With AuxDBFTable 
		If .Connect(CurDir$ + "\data\" + "CONCEPTS.DBF") = False Then
			MessageDlg("La tabla 'CONCEPTS.DBF' no es válida o no se encuentra en la carpeta 'data' del programa.", mtInformation, mbOk, 0)
			Exit Sub
		End If
		While .Fetch
			If Val(.FieldByName("KIND")) = 2 Then
				ExpensesConceptComboBox.AddItems(.FieldByName("DESC"))
			End If
		Wend
		.Close
	End With
	
	Dim CurrentDate As String
	
	CurrentDate = Date$
	
	ExpensesDayEdit.Text = Mid$(CurrentDate, 4, 2)
	ExpensesMonthEdit.Text = Left$(CurrentDate, 2)
	ExpensesYearEdit.Text = Right$(CurrentDate, 4)
	
	ExpensesDetailEdit.Text = ""
	ExpensesAmountEdit.Text = ""
	
	ExpensesMainForm.Parent = AppAccountsMainForm 
    ExpensesMainForm.ShowModal
END SUB

SUB AppBalancesItemClick(Sender AS QMENUITEM)
	If MainDBFTable.Connect(CurDir$ + "\data\" + "BALANCES.DBF") = False Then
        MessageDlg("La tabla 'BALANCES.DBF' no es válida o no se encuentra en la carpeta 'data' del programa.", mtError, mbOk, 0)
        Exit Sub
    End If
    
    If MainDBFTable.First Then
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
        
        BalancesYearEdit.Text = MainDBFTable.FieldByName("YEAR")
		BalancesAmountEdit.Text = MainDBFTable.FieldByName("AMOUNT")
    Else
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
    
    BalancesMainForm.Parent = AppAccountsMainForm 
    BalancesMainForm.ShowModal
END SUB

SUB AppConceptsBriefItemClick(Sender AS QMENUITEM)
	If MainDBFTable.Connect(CurDir$ + "\data\" + "CONCEPTS.DBF") = False Then
        MessageDlg("La tabla 'CONCEPTS.DBF' no es válida o no se encuentra en la carpeta 'data' del programa.", mtError, mbOk, 0)
        Exit Sub
    End If
	
	Dim CurrentDate As String
	
	CurrentDate = Date$
	
	ConceptsBriefParametersDayEdit.Text = Mid$(CurrentDate, 4, 2)
	ConceptsBriefParametersMonthEdit.Text = Left$(CurrentDate, 2)
	ConceptsBriefParametersYearEdit.Text = Right$(CurrentDate, 4)
	
	ConceptsBriefParametersPeriodComboBox.ItemIndex = 0
	
	ConceptsBriefDetailStringGrid.RowCount = 1
	
	ConceptsBriefSubtotalsTotalsRevenuesRichEdit.Text = "0.00"
	ConceptsBriefSubtotalsTotalsExpensesRichEdit.Text = "0.00"
	
	ConceptsBriefMainForm.Parent = AppAccountsMainForm 
    ConceptsBriefMainForm.ShowModal
END SUB

SUB AppDetailedBriefItemClick(Sender AS QMENUITEM)
	If MainDBFTable.Connect(CurDir$ + "\data\" + "MOVMENTS.DBF") = False Then
        MessageDlg("La tabla 'MOVMENTS.DBF' no es válida o no se encuentra en la carpeta 'data' del programa.", mtError, mbOk, 0)
        Exit Sub
    End If
	
	Dim CurrentDate As String
	
	CurrentDate = Date$
	
	DetailedBriefParametersDayEdit.Text = Mid$(CurrentDate, 4, 2)
	DetailedBriefParametersMonthEdit.Text = Left$(CurrentDate, 2)
	DetailedBriefParametersYearEdit.Text = Right$(CurrentDate, 4)
	
	DetailedBriefParametersPeriodComboBox.ItemIndex = 0
	
	DetailedBriefDetailStringGrid.RowCount = 1
	
	DetailedBriefHistoricTotalsBalanceRichEdit.Text = "0.00"
	DetailedBriefHistoricTotalsBalanceRichEdit.Font.Color = QBColor(2)
	DetailedBriefSubtotalsTotalsRevenuesRichEdit.Text = "0.00"
	DetailedBriefSubtotalsTotalsExpensesRichEdit.Text = "0.00"
	DetailedBriefSubtotalsTotalsBalanceRichEdit.Text = "0.00"
	DetailedBriefSubtotalsTotalsBalanceRichEdit.Font.Color = QBColor(2)
	
	DetailedBriefMainForm.Parent = AppAccountsMainForm 
    DetailedBriefMainForm.ShowModal
END SUB

SUB AppAboutItemClick(Sender AS QMENUITEM)
	AboutMainForm.Parent = AppAccountsMainForm
	AboutMainForm.ShowModal
END SUB


'====================================================================================
'
'  About.bas - Formulario Acerca de Account.
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

DECLARE SUB AboutOkButtonClick (Sender AS QBUTTON)

CREATE AboutMainForm AS QFORM
    Caption = "Acerca de Accounts"
    Width = 344
    Height = 264
	BorderStyle = bsDialog
    Center
    CREATE AboutGroupBox AS QGROUPBOX
        Left = 9
        Top = 8
        Width = 321
        Height = 188
        CREATE AboutLogoImage AS QIMAGE
            BMPHandle = APP_LOGO
            Left = 10
            Top = 16
            Width = 73
            Height = 73
        END CREATE
        CREATE AboutProgramNameLabel AS QLABEL
            Caption = "Accounts"
            Left = 102
            Top = 18
            Width = 208
			Transparent = 1
			Font.AddStyles(fsBold)
			Font.Size = 16
        END CREATE
        CREATE AboutVersionLabel AS QLABEL
            Caption = "Versión 1.0.0"
            Left = 102
            Top = 40
			Transparent = 1
			Font.AddStyles(fsBold)
			Font.Size = 12
        END CREATE
        CREATE AboutCreatedLabel AS QLABEL
            Caption = "Creado en 2022 por:"
            Left = 102
            Top = 66
			Transparent = 1
        END CREATE
        CREATE AboutCreatorNameLabel AS QLABEL
            Caption = "Pedro Luis Carballosa Mass"
            Left = 102
            Top = 80
			Transparent = 1
        END CREATE
        CREATE AboutCreatorEMAILLabel AS QLABEL
            Caption = "pedro.l.carballosa@gmail.com"
            Left = 102
            Top = 94
			Transparent = 1
        END CREATE
		CREATE AboutCreatorWebSiteLabel AS QLABEL
            Caption = "https://cubansolutions.blogspot.com"
            Left = 102
            Top = 108
			Transparent = 1
        END CREATE
		CREATE AboutSeparatorGroupBox AS QGROUPBOX
            Left = 5
            Top = 128
            Width = 310
            Height = 4
            Color = 0
        END CREATE
        CREATE AboutCommentLabel AS QLABEL
            Caption = "Comentario:"
            Left = 7
            Top = 134
			Transparent = 1
        END CREATE
        CREATE AboutCommentTextLabel AS QLABEL
            Caption = "El presente programa para llevar los gastos e ingresos es gratuito para todo uso."
            Left = 7
            Top = 148
			Height = 20
			Width = 305
            Transparent = 1
            Wordwrap = 1
        END CREATE
    END CREATE
    CREATE AboutOkButton AS QBUTTON
        Caption = "&Aceptar"
        Left = 137
        Top = 204
		Width = 70
        TabOrder = 0
		OnClick = AboutOkButtonClick
    END CREATE
END CREATE

'--------- Subroutines ---------

SUB AboutOkButtonClick (Sender AS QBUTTON)
	AboutMainForm.Close
END SUB
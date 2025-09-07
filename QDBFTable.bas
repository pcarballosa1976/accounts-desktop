'==================================================================================
'
'  QDBFTable.bas - Clase básica para manipular una tabla (".dbf") de Foxbase.
'
'  Copyright (C) 2017 Cuban Solutions
'
'  Este código de demostración fue elaborado como parte de un tutorial sobre cómo usar una tabla ".dbf"
'  desde RapidQ por medios propios (sin usar otras librerías).
'
'  El tutorial original, así como software gratis, además de info adicional sobre matemáticas y programación
'  en diferentes lenguajes de cómputo y artículos de temática variada, lo puede encontrar en la dirección de 
'  mi blog personal https://cubansolutions.blogspot.com
'
'==================================================================================

Const QDBF_MAX_FIELD As Byte = 50

Const QDBF_NORMAL_MODE As Word = 0
Const QDBF_EDIT_MODE As Word = 1
Const QDBF_APPEND_MODE As Word = 2

Const QDBF_ERR_NONE As Byte = 0
Const QDBF_ERR_FILE_NOT_EXIST As Byte = 1
Const QDBF_ERR_FILE_NOT_SUPPORTED As Byte = 2
Const QDBF_ERR_FILE_ENCRYPTED As Byte = 3
Const QDBF_ERR_TOO_MANY_FIELD As Byte = 4
Const QDBF_ERR_FILE_PROBABLY_CORRUPTED As Byte = 5
Const QDBF_ERR_FIELD_INDEX_OUT_BOUNDS As Byte = 6
Const QDBF_ERR_RECORD_INDEX_OUT_BOUNDS As Byte = 7
Const QDBF_ERR_INVALID_OPERATION As Byte = 8

Type QDBFErrorDesc
    Id As Byte
    Description As String*50
End Type

Type QDBFFieldDesc
    Name As String*11
    TypeId As String*1
    Displacement As DWord
    Length As Byte
    DecimalPlaces As Byte
    Reserved1 As Word
    WorkAreaId As Byte
    Reserved2(1 TO 10) As Byte
    HaveIndex As Byte
End Type

Type QDBFTable Extends QObject
              
    PRIVATE:
          
        _Id As Byte
        _Year As Byte
        _Month As Byte
        _Day As Byte
        _RecordCount As DWord
        _HeaderBytes As Word
        _RecordBytes As Word
        _Reserved1 As Word
        _TransactionFlag As Byte
        _EncryptionFlag As Byte
        _MultiUserEnv(1 TO 12) As Byte
        _ProductionIndexExist As Byte
        _DB4LanguageDriverId As Byte
        _Reserved2 As Word
              
        _BOF As Integer
        _EOF As Integer
                
        _FieldCount As Word
        _FieldDesc(1 To QDBF_MAX_FIELD) As QDBFFieldDesc
        
        _Connected As Integer
        _Mode As Word
           
        _CurrentPath As String
        _CurrentField As Word
        _CurrentRecord As String
        _CurrentPosition As DWord
        _OldCurrentPosition As DWord
        _CurrentError As QDBFErrorDesc
        
        Function _OpenFile(File As QFileStream, Path As String, Mode As Word) As Integer
            With This
                ._OpenFile = False
                If File.Open(Path, Mode) = False Then
                    ._CurrentError.Id = QDBF_ERR_FILE_NOT_EXIST
                    ._CurrentError.Description = "File does not exist or could not be open"
                    Exit Function
                End If
                ._OpenFile = True
            End With
        End Function
        
        Function _IsConnected() As Integer
            With This
                ._IsConnected = False
                If ._Connected = False Then
                    ._CurrentError.Id = QDBF_ERR_INVALID_OPERATION
                    ._CurrentError.Description = "Invalid operation with close database file"
                    Exit Function
                End If
                ._IsConnected = True
            End With
        End Function
        
        Function _IsValidFieldIndex(Index As Word) As Integer
            With This
                ._IsValidFieldIndex = False
                If ._Connected = False Then
                    Exit Function
                ElseIf Index < 1 Or Index > ._FieldCount Then
                    ._CurrentError.Id = QDBF_ERR_FIELD_INDEX_OUT_BOUNDS
                    ._CurrentError.Description = "Field index out of bounds"
                    Exit Function
                End If
                ._IsValidFieldIndex = True
            End With
        End Function
        
        Function _IsValidRecordIndex(Index As DWord) As Integer
            With This
                ._IsValidRecordIndex = False
                If ._Connected = False Then
                    Exit Function
                ElseIf Index < 1 Or Index > ._RecordCount Then
                    ._CurrentError.Id = QDBF_ERR_RECORD_INDEX_OUT_BOUNDS
                    ._CurrentError.Description = "Record index out of bounds"
                    Exit Function
                End If
                ._IsValidRecordIndex = True
            End With
        End Function
        
        Sub _NormalMode()
            With This
                If ._Mode = QDBF_APPEND_MODE Then
                    ._RecordCount -= 1
                    ._CurrentPosition = ._RecordCount
                End If
                ._Mode = QDBF_NORMAL_MODE
            End With
        End Sub
		
		Function _CopyString(Str1 As String) As String
			Dim i As Word
			With This
				Str1 = RTrim$(LTrim$(Str1))
				i = Len(Str1)
				If i > 0 Then
					While Asc(Str1[i]) = 0
						i = i - 1
					Wend
					._CopyString = Mid$(Str1, 1, i)
				Else
					._CopyString = ""
				End If
			End With
		End Function

        Function _CompareString(Str1 As String, Str2 As String, StrictCompare As Integer) As Integer
            Dim Length As Integer
			Dim i As Word
            With This
				Str1 = ._CopyString(Str1)
				Str2 = ._CopyString(Str2)
				Length = IIF(Len(Str1) < Len(Str2), Len(Str1), Len(Str2))
				If StrictCompare = True Then
					._CompareString = True
					If (Len(Str1)) <> (Len(Str2)) Then
						._CompareString = False
					Else
						If Length > 1 Then
							For i = 1 To Length
								If Str1[i] <> Str2[i] Then
									._CompareString = False
									Exit For
								End If
							Next
						Else
							If Length = 1 Then
								If Str1[1] <> Str2[1] Then
									._CompareString = False
								End If
							End If
						End If
					End If
				Else
					._CompareString = True
					If Length > 1 Then
						For i = 1 To Length
							If Str1[i] <> Str2[i] Then
								._CompareString = False
								Exit For
							End If
						Next
					Else
						If Length = 1 Then
							If Str1[1] <> Str2[1] Then
								._CompareString = False
							End If
						End If
					End If
				End If
            End With
        End Function
		
		Function _DateValid (Year As Integer, Month As Integer, Day As Integer) As Integer
		    With This
				Select Case Month
					Case 4, 6, 9, 11                  
						Select Case Day
							Case 1 TO 30
							Case Else
								._DateValid = False
								Exit Function
						End Select
					Case 1, 3, 5, 7, 8, 10, 12
						Select Case Day
							Case 1 TO 31
							Case Else
								._DateValid = False
								Exit Function
						End Select
					Case 2                            
						Select Case Year MOD 4      
							Case 0  
								Select Case Day
									Case 1 TO 29
									Case Else
										._DateValid = False
										Exit Function
								End Select
							Case Else
								Select Case Day
									Case 1 TO 28 
									Case Else
										._DateValid = False
										Exit Function
									End Select
						End Select                  
					Case Else
						._DateValid = False
						Exit Function
				End Select
				._DateValid = True
			End With
		End Function
                   
    PUBLIC:
            
        Function Path() As String
            With This
                .Path = ._CurrentPath
            End With
        End Function
            
        Function Connected() As Integer
            With This
                .Connected = ._Connected
            End With
        End Function
        
        Function Mode() As Word
            With This
                .Mode = ._Mode
            End With
        End Function
        
        Function Position() As DWord
            With This
                .Position = ._CurrentPosition
            End With
        End Function
            
        Function FieldCount() As Word
            With This
                .FieldCount = ._FieldCount
            End With
        End Function
        
        Function BOF() As Word
            With This
                .BOF = ._BOF
            End With
        End Function
        
        Function EOF() As Word
            With This
                .EOF = ._EOF
            End With
        End Function
        
        Sub ClearError()
            With This
                ._CurrentError.Id = QDBF_ERR_NONE
                ._CurrentError.Description = ""
            End With
        End Sub
            
        Sub LastError(ByRef ErrorDesc As QDBFErrorDesc)
            With This
                ErrorDesc.Id = ._CurrentError.Id
                ErrorDesc.Description = ._CurrentError.Description
            End With
        End Sub
        
        Function FieldName(Index As Word) As String
            With This
                .FieldName = ""
                If ._IsConnected = False Then
                    Exit Function
                ElseIf ._IsValidFieldIndex(Index) = False Then
                    Exit Function
                End If
                .FieldName = ._CopyString(._FieldDesc(Index).Name)
            End With
        End Function
        
        Function FieldType(Index As Word) As String
            With This
                .FieldType = ""
                If ._IsConnected = False Then
                    Exit Function
                ElseIf ._IsValidFieldIndex(Index) = False Then
                    Exit Function
                End If
                .FieldType = ._FieldDesc(Index).TypeId
            End With
        End Function
        
        Function FieldLength(Index As Word) As Word
            With This
                .FieldLength = 0
                If ._IsConnected = False Then
                    Exit Function
                ElseIf ._IsValidFieldIndex(Index) = False Then
                    Exit Function
                End If
                .FieldLength = ._FieldDesc(Index).Length
            End With
        End Function

        Function FieldDecimals(Index As Word) As Word
            With This
                .FieldDecimals = 0
                If ._IsConnected = False Then
                    Exit Function
                ElseIf ._IsValidFieldIndex(Index) = False Then
                    Exit Function
                End If
                .FieldDecimals = ._FieldDesc(Index).DecimalPlaces
            End With
        End Function

        Function FieldIndex(Name As String) As Word
            Dim Found As Integer, i As Word
            Found = False
            With This
                .FieldIndex = 0
                If ._IsConnected = False Then
                    Exit Function
                End If
                For i = 1 To ._FieldCount
                    If ._CompareString(UCase$(._FieldDesc(i).Name), UCase$(Name), True) = True Then
						Found = True
                        Exit For
                    End If
				Next i
                If Found Then
                    .FieldIndex = i
                End If
            End With
        End Function
            
        Function FieldByIndex(Index As Word) As String
            With This
                .FieldByIndex = ""
                If ._IsConnected = False Then
                    Exit Function
                ElseIf ._IsValidFieldIndex(Index) = False Then
                    Exit Function
                ElseIf ._IsValidRecordIndex(._CurrentPosition) = False Then
                    Exit Function
                End If
                .FieldByIndex = RTrim$(LTrim$(Mid$(._CurrentRecord, ._FieldDesc(Index).Displacement + 1, ._FieldDesc(Index).Length)))
            End With
        End Function
        
        Function FieldByName(Name As String) As String
            Dim FieldIndex As Word
            With This
                .FieldByName = ""
                FieldIndex = .FieldIndex(Name)
                If FieldIndex > 0 Then
                    .FieldByName = .FieldByIndex(FieldIndex)
                End If
            End With 
        End Function
		
		Function RecordNo() As DWord
            With This
                .RecordNo = ._CurrentPosition
            End With
        End Function
            
        Function RecordCount() As DWord
            With This
                .RecordCount = ._RecordCount
            End With
        End Function
                
        Function LastUpdate() As String
            With This
                .LastUpdate = ""
                If ._IsConnected = False Then Exit Function
                .LastUpdate = IIF(._Month < 9, "0" + Str$(._Month), Str$(._Month)) + "/" + IIF(._Day < 9, "0" + Str$(._Day), Str$(._Day)) + "/" + Str$(._Year)
            End With
        End Function
            
        Function Go(Index As DWord) As Integer
            Dim File As QFileStream
            With This
                .Go = False
                '._NormalMode Deshabilitado para obligar a usar el método Cancel
                ._CurrentPosition = Index
                ._BOF = False
                ._EOF = False
                If ._IsConnected = False Then
                    Exit Function
                ElseIf ((._Mode = QDBF_EDIT_MODE) Or (._Mode = QDBF_APPEND_MODE)) Then
                    ._CurrentError.Id = QDBF_ERR_INVALID_OPERATION
                    ._CurrentError.Description = "Database is in edit or append mode"
                    Exit Function    
                ElseIf ._RecordCount < 1 Then
                    ._BOF = True
                    ._EOF = True
                    ._CurrentPosition = 0
                    Exit Function
                ElseIf ._OpenFile(File, ._CurrentPath, fmOpenRead) = False Then
                    ._BOF = True
                    ._EOF = True
                    Exit Function
                ElseIf ._IsValidRecordIndex(Index) = False Then
                    If Index < 1 Then
                        ._BOF = True 
                        ._CurrentPosition = 1
                    ElseIf Index > ._RecordCount Then
                        ._EOF = True
                        ._CurrentPosition = ._RecordCount
                    End If
                End If
                File.Position = (._CurrentPosition - 1) * ._RecordBytes + ._HeaderBytes
                ._CurrentRecord = File.ReadStr(._RecordBytes)
                File.Close
                .Go = True
            End With
        End Function
        
        Function GoTop() As Integer
            With This
                .GoTop = .Go(1)
            End With
        End Function
        
        Function Skip(Offset As Integer) As Integer
            With This
                .Skip = .Go(._CurrentPosition + Offset)
            End With 
        End Function
                
        Function GoBottom() As Integer
            With This
                .GoBottom = .Go(._RecordCount)
            End With
        End Function
        
        Function Connect(Path As String) As Integer
            Dim File As QFileStream
            Dim FieldDesc As QDBFFieldDesc
            Dim i As Word
            Dim Aux As Byte
            
			With This
                .Connect = False
                If ._Connected Then Exit Function
                If ._OpenFile(File, Path, fmOpenRead) = False Then Exit Function
                File.Read(._Id)
                File.Read(._Year)
                File.Read(._Month)
                File.Read(._Day)
                File.Read(._RecordCount)
                File.Read(._HeaderBytes)
                File.Read(._RecordBytes)
                File.Read(._Reserved1)
                File.Read(._TransactionFlag)
                File.Read(._EncryptionFlag)
                For i = 1 To 12
                    File.Read(Aux)
                    ._MultiUserEnv(i) = Aux
                Next
                File.Read(._ProductionIndexExist)
                File.Read(._DB4LanguageDriverId)
                File.Read(._Reserved2)
                If ._Id <> 3 Then
                    ._CurrentError.Id = QDBF_ERR_FILE_NOT_SUPPORTED
                    ._CurrentError.Description = "File not supported"
                    File.Close
                    Exit Function
                End If
                If ._EncryptionFlag <> 0 Then
                    ._CurrentError.Id = QDBF_ERR_FILE_ENCRYPTED
                    ._CurrentError.Description = "The database file is encrypted"
                    File.Close
                    Exit Function
                End If
                ._CurrentPath = Path
				._FieldCount = (._HeaderBytes - 33)/32
                If ._FieldCount > QDBF_MAX_FIELD Then
                    ._CurrentError.Id = QDBF_ERR_TOO_MANY_FIELD
                    ._CurrentError.Description = "Too many fields in dbf file"
                    File.Close
                    Exit Function
                End If
                For i = 1 To ._FieldCount
                    File.ReadUDT(FieldDesc)
					._FieldDesc(i) = FieldDesc
				Next
                File.Read(Aux)
                If Aux <> 13 Then
                    ._CurrentError.Id = QDBF_ERR_FILE_PROBABLY_CORRUPTED
                    ._CurrentError.Description = "Database file probably corrupted"
                    File.Close
                    Exit Function
                End If
                File.Close
                ._Connected = True
                .GoTop
                .Connect = True
            End With
        End Function
        
        Sub Close()
            With This
                ._NormalMode
                ._CurrentRecord = ""
                ._Connected = False
            End With
        End Sub
        
        Function Deleted() As Integer
            With This
                .Deleted = False
                If ._IsConnected = False Then
                    Exit Function
                ElseIf RTrim$(LTrim$(._CurrentRecord)) = "" Then
                    Exit Function
                End If
                .Deleted = IIF(Mid$(._CurrentRecord, 1, 1) = "*", True, False)
            End With
        End Function
		
		Function First() As Integer
			With This
				.First = False
                If .BOF = True And .EOF = True Then Exit Function
                .First = .GoTop
				While .Deleted = True
					.First = .Skip(1)
					If .EOF = True Then
						.ClearError()
						.First = False
						Exit Function
					End If
				Wend
            End With
		End Function

        Function Prior() As Integer
            With This
				.Prior = False
                If .BOF = True And .EOF = True Then Exit Function
                .Prior = .Skip(-1)
				While .Deleted = True And .BOF = False
					.Prior = .Skip(-1)
				Wend
                If .Deleted = True And .BOF = True Then
                    While .Deleted = True
						.Prior = .Skip(1)
						If .EOF = True Then
							.ClearError()
							.Prior = False
							Exit Function
						End If
					Wend
                End If
            End With
        End Function
        
        Function Next() As Integer
            With This
				.Next = False
                If .BOF = True And .EOF = True Then Exit Function
                .Next = .Skip(1)
				While .Deleted = True And .EOF = False
					.Next = .Skip(1)
				Wend
                If .Deleted = True And .EOF = True Then
                    While .Deleted = True
						.Next = .Skip(-1)
						If .BOF = True Then
							.ClearError()
							.Next = False
							Exit Function
						End If
					Wend
                End If
            End With
        End Function
		
		Function Last() As Integer
			With This
				.Last = False
                If .BOF = True And .EOF = True Then Exit Function
                .Last = .GoBottom
				While .Deleted = True
					.Last = .Skip(-1)
					If .BOF = True Then
						.ClearError()
						.Last = False
						Exit Function
					End If
				Wend
            End With
		End Function
		
        Function Fetch() As Integer
            STATIC Index As DWord
            With This
                .Fetch = False
                If Index > ._CurrentPosition Then Index = 0
                Do
                    Index += 1
                    If .Go(Index) = False Then
                        Exit Function
                    ElseIf .Deleted = False Then
                        Exit Do
                    Elseif ._EOF = True Then
                        Exit Do    
                    End If
					DoEvents
                Loop
                .Fetch = (._EOF <> True)
            End With
        End Function
            
        Function Delete() As Integer
            Dim File As QFileStream
            Dim Err As QDBFErrorDesc
            With This
                .Delete = False
                .ClearError
                If ._IsConnected = False Then
                    Exit Function
                ElseIf ._IsValidRecordIndex(._CurrentPosition) = False Then
                    Exit Function
                ElseIf .Deleted = True Then
                    .LastError(Err)
                    If Err.Id <> QDBF_ERR_NONE Then
                        ._CurrentError.Id = Err.Id
                        ._CurrentError.Description = Err.Description
                    Else
                        ._CurrentError.Id = QDBF_ERR_INVALID_OPERATION
                        ._CurrentError.Description = "Invalid operation with deleted data record"
                    End If
                    Exit Function
                ElseIf ._OpenFile(File, ._CurrentPath, fmOpenWrite) = False Then
                    Exit Function
                End If
                File.Position = (._CurrentPosition - 1) * ._RecordBytes + ._HeaderBytes
                ._CurrentRecord = "*" + Mid$(._CurrentRecord, 2, Len(._CurrentRecord)-1)
                File.WriteStr(._CurrentRecord, ._RecordBytes)
                File.Close
                .Delete = True
            End With
        End Function
		
		Function Locate(FieldName As String, Value As String, PartialKey As Integer) As Integer
            Dim FieldValue As String
            Dim FieldIndex As Word
            With This
				.Locate = False
				If ._IsConnected = False Then
                    Exit Function
                ElseIf ._IsValidRecordIndex(._CurrentPosition) = False Then
                    Exit Function
                End If
                FieldIndex = .FieldIndex(FieldName)
                If FieldIndex = 0 Then
                    ._CurrentError.Id = QDBF_ERR_INVALID_OPERATION
                    ._CurrentError.Description = "The especified field do not exist in the database"
                    Exit Function
                End If
				If ._RecordCount = 0 Then
					._CurrentError.Id = QDBF_ERR_INVALID_OPERATION
                    ._CurrentError.Description = "The recordset can not be empty"
                    Exit Function
				End If
				If ._Mode <> QDBF_NORMAL_MODE Then
					._CurrentError.Id = QDBF_ERR_INVALID_OPERATION
                    ._CurrentError.Description = "Database is not in normal mode"
                    Exit Function
				End If
				While ._EOF <> True
                    If .Deleted = False Then
                        FieldValue = .FieldByIndex(FieldIndex)
                        If ._CompareString(UCase$(FieldValue), UCase$(Value), IIF(PartialKey = True, False, True)) = True Then
                            .Locate = True
                            Exit Function
                        End If
                    End If
					.Skip(1)
					DoEvents
                Wend
				.Locate = False
            End With
        End Function
            
        Function Recall() As Integer
            Dim File As QFileStream
            With This
                .Recall = False
                If ._IsConnected = False Then
                    Exit Function
                ElseIf ._IsValidRecordIndex(._CurrentPosition) = False Then
                    Exit Function
                ElseIf .Deleted = False Then
                    ._CurrentError.Id = QDBF_ERR_INVALID_OPERATION
                    ._CurrentError.Description = "Invalid operation with undeleted data record"
                    Exit Function
                ElseIf ._OpenFile(File, ._CurrentPath, fmOpenWrite) = False Then
                    Exit Function
                End If
                File.Position = (._CurrentPosition - 1) * ._RecordBytes + ._HeaderBytes
                ._CurrentRecord = " " + Mid$(._CurrentRecord, 2, Len(._CurrentRecord)-1)
                File.WriteStr(._CurrentRecord, ._RecordBytes)
                File.Close
                .Recall = True
            End With
        End Function
        
        Function Edit() As Integer
            With This
                .Edit = False
                If ._IsConnected = False Then
                    Exit Function
                ElseIf ((._Mode = QDBF_EDIT_MODE) Or (._Mode = QDBF_APPEND_MODE)) Then
                    ._CurrentError.Id = QDBF_ERR_INVALID_OPERATION
                    ._CurrentError.Description = "Database already is in append or edit mode"
                    Exit Function
                ElseIf ._IsValidRecordIndex(._CurrentPosition) = False Then
                    Exit Function
                End If
                ._Mode = QDBF_EDIT_MODE
                .Edit = True
            End With
        End Function
            
        Function Append() As Integer
            With This
                .Append = False
                If ._IsConnected = False Then
                    Exit Function
                ElseIf ((._Mode = QDBF_EDIT_MODE) Or (._Mode = QDBF_APPEND_MODE)) Then
                    ._CurrentError.Id = QDBF_ERR_INVALID_OPERATION
                    ._CurrentError.Description = "Database already is in append or edit mode"
                    Exit Function
                End If
                ._RecordCount += 1
                ._OldCurrentPosition = ._CurrentPosition
                ._CurrentPosition = ._RecordCount
                ._CurrentRecord = Space$(._RecordBytes)
                ._Mode = QDBF_APPEND_MODE
                .Append = True
            End With
        End Function
            
        Function SetField(Index As Word, Value As String) As Integer
            Dim FieldValue As String
            Dim Fracc As String
            With This
                .SetField = False
                If ._IsConnected = False Then
                    Exit Function
                ElseIf (._Mode = QDBF_NORMAL_MODE) Then
                    ._CurrentError.Id = QDBF_ERR_INVALID_OPERATION
                    ._CurrentError.Description = "Database is not in edit or append mode"
                    Exit Function
                ElseIf ((._Mode <> QDBF_APPEND_MODE) And (.Deleted = True)) Then
                    ._CurrentError.Id = QDBF_ERR_INVALID_OPERATION
                    ._CurrentError.Description = "Records deleted can not be modified"
                    Exit Function
                ElseIf ._IsValidFieldIndex(Index) = False Then
                    Exit Function
                ElseIf ._IsValidRecordIndex(._CurrentPosition) = False Then
                    Exit Function   
                ElseIf Len(Value) > ._FieldDesc(Index).Length Then
                    Value = Mid$(Value, 1, ._FieldDesc(Index).Length)
                End If
                
		    Select Case ._FieldDesc(Index).TypeId
                    Case "C"  
                        FieldValue = Value + Space$(._FieldDesc(Index).Length - Len(Value))
                    Case "D"
                        'Las fechas deben recibirse y se guardan como una cadena de texto con el formato Año Mes Día
						'sin espacios entre ellos.
						FieldValue = Value + Space$(._FieldDesc(Index).Length - Len(Value))
						If ._DateValid(Val(Mid$(Value, 1, 4)), Val(Mid$(Value, 5, 2)), Val(Mid$(Value, 7, 2))) = True Then
							FieldValue = Value
						Else
							FieldValue = "18990101"
						End If
                    Case "L"
                        If ((UCase$(Value) <> "T") And (UCase$(Value) <> "F" )) Then
                            FieldValue = "F"
                        Else
                            FieldValue = UCase$(Value)
                        End If
                    Case "N", "F"
                        Value = Str$(Val(Value))
                        If ._FieldDesc(Index).DecimalPlaces > 0 Then
                            If InStr(0, Value, ".") = False Then
                               Value = Value + "." + String$(._FieldDesc(Index).DecimalPlaces, "0")
                            Else
                               Fracc = Mid$(Value, InStr(0, Value, ".") + 1)
                               If Len(Fracc) < ._FieldDesc(Index).DecimalPlaces Then
                                   Fracc = Fracc + STRING$(._FieldDesc(Index).DecimalPlaces - Len(Fracc), "0") 
                               ElseIf Len(Fracc) > ._FieldDesc(Index).DecimalPlaces Then
                                   Fracc = Mid$(Fracc, 1, ._FieldDesc(Index).DecimalPlaces)
                               End if
                               Value = Mid$(Value, 1, InStr(0, Value, ".") - 1) + "." + Fracc
                            End If
                        Else
                            If InStr(0, Value, ".") <> False Then
                                Value = Mid$(Value, 1, InStr(0, Value, ".") - 1) 
                            End If
                        End If
                        FieldValue = Space$(._FieldDesc(Index).Length - Len(Value)) + Value
                    Case Else
                        ._CurrentError.Id = QDBF_ERR_INVALID_OPERATION
                        ._CurrentError.Description = "Field type not supported"
                        Exit Function
                End Select
                If Len(FieldValue) > ._FieldDesc(Index).Length Then
                    FieldValue = Mid$(FieldValue, 1, ._FieldDesc(Index).Length)
                End If
                ._CurrentRecord = Mid$(._CurrentRecord, 1,._FieldDesc(Index).Displacement) + FieldValue + Mid$(._CurrentRecord, ._FieldDesc(Index).Displacement + ._FieldDesc(Index).Length + 1) 
                .SetField = True
            End With
        End Function
            
        Function Save() As Integer
            Dim File As QFileStream
            Dim Term As Byte
            With This
                .Save = False
                If ._IsConnected = False Then
                    Exit Function
                ElseIf ._Mode = QDBF_NORMAL_MODE Then
                    ._CurrentError.Id = QDBF_ERR_INVALID_OPERATION
                    ._CurrentError.Description = "Database is not in append or edit mode"
                    Exit Function 
                ElseIf ._IsValidRecordIndex(._CurrentPosition) = False Then
                    Exit Function
                ElseIf ._OpenFile(File, ._CurrentPath, fmOpenWrite) = False Then
                    Exit Function
                End If
                ._Year = Val(Right$(Date$,2))
                ._Month = Val(Left$(Date$,2))
                ._Day = Val(Mid$(Date$,4,2))
                File.Position = 1
                File.Write(._Year)
                File.Write(._Month)
                File.Write(._Day)
                File.Write(._RecordCount)
                File.Position = (._CurrentPosition - 1) * ._RecordBytes + ._HeaderBytes
                File.WriteStr(._CurrentRecord, ._RecordBytes)
                If ._Mode = QDBF_APPEND_MODE Then
                    Term = 26
                    File.Write(Term)
                End If
                ._Mode = QDBF_NORMAL_MODE
                File.Close
				.Save = True
            End With
        End Function
        
        Function Cancel() As Integer
	      With This
                .Cancel = False
                If ._IsConnected = False Then
                    Exit Function
                ElseIf ((._Mode <> QDBF_EDIT_MODE) And (._Mode <> QDBF_APPEND_MODE)) Then
                    ._CurrentError.Id = QDBF_ERR_INVALID_OPERATION
                    ._CurrentError.Description = "Database is not in append or edit mode"
                    Exit Function
                ElseIf ._Mode = QDBF_APPEND_MODE Then
                    ._RecordCount -= 1
                    ._CurrentPosition = ._OldCurrentPosition
                    ._Mode = QDBF_NORMAL_MODE
					If ._RecordCount > 0 Then
						If .Go(._CurrentPosition) = False Then
							Exit Function
						End If
					End If
                End If
                ._Mode = QDBF_NORMAL_MODE
                .Cancel = True
            End With
        End Function
        
        Function Pack() As Integer
            'Dim SourceFile As QFileStream
            Dim TargetFile As QFileStream
            Dim RecordCount As DWord
            Dim FieldDesc As QDBFFieldDesc
            Dim Year As Byte
            Dim Month As Byte
            Dim Day As Byte
            Dim i As Word
            Dim Aux As Byte
            Dim TempDBFPath As String
            Dim Term As Byte
            With This
                .Pack = False
                If ._IsConnected = False Then
                    Exit Function
                ElseIf ((._Mode = QDBF_EDIT_MODE) Or (._Mode = QDBF_APPEND_MODE)) Then
                    ._CurrentError.Id = QDBF_ERR_INVALID_OPERATION
                    ._CurrentError.Description = "Database is in append or edit mode"
                    Exit Function
                End If
                TempDBFPath = Mid$(._CurrentPath, 1, RInStr(._CurrentPath, "\")) + "__" + Mid$(._CurrentPath, RInStr(._CurrentPath, "\") + 1, Len(._CurrentPath) - RInStr(._CurrentPath, "\"))
                If ._OpenFile(TargetFile, TempDBFPath, fmCreate) = True Then
                    TargetFile.Write(._Id)
                    Year = Val(Right$(Date$,2))
                    Month = Val(Left$(Date$,2))
                    Day = Val(Mid$(Date$,4,2))
                    TargetFile.Write(Year)
                    TargetFile.Write(Month)
                    TargetFile.Write(Day)
                    TargetFile.Write(._RecordCount)
                    TargetFile.Write(._HeaderBytes)
                    TargetFile.Write(._RecordBytes)
                    TargetFile.Write(._Reserved1)
                    TargetFile.Write(._TransactionFlag)
                    TargetFile.Write(._EncryptionFlag)
                    For i = 1 To 12
                        Aux = ._MultiUserEnv(i)
                        TargetFile.Write(Aux)
                    Next
                    TargetFile.Write(._ProductionIndexExist)
                    TargetFile.Write(._DB4LanguageDriverId)
                    TargetFile.Write(._Reserved2)
                    For i = 1 To ._FieldCount
                        FieldDesc.Name = ._FieldDesc(i).Name
                        FieldDesc.TypeId = ._FieldDesc(i).TypeId
                        FieldDesc.Displacement = ._FieldDesc(i).Displacement
                        FieldDesc.Length = ._FieldDesc(i).Length
                        FieldDesc.DecimalPlaces = ._FieldDesc(i).DecimalPlaces
                        FieldDesc.Reserved1 = ._FieldDesc(i).Reserved1
                        FieldDesc.WorkAreaId = ._FieldDesc(i).WorkAreaId
                        FieldDesc.Reserved2() = ._FieldDesc(i).Reserved2()
                        FieldDesc.HaveIndex = ._FieldDesc(i).HaveIndex
                        TargetFile.WriteUDT(FieldDesc)
					Next
                    Aux = 13
                    TargetFile.Write(Aux)
                    RecordCount = 0
                    ._CurrentPosition = 0
                    
                    While .Fetch
                        RecordCount += 1
                        TargetFile.Position = (RecordCount - 1) * ._RecordBytes + ._HeaderBytes
                        TargetFile.WriteStr(._CurrentRecord, ._RecordBytes)   
                    Wend
                    Term = 26
                    TargetFile.Write(Term)
                    TargetFile.Position = 4
                    TargetFile.Write(RecordCount)
                    TargetFile.Close
                    
                    If FileExists(Mid$(._CurrentPath, 1, RInStr(._CurrentPath, ".")) + "BAK") = True Then
                        Kill Mid$(._CurrentPath, 1, RInStr(._CurrentPath, ".")) + "BAK"
                    End If
                    Rename ._CurrentPath, Mid$(._CurrentPath, 1, RInStr(._CurrentPath, ".")) + "BAK"  
                    Rename TempDBFPath, ._CurrentPath
                    .Close
                    If .Connect(._CurrentPath) = False Then Exit Function
                Else
                    ._CurrentError.Id = QDBF_ERR_INVALID_OPERATION
                    ._CurrentError.Description = "Temporal database file could not be created"
                    Exit Function
                End If
                .Pack = True
            End With
        End Function
          
End Type

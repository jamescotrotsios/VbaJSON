Attribute VB_Name = "RecordsetToJsonHelper"
'@Description("Converts a DAO Recordset to JSON string following the simple schema format")
Public Function RecordsetToJSON(ByVal rs As DAO.Recordset) As String
    Dim sb As StringBuilder
    Set sb = StringBuilder.Create("{", 1024)
    
    ' Add fields array
    sb.Append """fields"":["
    
    Dim fld As DAO.field
    Dim fieldIndex As Long
    fieldIndex = 0
    
    ' Build fields array
    For Each fld In rs.fields
        If fieldIndex > 0 Then sb.Append ","
        
        sb.Append "{"
        sb.Append """name"":""" & EscapeJsonString(fld.Name) & ""","
        sb.Append """type"":""" & GetJsonDataType(fld.Type) & """"
        sb.Append "}"
        
        fieldIndex = fieldIndex + 1
    Next fld
    
    sb.Append "],"
    
    ' Add data array
    sb.Append """data"":["
    
    ' Move to first record if not at BOF
    If Not rs.BOF And Not rs.EOF Then
        rs.MoveFirst
    End If
    
    Dim rowIndex As Long
    rowIndex = 0
    
    ' Build data array
    Do While Not rs.EOF
        If rowIndex > 0 Then sb.Append ","
        
        sb.Append "{"
        
        fieldIndex = 0
        For Each fld In rs.fields
            If fieldIndex > 0 Then sb.Append ","
            
            sb.Append """" & EscapeJsonString(fld.Name) & """:"
            sb.Append FormatFieldValue(fld)
            
            fieldIndex = fieldIndex + 1
        Next fld
        
        sb.Append "}"
        
        rowIndex = rowIndex + 1
        rs.MoveNext
    Loop
    
    sb.Append "]"
    sb.Append "}"
    
    RecordsetToJSON = sb.ToString()
End Function

'@Description("Maps DAO field types to JSON schema data types")
Private Function GetJsonDataType(ByVal daoType As Long) As String
    Select Case daoType
        Case dbBoolean
            GetJsonDataType = "boolean"
        Case dbByte, dbInteger, dbLong
            GetJsonDataType = "integer"
        Case dbSingle, dbDouble
            GetJsonDataType = "number"
        Case dbCurrency, dbDecimal, dbNumeric
            GetJsonDataType = "decimal"
        Case dbDate, dbTime
            GetJsonDataType = "datetime"
        Case dbText, dbChar
            GetJsonDataType = "string"
        Case dbMemo, dbLongBinary
            GetJsonDataType = "text"
        Case dbBinary, dbVarBinary
            GetJsonDataType = "binary"
        Case dbGUID
            GetJsonDataType = "string"
        Case Else
            GetJsonDataType = "string"
    End Select
End Function

'@Description("Formats a field value as JSON")
Private Function FormatFieldValue(ByVal fld As DAO.field) As String
    ' Handle NULL values
    If IsNull(fld.value) Then
        FormatFieldValue = "null"
        Exit Function
    End If
    
    ' Format based on data type
    Select Case fld.Type
        Case dbBoolean
            FormatFieldValue = IIf(fld.value, "true", "false")
            
        Case dbByte, dbInteger, dbLong, dbSingle, dbDouble, dbCurrency, dbDecimal, dbNumeric
            ' Numbers don't need quotes
            FormatFieldValue = CStr(fld.value)
            
        Case dbDate, dbTime
            ' Format dates as ISO 8601 strings
            If IsDate(fld.value) Then
                FormatFieldValue = """" & Format$(fld.value, "yyyy-mm-dd\Thh:nn:ss") & """"
            Else
                FormatFieldValue = "null"
            End If
            
        Case dbText, dbMemo, dbChar, dbGUID
            ' Strings need quotes and escaping
            FormatFieldValue = """" & EscapeJsonString(CStr(fld.value)) & """"
            
        Case dbBinary, dbLongBinary, dbVarBinary
            ' Binary data - convert to base64 or hex string
            ' For simplicity, marking as empty string
            FormatFieldValue = """"""""
            
        Case Else
            ' Default to string
            FormatFieldValue = """" & EscapeJsonString(CStr(fld.value)) & """"
    End Select
End Function

'@Description("Escapes special characters in a string for JSON")
Private Function EscapeJsonString(ByVal text As String) As String
    Dim result As String
    result = text
    
    ' Replace backslashes first (must be done before other escapes)
    result = Replace(result, "\", "\\")
    
    ' Replace other special characters
    result = Replace(result, """", "\""")
    result = Replace(result, "/", "\/")
    result = Replace(result, vbBack, "\b")
    result = Replace(result, vbFormFeed, "\f")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbCr, "\r")
    result = Replace(result, vbTab, "\t")
    
    ' Handle other control characters (0x00-0x1F)
    Dim i As Long
    Dim char As String
    Dim charCode As Long
    Dim escapedResult As String
    escapedResult = ""
    
    For i = 1 To Len(result)
        char = Mid$(result, i, 1)
        charCode = AscW(char)
        
        If charCode < 32 And charCode <> 8 And charCode <> 9 And charCode <> 10 And charCode <> 12 And charCode <> 13 Then
            ' Control character that needs Unicode escape
            escapedResult = escapedResult & "\u" & Right$("0000" & Hex$(charCode), 4)
        Else
            escapedResult = escapedResult & char
        End If
    Next i
    
    EscapeJsonString = escapedResult
End Function

'@Description("Alternative version with pretty printing option")
Public Function RecordsetToJSONPretty(ByVal rs As DAO.Recordset, Optional ByVal indent As Boolean = True) As String
    Dim sb As StringBuilder
    Set sb = StringBuilder.Create("{", 1024)
    Dim newLine As String
    Dim lvTab As String
    
    If indent Then
        newLine = vbCrLf
        lvTab = "  "
    Else
        newLine = ""
        lvTab = ""
    End If
    
    ' Add fields array
    sb.Append newLine & lvTab & """fields"": ["
    
    Dim fld As DAO.field
    Dim fieldIndex As Long
    fieldIndex = 0
    
    ' Build fields array
    For Each fld In rs.fields
        If fieldIndex > 0 Then sb.Append ","
        
        sb.Append newLine & lvTab & lvTab & "{"
        sb.Append newLine & lvTab & lvTab & lvTab & """name"": """ & EscapeJsonString(fld.Name) & ""","
        sb.Append newLine & lvTab & lvTab & lvTab & """type"": """ & GetJsonDataType(fld.Type) & """"
        sb.Append newLine & lvTab & lvTab & "}"
        
        fieldIndex = fieldIndex + 1
    Next fld
    
    sb.Append newLine & lvTab & "],"
    
    ' Add data array
    sb.Append newLine & lvTab & """data"": ["
    
    ' Move to first record if not at BOF
    If Not rs.BOF And Not rs.EOF Then
        rs.MoveFirst
    End If
    
    Dim rowIndex As Long
    rowIndex = 0
    
    ' Build data array
    Do While Not rs.EOF
        If rowIndex > 0 Then sb.Append ","
        
        sb.Append newLine & lvTab & lvTab & "{"
        
        fieldIndex = 0
        For Each fld In rs.fields
            If fieldIndex > 0 Then sb.Append ","
            
            sb.Append newLine & lvTab & lvTab & lvTab & """" & EscapeJsonString(fld.Name) & """: "
            sb.Append FormatFieldValue(fld)
            
            fieldIndex = fieldIndex + 1
        Next fld
        
        sb.Append newLine & lvTab & lvTab & "}"
        
        rowIndex = rowIndex + 1
        rs.MoveNext
    Loop
    
    sb.Append newLine & lvTab & "]"
    sb.Append newLine & "}"
    
    RecordsetToJSONPretty = sb.ToString()
End Function

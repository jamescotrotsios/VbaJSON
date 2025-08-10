Attribute VB_Name = "JsonValidationHelper"
Option Compare Database
Option Explicit

'@Description("Validates that a JSON string complies with the DAO Recordset JSON Schema")
Public Function ValidateJSON(ByVal jsonString As String, _
                            Optional ByRef errorMessage As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ' Clear any previous error message
    errorMessage = ""
    
    ' Check for empty or null input
    If Len(Trim(jsonString)) = 0 Then
        errorMessage = "JSON string is empty"
        ValidateJSON = False
        Exit Function
    End If
    
    ' Check basic JSON structure (must start with { and end with })
    Dim trimmedJson As String
    trimmedJson = Trim(jsonString)
    
    If Left(trimmedJson, 1) <> "{" Or Right(trimmedJson, 1) <> "}" Then
        errorMessage = "JSON must be an object (start with { and end with })"
        ValidateJSON = False
        Exit Function
    End If
    
    ' Validate required top-level properties
    If Not ValidateRequiredProperties(jsonString, errorMessage) Then
        ValidateJSON = False
        Exit Function
    End If
    
    ' Validate fields array structure
    If Not ValidateFieldsArray(jsonString, errorMessage) Then
        ValidateJSON = False
        Exit Function
    End If
    
    ' Validate data array structure
    If Not ValidateDataArray(jsonString, errorMessage) Then
        ValidateJSON = False
        Exit Function
    End If
    
    ' Cross-validate that data objects use only defined field names
    If Not ValidateDataFieldNames(jsonString, errorMessage) Then
        ValidateJSON = False
        Exit Function
    End If
    
    ' Optional: Validate data types match field definitions
    If Not ValidateDataTypes(jsonString, errorMessage) Then
        ValidateJSON = False
        Exit Function
    End If
    
    ' All validations passed
    ValidateJSON = True
    Exit Function
    
ErrorHandler:
    errorMessage = "Validation error: " & Err.Description
    ValidateJSON = False
End Function

'@Description("Validates that required top-level properties exist")
Private Function ValidateRequiredProperties(ByVal jsonString As String, _
                                           ByRef errorMessage As String) As Boolean
    
    ' Check for "fields" property
    If InStr(jsonString, """fields""") = 0 Then
        errorMessage = "Missing required property: 'fields'"
        ValidateRequiredProperties = False
        Exit Function
    End If
    
    ' Check for "data" property
    If InStr(jsonString, """data""") = 0 Then
        errorMessage = "Missing required property: 'data'"
        ValidateRequiredProperties = False
        Exit Function
    End If
    
    ValidateRequiredProperties = True
End Function

'@Description("Validates the fields array structure")
Private Function ValidateFieldsArray(ByVal jsonString As String, _
                                    ByRef errorMessage As String) As Boolean
    On Error GoTo ValidationError
    
    ' Extract fields array content
    Dim fieldsStart As Long
    Dim fieldsEnd As Long
    
    fieldsStart = InStr(jsonString, """fields""")
    fieldsStart = InStr(fieldsStart, jsonString, "[")
    
    If fieldsStart = 0 Then
        errorMessage = "'fields' must be an array"
        ValidateFieldsArray = False
        Exit Function
    End If
    
    fieldsEnd = FindMatchingBracket(jsonString, fieldsStart, "[", "]")
    
    Dim fieldsContent As String
    fieldsContent = Mid$(jsonString, fieldsStart + 1, fieldsEnd - fieldsStart - 1)
    
    ' Check if fields array is empty
    If Len(Trim(fieldsContent)) = 0 Then
        errorMessage = "'fields' array cannot be empty"
        ValidateFieldsArray = False
        Exit Function
    End If
    
    ' Parse and validate each field object
    Dim pos As Long
    Dim fieldCount As Long
    pos = 1
    fieldCount = 0
    
    Do While pos < Len(fieldsContent)
        Dim objectStart As Long
        Dim objectEnd As Long
        
        objectStart = InStr(pos, fieldsContent, "{")
        If objectStart = 0 Then Exit Do
        
        objectEnd = FindMatchingBracket(fieldsContent, objectStart, "{", "}")
        
        Dim fieldObj As String
        fieldObj = Mid$(fieldsContent, objectStart + 1, objectEnd - objectStart - 1)
        
        ' Validate field object has required properties
        If Not ValidateFieldObject(fieldObj, fieldCount + 1, errorMessage) Then
            ValidateFieldsArray = False
            Exit Function
        End If
        
        fieldCount = fieldCount + 1
        pos = objectEnd + 1
    Loop
    
    If fieldCount = 0 Then
        errorMessage = "'fields' array must contain at least one field definition"
        ValidateFieldsArray = False
        Exit Function
    End If
    
    ValidateFieldsArray = True
    Exit Function
    
ValidationError:
    errorMessage = "Error validating fields array: " & Err.Description
    ValidateFieldsArray = False
End Function

'@Description("Validates individual field object structure")
Private Function ValidateFieldObject(ByVal fieldObj As String, _
                                    ByVal fieldIndex As Long, _
                                    ByRef errorMessage As String) As Boolean
    
    ' Check for "name" property
    If InStr(fieldObj, """name""") = 0 Then
        errorMessage = "Field at index " & fieldIndex & " missing required property: 'name'"
        ValidateFieldObject = False
        Exit Function
    End If
    
    ' Check for "type" property
    If InStr(fieldObj, """type""") = 0 Then
        errorMessage = "Field at index " & fieldIndex & " missing required property: 'type'"
        ValidateFieldObject = False
        Exit Function
    End If
    
    ' Extract and validate type value
    Dim typeValue As String
    typeValue = ExtractJsonValue(fieldObj, "type")
    
    If Not IsValidDataType(typeValue) Then
        errorMessage = "Field at index " & fieldIndex & " has invalid type: '" & typeValue & "'"
        ValidateFieldObject = False
        Exit Function
    End If
    
    ' Extract and validate name value
    Dim nameValue As String
    nameValue = ExtractJsonValue(fieldObj, "name")
    
    If Len(Trim(nameValue)) = 0 Then
        errorMessage = "Field at index " & fieldIndex & " has empty name"
        ValidateFieldObject = False
        Exit Function
    End If
    
    ValidateFieldObject = True
End Function

'@Description("Checks if a data type is valid according to schema")
Private Function IsValidDataType(ByVal dataType As String) As Boolean
    Select Case LCase(dataType)
        Case "string", "number", "integer", "boolean", _
             "date", "datetime", "decimal", "text", "binary"
            IsValidDataType = True
        Case Else
            IsValidDataType = False
    End Select
End Function

'@Description("Validates the data array structure")
Private Function ValidateDataArray(ByVal jsonString As String, _
                                  ByRef errorMessage As String) As Boolean
    On Error GoTo ValidationError
    
    ' Extract data array content
    Dim dataStart As Long
    Dim dataEnd As Long
    
    dataStart = InStr(jsonString, """data""")
    dataStart = InStr(dataStart, jsonString, "[")
    
    If dataStart = 0 Then
        errorMessage = "'data' must be an array"
        ValidateDataArray = False
        Exit Function
    End If
    
    dataEnd = FindMatchingBracket(jsonString, dataStart, "[", "]")
    
    Dim dataContent As String
    dataContent = Mid$(jsonString, dataStart + 1, dataEnd - dataStart - 1)
    
    ' Empty data array is valid
    If Len(Trim(dataContent)) = 0 Then
        ValidateDataArray = True
        Exit Function
    End If
    
    ' Validate each data object is properly formed
    Dim pos As Long
    Dim rowCount As Long
    pos = 1
    rowCount = 0
    
    Do While pos < Len(dataContent)
        Dim objectStart As Long
        Dim objectEnd As Long
        
        objectStart = InStr(pos, dataContent, "{")
        If objectStart = 0 Then Exit Do
        
        objectEnd = FindMatchingBracket(dataContent, objectStart, "{", "}")
        
        If objectEnd = 0 Then
            errorMessage = "Malformed object in data array at row " & (rowCount + 1)
            ValidateDataArray = False
            Exit Function
        End If
        
        rowCount = rowCount + 1
        pos = objectEnd + 1
    Loop
    
    ValidateDataArray = True
    Exit Function
    
ValidationError:
    errorMessage = "Error validating data array: " & Err.Description
    ValidateDataArray = False
End Function

'@Description("Validates that data objects only use defined field names")
Private Function ValidateDataFieldNames(ByVal jsonString As String, _
                                       ByRef errorMessage As String) As Boolean
    On Error GoTo ValidationError
    
    ' First, collect all defined field names
    Dim definedFields As Collection
    Set definedFields = GetDefinedFieldNames(jsonString)
    
    If definedFields.Count = 0 Then
        errorMessage = "No field definitions found"
        ValidateDataFieldNames = False
        Exit Function
    End If
    
    ' Extract data array
    Dim dataStart As Long
    Dim dataEnd As Long
    
    dataStart = InStr(jsonString, """data""")
    dataStart = InStr(dataStart, jsonString, "[")
    dataEnd = FindMatchingBracket(jsonString, dataStart, "[", "]")
    
    Dim dataContent As String
    dataContent = Mid$(jsonString, dataStart + 1, dataEnd - dataStart - 1)
    
    ' Check each data object
    Dim pos As Long
    Dim rowIndex As Long
    pos = 1
    rowIndex = 0
    
    Do While pos < Len(dataContent)
        Dim objectStart As Long
        Dim objectEnd As Long
        
        objectStart = InStr(pos, dataContent, "{")
        If objectStart = 0 Then Exit Do
        
        objectEnd = FindMatchingBracket(dataContent, objectStart, "{", "}")
        
        Dim rowObj As String
        rowObj = Mid$(dataContent, objectStart + 1, objectEnd - objectStart - 1)
        
        ' Extract field names from this data object
        Dim dataFieldNames As Collection
        Set dataFieldNames = ExtractFieldNamesFromObject(rowObj)
        
        ' Validate each field name exists in defined fields
        Dim fieldName As Variant
        For Each fieldName In dataFieldNames
            If Not IsFieldDefined(CStr(fieldName), definedFields) Then
                errorMessage = "Row " & (rowIndex + 1) & " contains undefined field: '" & fieldName & "'"
                ValidateDataFieldNames = False
                Exit Function
            End If
        Next fieldName
        
        rowIndex = rowIndex + 1
        pos = objectEnd + 1
    Loop
    
    ValidateDataFieldNames = True
    Exit Function
    
ValidationError:
    errorMessage = "Error validating field names: " & Err.Description
    ValidateDataFieldNames = False
End Function

'@Description("Gets all defined field names from the fields array")
Private Function GetDefinedFieldNames(ByVal jsonString As String) As Collection
    Dim fields As New Collection
    
    ' Extract fields array
    Dim fieldsStart As Long
    Dim fieldsEnd As Long
    
    fieldsStart = InStr(jsonString, """fields""")
    fieldsStart = InStr(fieldsStart, jsonString, "[")
    fieldsEnd = FindMatchingBracket(jsonString, fieldsStart, "[", "]")
    
    Dim fieldsContent As String
    fieldsContent = Mid$(jsonString, fieldsStart + 1, fieldsEnd - fieldsStart - 1)
    
    ' Parse each field object
    Dim pos As Long
    pos = 1
    
    Do While pos < Len(fieldsContent)
        Dim objectStart As Long
        Dim objectEnd As Long
        
        objectStart = InStr(pos, fieldsContent, "{")
        If objectStart = 0 Then Exit Do
        
        objectEnd = FindMatchingBracket(fieldsContent, objectStart, "{", "}")
        
        Dim fieldObj As String
        fieldObj = Mid$(fieldsContent, objectStart + 1, objectEnd - objectStart - 1)
        
        Dim fieldName As String
        fieldName = ExtractJsonValue(fieldObj, "name")
        
        If Len(fieldName) > 0 Then
            fields.Add fieldName, fieldName
        End If
        
        pos = objectEnd + 1
    Loop
    
    Set GetDefinedFieldNames = fields
End Function

'@Description("Extracts field names from a data object")
Private Function ExtractFieldNamesFromObject(ByVal objContent As String) As Collection
    Dim fieldNames As New Collection
    Dim pos As Long
    pos = 1
    
    Do While pos < Len(objContent)
        ' Find next field name
        Dim nameStart As Long
        Dim nameEnd As Long
        
        nameStart = InStr(pos, objContent, """")
        If nameStart = 0 Then Exit Do
        
        nameEnd = InStr(nameStart + 1, objContent, """")
        If nameEnd = 0 Then Exit Do
        
        Dim fieldName As String
        fieldName = Mid$(objContent, nameStart + 1, nameEnd - nameStart - 1)
        
        ' Add to collection if not already there
        On Error Resume Next
        fieldNames.Add fieldName, fieldName
        On Error GoTo 0
        
        ' Find the colon and skip to next field
        Dim colonPos As Long
        colonPos = InStr(nameEnd, objContent, ":")
        If colonPos = 0 Then Exit Do
        
        ' Skip the value to find next field
        pos = SkipJsonValue(objContent, colonPos + 1)
    Loop
    
    Set ExtractFieldNamesFromObject = fieldNames
End Function

'@Description("Skips over a JSON value to find the next position")
Private Function SkipJsonValue(ByVal text As String, ByVal startPos As Long) As Long
    Dim pos As Long
    pos = startPos
    
    ' Skip whitespace
    Do While pos <= Len(text) And Mid$(text, pos, 1) = " "
        pos = pos + 1
    Loop
    
    If pos > Len(text) Then
        SkipJsonValue = pos
        Exit Function
    End If
    
    Select Case Mid$(text, pos, 1)
        Case """"
            ' String - find closing quote
            pos = pos + 1
            Do While pos <= Len(text)
                If Mid$(text, pos, 1) = """" And Mid$(text, pos - 1, 1) <> "\" Then
                    pos = pos + 1
                    Exit Do
                End If
                pos = pos + 1
            Loop
            
        Case "{"
            ' Object - find closing brace
            pos = FindMatchingBracket(text, pos, "{", "}") + 1
            
        Case "["
            ' Array - find closing bracket
            pos = FindMatchingBracket(text, pos, "[", "]") + 1
            
        Case "t"
            ' true
            pos = pos + 4
            
        Case "f"
            ' false
            pos = pos + 5
            
        Case "n"
            ' null
            pos = pos + 4
            
        Case Else
            ' Number - continue until comma, space, or closing brace
            Do While pos <= Len(text)
                Dim ch As String
                ch = Mid$(text, pos, 1)
                If ch = "," Or ch = "}" Or ch = " " Or ch = vbCr Or ch = vbLf Then
                    Exit Do
                End If
                pos = pos + 1
            Loop
    End Select
    
    SkipJsonValue = pos
End Function

'@Description("Checks if a field name is defined")
Private Function IsFieldDefined(ByVal fieldName As String, _
                               ByVal definedFields As Collection) As Boolean
    On Error Resume Next
    Dim temp As Variant
    temp = definedFields(fieldName)
    IsFieldDefined = (Err.Number = 0)
    Err.Clear
End Function

'@Description("Validates that data values match their defined types")
Private Function ValidateDataTypes(ByVal jsonString As String, _
                                  ByRef errorMessage As String) As Boolean
    ' This is optional and can be computationally expensive
    ' Set to True to skip type validation for performance
    Dim skipTypeValidation As Boolean
    skipTypeValidation = False
    
    If skipTypeValidation Then
        ValidateDataTypes = True
        Exit Function
    End If
    
    ' Build field type map
    Dim fieldTypes As Object
    Set fieldTypes = CreateObject("Scripting.Dictionary")
    
    Dim fields As Collection
    Set fields = ParseFields(jsonString)
    
    Dim fieldInfo As Object
    For Each fieldInfo In fields
        fieldTypes(fieldInfo("name")) = fieldInfo("type")
    Next fieldInfo
    
    ' Validate each data row
    Dim dataRows As Collection
    Set dataRows = ParseData(jsonString)
    
    Dim rowIndex As Long
    Dim rowData As Object
    
    rowIndex = 0
    For Each rowData In dataRows
        rowIndex = rowIndex + 1
        
        Dim key As Variant
        For Each key In rowData.Keys
            Dim value As Variant
            value = rowData(key)
            
            If fieldTypes.Exists(key) Then
                If Not ValidateValueType(value, fieldTypes(key), key, rowIndex, errorMessage) Then
                    ValidateDataTypes = False
                    Exit Function
                End If
            End If
        Next key
    Next rowData
    
    ValidateDataTypes = True
End Function

'@Description("Validates that a value matches its expected type")
Private Function ValidateValueType(ByVal value As Variant, _
                                  ByVal expectedType As String, _
                                  ByVal fieldName As String, _
                                  ByVal rowIndex As Long, _
                                  ByRef errorMessage As String) As Boolean
    
    ' Null values are always valid
    If IsNull(value) Then
        ValidateValueType = True
        Exit Function
    End If
    
    Select Case LCase(expectedType)
        Case "boolean"
            If VarType(value) <> vbBoolean Then
                errorMessage = "Row " & rowIndex & ", field '" & fieldName & "': expected boolean, got " & TypeName(value)
                ValidateValueType = False
                Exit Function
            End If
            
        Case "integer"
            If Not IsNumeric(value) Then
                errorMessage = "Row " & rowIndex & ", field '" & fieldName & "': expected integer, got " & TypeName(value)
                ValidateValueType = False
                Exit Function
            End If
            
        Case "number", "decimal"
            If Not IsNumeric(value) Then
                errorMessage = "Row " & rowIndex & ", field '" & fieldName & "': expected number, got " & TypeName(value)
                ValidateValueType = False
                Exit Function
            End If
            
        Case "date", "datetime"
            If VarType(value) = vbString Then
                ' Try to parse as date
                If Not IsDate(value) And Not IsISODate(CStr(value)) Then
                    errorMessage = "Row " & rowIndex & ", field '" & fieldName & "': invalid date format"
                    ValidateValueType = False
                    Exit Function
                End If
            ElseIf VarType(value) <> vbDate Then
                errorMessage = "Row " & rowIndex & ", field '" & fieldName & "': expected date, got " & TypeName(value)
                ValidateValueType = False
                Exit Function
            End If
            
        Case "string", "text", "binary"
            ' Strings can be almost anything when parsed from JSON
            ' So we're lenient here
    End Select
    
    ValidateValueType = True
End Function

'@Description("Checks if a string is in ISO date format")
Private Function IsISODate(ByVal dateStr As String) As Boolean
    ' Basic check for ISO 8601 format: yyyy-mm-dd or yyyy-mm-ddThh:nn:ss
    If Len(dateStr) < 10 Then
        IsISODate = False
        Exit Function
    End If
    
    ' Check basic format
    If Mid$(dateStr, 5, 1) = "-" And Mid$(dateStr, 8, 1) = "-" Then
        IsISODate = True
    Else
        IsISODate = False
    End If
End Function

'@Description("Helper functions from previous artifacts")
Private Function FindMatchingBracket(ByVal text As String, ByVal startPos As Long, _
                                    ByVal openChar As String, ByVal closeChar As String) As Long
    Dim pos As Long
    Dim depth As Long
    Dim inString As Boolean
    Dim escapeNext As Boolean
    
    pos = startPos + 1
    depth = 1
    inString = False
    escapeNext = False
    
    Do While pos <= Len(text) And depth > 0
        Dim ch As String
        ch = Mid$(text, pos, 1)
        
        If escapeNext Then
            escapeNext = False
        ElseIf ch = "\" Then
            escapeNext = True
        ElseIf ch = """" And Not escapeNext Then
            inString = Not inString
        ElseIf Not inString Then
            If ch = openChar Then
                depth = depth + 1
            ElseIf ch = closeChar Then
                depth = depth - 1
            End If
        End If
        
        pos = pos + 1
    Loop
    
    FindMatchingBracket = pos - 1
End Function

Private Function ExtractJsonValue(ByVal jsonObj As String, ByVal key As String) As String
    Dim keyPos As Long
    keyPos = InStr(jsonObj, """" & key & """")
    
    If keyPos = 0 Then
        ExtractJsonValue = ""
        Exit Function
    End If
    
    Dim colonPos As Long
    colonPos = InStr(keyPos, jsonObj, ":")
    
    Dim valueStart As Long
    valueStart = InStr(colonPos, jsonObj, """")
    
    If valueStart = 0 Then
        ExtractJsonValue = ""
        Exit Function
    End If
    
    Dim valueEnd As Long
    valueEnd = InStr(valueStart + 1, jsonObj, """")
    
    ExtractJsonValue = Mid$(jsonObj, valueStart + 1, valueEnd - valueStart - 1)
End Function

Private Function ParseFields(ByVal jsonString As String) As Collection
    Dim fields As New Collection
    Dim fieldsStart As Long
    Dim fieldsEnd As Long
    
    fieldsStart = InStr(jsonString, """fields""")
    fieldsStart = InStr(fieldsStart, jsonString, "[")
    fieldsEnd = FindMatchingBracket(jsonString, fieldsStart, "[", "]")
    
    Dim fieldsContent As String
    fieldsContent = Mid$(jsonString, fieldsStart + 1, fieldsEnd - fieldsStart - 1)
    
    Dim pos As Long
    pos = 1
    
    Do While pos < Len(fieldsContent)
        Dim objectStart As Long
        Dim objectEnd As Long
        
        objectStart = InStr(pos, fieldsContent, "{")
        If objectStart = 0 Then Exit Do
        
        objectEnd = FindMatchingBracket(fieldsContent, objectStart, "{", "}")
        
        Dim fieldObj As String
        fieldObj = Mid$(fieldsContent, objectStart + 1, objectEnd - objectStart - 1)
        
        Dim fieldInfo As Object
        Set fieldInfo = CreateObject("Scripting.Dictionary")
        
        fieldInfo("name") = ExtractJsonValue(fieldObj, "name")
        fieldInfo("type") = ExtractJsonValue(fieldObj, "type")
        
        fields.Add fieldInfo
        
        pos = objectEnd + 1
    Loop
    
    Set ParseFields = fields
End Function

Private Function ParseData(ByVal jsonString As String) As Collection
    Dim dataRows As New Collection
    Dim dataStart As Long
    Dim dataEnd As Long
    
    dataStart = InStr(jsonString, """data""")
    dataStart = InStr(dataStart, jsonString, "[")
    dataEnd = FindMatchingBracket(jsonString, dataStart, "[", "]")
    
    Dim dataContent As String
    dataContent = Mid$(jsonString, dataStart + 1, dataEnd - dataStart - 1)
    
    Dim pos As Long
    pos = 1
    
    Do While pos < Len(dataContent)
        Dim objectStart As Long
        Dim objectEnd As Long
        
        objectStart = InStr(pos, dataContent, "{")
        If objectStart = 0 Then Exit Do
        
        objectEnd = FindMatchingBracket(dataContent, objectStart, "{", "}")
        
        Dim rowObj As String
        rowObj = Mid$(dataContent, objectStart + 1, objectEnd - objectStart - 1)
        
        Dim rowData As Object
        Set rowData = ParseRowObject(rowObj)
        
        dataRows.Add rowData
        
        pos = objectEnd + 1
    Loop
    
    Set ParseData = dataRows
End Function

Private Function ParseRowObject(ByVal rowJson As String) As Object
    Dim rowData As Object
    Set rowData = CreateObject("Scripting.Dictionary")
    
    Dim pos As Long
    pos = 1
    
    Do While pos < Len(rowJson)
        Dim nameStart As Long
        Dim nameEnd As Long
        
        nameStart = InStr(pos, rowJson, """")
        If nameStart = 0 Then Exit Do
        
        nameEnd = InStr(nameStart + 1, rowJson, """")
        If nameEnd = 0 Then Exit Do
        
        Dim fieldName As String
        fieldName = Mid$(rowJson, nameStart + 1, nameEnd - nameStart - 1)
        
        Dim colonPos As Long
        colonPos = InStr(nameEnd, rowJson, ":")
        If colonPos = 0 Then Exit Do
        
        Dim valueStart As Long
        valueStart = colonPos + 1
        
        Do While valueStart <= Len(rowJson) And Mid$(rowJson, valueStart, 1) = " "
            valueStart = valueStart + 1
        Loop
        
        Dim value As Variant
        Dim valueEnd As Long
        
        Select Case Mid$(rowJson, valueStart, 1)
            Case """"
                valueEnd = InStr(valueStart + 1, rowJson, """")
                value = Mid$(rowJson, valueStart + 1, valueEnd - valueStart - 1)
                pos = valueEnd + 1
                
            Case "t", "f"
                If Mid$(rowJson, valueStart, 4) = "true" Then
                    value = True
                    pos = valueStart + 4
                ElseIf Mid$(rowJson, valueStart, 5) = "false" Then
                    value = False
                    pos = valueStart + 5
                End If
                
            Case "n"
                If Mid$(rowJson, valueStart, 4) = "null" Then
                    value = Null
                    pos = valueStart + 4
                End If
                
            Case Else
                valueEnd = valueStart
                Do While valueEnd <= Len(rowJson)
                    Dim ch As String
                    ch = Mid$(rowJson, valueEnd, 1)
                    If ch = "," Or ch = "}" Or ch = " " Or ch = vbCr Or ch = vbLf Then
                        Exit Do
                    End If
                    valueEnd = valueEnd + 1
                Loop
                
                Dim numStr As String
                numStr = Trim$(Mid$(rowJson, valueStart, valueEnd - valueStart))
                
                If InStr(numStr, ".") > 0 Then
                    value = CDbl(numStr)
                Else
                    value = CLng(numStr)
                End If
                pos = valueEnd
        End Select
        
        rowData(fieldName) = value
        
        Dim commaPos As Long
        commaPos = InStr(pos, rowJson, ",")
        If commaPos > 0 Then
            pos = commaPos + 1
        Else
            Exit Do
        End If
    Loop
    
    Set ParseRowObject = rowData
End Function


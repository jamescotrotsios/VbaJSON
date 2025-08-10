Attribute VB_Name = "JsonToRecordsetHelper"
'Requires reference to: Microsoft ActiveX Data Objects Library (any version 2.x or higher)

'@Description("Converts a JSON string to an in-memory ADODB Recordset (no temp table required)")
Public Function JSONToRecordset(ByVal jsonString As String) As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    ' Create a new disconnected recordset
    Dim rs As New ADODB.Recordset
    
    ' Parse the JSON to extract fields and data
    Dim fields As Collection
    Dim dataRows As Collection
    
    Set fields = ParseFields(jsonString)
    Set dataRows = ParseData(jsonString)
    
    ' Define the recordset structure based on fields
    Dim fieldInfo As Object
    For Each fieldInfo In fields
        rs.fields.Append fieldInfo("name"), GetADODataType(fieldInfo("type")), _
                         GetFieldSize(fieldInfo("type")), adFldMayBeNull
    Next fieldInfo
    
    ' Open the recordset in memory (disconnected)
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic
    rs.Open
    
    ' Populate the recordset with data
    Dim rowData As Object
    For Each rowData In dataRows
        rs.AddNew
        
        For Each fieldInfo In fields
            Dim fieldName As String
            fieldName = fieldInfo("name")
            
            If rowData.Exists(fieldName) Then
                Dim value As Variant
                value = rowData(fieldName)
                
                If Not IsNull(value) Then
                    ' Handle date strings
                    If fieldInfo("type") = "date" Or fieldInfo("type") = "datetime" Then
                        If VarType(value) = vbString Then
                            value = ParseISODate(CStr(value))
                        End If
                    End If
                    
                    rs.fields(fieldName).value = value
                Else
                    rs.fields(fieldName).value = Null
                End If
            End If
        Next fieldInfo
        
        rs.Update
    Next rowData
    
    ' Move to first record if we have data
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    
    Set JSONToRecordset = rs
    Exit Function
    
ErrorHandler:
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    Err.Raise Err.Number, "JSONToRecordset", "Error converting JSON to Recordset: " & Err.Description
End Function

'@Description("Maps JSON types to ADO data types")
Private Function GetADODataType(ByVal jsonType As String) As ADODB.DataTypeEnum
    Select Case LCase(jsonType)
        Case "boolean"
            GetADODataType = adBoolean
        Case "integer"
            GetADODataType = adInteger
        Case "number"
            GetADODataType = adDouble
        Case "decimal"
            GetADODataType = adCurrency
        Case "date", "datetime"
            GetADODataType = adDate
        Case "text"
            GetADODataType = adLongVarWChar
        Case "binary"
            GetADODataType = adLongVarBinary
        Case Else ' "string" and default
            GetADODataType = adVarWChar
    End Select
End Function

'@Description("Gets appropriate field size for data type")
Private Function GetFieldSize(ByVal jsonType As String) As Long
    Select Case LCase(jsonType)
        Case "string"
            GetFieldSize = 255
        Case "text"
            GetFieldSize = 2147483647 ' Max long value for memo
        Case "binary"
            GetFieldSize = 2147483647
        Case Else
            GetFieldSize = 0 ' Size not applicable for numbers, dates, etc.
    End Select
End Function

'@Description("Alternative version that returns a DAO-compatible recordset wrapper")
Public Function JSONToDAOCompatible(ByVal jsonString As String) As Object
    ' This returns an ADODB recordset that can be used similarly to DAO
    Set JSONToDAOCompatible = JSONToRecordset(jsonString)
End Function

'@Description("Parses the fields array from JSON")
Private Function ParseFields(ByVal jsonString As String) As Collection
    Dim fields As New Collection
    Dim fieldsStart As Long
    Dim fieldsEnd As Long
    
    ' Find the "fields" array
    fieldsStart = InStr(jsonString, """fields""")
    If fieldsStart = 0 Then
        Err.Raise vbObjectError + 1001, "ParseFields", "No 'fields' array found in JSON"
    End If
    
    ' Find the opening bracket of the fields array
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
        
        ' Extract name and type from field object
        Dim fieldInfo As Object
        Set fieldInfo = CreateObject("Scripting.Dictionary")
        
        fieldInfo("name") = ExtractJsonValue(fieldObj, "name")
        fieldInfo("type") = ExtractJsonValue(fieldObj, "type")
        
        fields.Add fieldInfo
        
        pos = objectEnd + 1
    Loop
    
    Set ParseFields = fields
End Function

'@Description("Parses the data array from JSON")
Private Function ParseData(ByVal jsonString As String) As Collection
    Dim dataRows As New Collection
    Dim dataStart As Long
    Dim dataEnd As Long
    
    ' Find the "data" array
    dataStart = InStr(jsonString, """data""")
    If dataStart = 0 Then
        Err.Raise vbObjectError + 1002, "ParseData", "No 'data' array found in JSON"
    End If
    
    ' Find the opening bracket of the data array
    dataStart = InStr(dataStart, jsonString, "[")
    dataEnd = FindMatchingBracket(jsonString, dataStart, "[", "]")
    
    Dim dataContent As String
    dataContent = Mid$(jsonString, dataStart + 1, dataEnd - dataStart - 1)
    
    ' Parse each data object
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
        
        ' Parse the row object into a dictionary
        Dim rowData As Object
        Set rowData = ParseRowObject(rowObj)
        
        dataRows.Add rowData
        
        pos = objectEnd + 1
    Loop
    
    Set ParseData = dataRows
End Function

'@Description("Parses a single row object into a dictionary")
Private Function ParseRowObject(ByVal rowJson As String) As Object
    Dim rowData As Object
    Set rowData = CreateObject("Scripting.Dictionary")
    
    Dim pos As Long
    pos = 1
    
    Do While pos < Len(rowJson)
        ' Find next field name
        Dim nameStart As Long
        Dim nameEnd As Long
        
        nameStart = InStr(pos, rowJson, """")
        If nameStart = 0 Then Exit Do
        
        nameEnd = InStr(nameStart + 1, rowJson, """")
        If nameEnd = 0 Then Exit Do
        
        Dim fieldName As String
        fieldName = Mid$(rowJson, nameStart + 1, nameEnd - nameStart - 1)
        
        ' Find the colon
        Dim colonPos As Long
        colonPos = InStr(nameEnd, rowJson, ":")
        If colonPos = 0 Then Exit Do
        
        ' Find the value
        Dim valueStart As Long
        valueStart = colonPos + 1
        
        ' Skip whitespace
        Do While valueStart <= Len(rowJson) And Mid$(rowJson, valueStart, 1) = " "
            valueStart = valueStart + 1
        Loop
        
        ' Parse the value
        Dim value As Variant
        Dim valueEnd As Long
        
        Select Case Mid$(rowJson, valueStart, 1)
            Case """"
                ' String value
                valueEnd = InStr(valueStart + 1, rowJson, """")
                value = UnescapeJsonString(Mid$(rowJson, valueStart + 1, valueEnd - valueStart - 1))
                pos = valueEnd + 1
                
            Case "{"
                ' Object (not supported in simple schema)
                valueEnd = FindMatchingBracket(rowJson, valueStart, "{", "}")
                value = Null
                pos = valueEnd + 1
                
            Case "["
                ' Array (not supported in simple schema)
                valueEnd = FindMatchingBracket(rowJson, valueStart, "[", "]")
                value = Null
                pos = valueEnd + 1
                
            Case "t", "f"
                ' Boolean
                If Mid$(rowJson, valueStart, 4) = "true" Then
                    value = True
                    pos = valueStart + 4
                ElseIf Mid$(rowJson, valueStart, 5) = "false" Then
                    value = False
                    pos = valueStart + 5
                End If
                
            Case "n"
                ' null
                If Mid$(rowJson, valueStart, 4) = "null" Then
                    value = Null
                    pos = valueStart + 4
                End If
                
            Case Else
                ' Number - find the end of the number
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
        
        ' Find next comma or end
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

'@Description("Finds the matching closing bracket")
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

'@Description("Extracts a value from a JSON object string")
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
    
    Dim valueEnd As Long
    valueEnd = InStr(valueStart + 1, jsonObj, """")
    
    ExtractJsonValue = Mid$(jsonObj, valueStart + 1, valueEnd - valueStart - 1)
End Function

'@Description("Unescapes JSON string special characters")
Private Function UnescapeJsonString(ByVal text As String) As String
    Dim result As String
    result = text
    
    ' Replace escape sequences
    result = Replace(result, "\""", """")
    result = Replace(result, "\\", "\")
    result = Replace(result, "\/", "/")
    result = Replace(result, "\b", vbBack)
    result = Replace(result, "\f", vbFormFeed)
    result = Replace(result, "\n", vbLf)
    result = Replace(result, "\r", vbCr)
    result = Replace(result, "\t", vbTab)
    
    ' Handle Unicode escapes \uXXXX
    Dim pos As Long
    pos = InStr(result, "\u")
    
    While pos > 0
        If pos + 5 <= Len(result) Then
            Dim hexCode As String
            hexCode = Mid$(result, pos + 2, 4)
            
            If IsHex(hexCode) Then
                Dim charCode As Long
                charCode = CLng("&H" & hexCode)
                result = Left$(result, pos - 1) & ChrW$(charCode) & Mid$(result, pos + 6)
            End If
        End If
        
        pos = InStr(pos + 1, result, "\u")
    Wend
    
    UnescapeJsonString = result
End Function

'@Description("Checks if a string is valid hexadecimal")
Private Function IsHex(ByVal text As String) As Boolean
    Dim i As Long
    For i = 1 To Len(text)
        Dim ch As String
        ch = Mid$(text, i, 1)
        If Not ((ch >= "0" And ch <= "9") Or _
                (ch >= "A" And ch <= "F") Or _
                (ch >= "a" And ch <= "f")) Then
            IsHex = False
            Exit Function
        End If
    Next i
    IsHex = True
End Function

'@Description("Parses ISO date string to Date value")
Private Function ParseISODate(ByVal isoDate As String) As Date
    ' Handle ISO 8601 format: yyyy-mm-ddThh:nn:ss
    Dim cleanDate As String
    cleanDate = Replace(isoDate, "T", " ")
    cleanDate = Replace(cleanDate, "Z", "")
    
    ' Remove milliseconds if present
    Dim dotPos As Long
    dotPos = InStr(cleanDate, ".")
    If dotPos > 0 Then
        cleanDate = Left$(cleanDate, dotPos - 1)
    End If
    
    On Error Resume Next
    ParseISODate = CDate(cleanDate)
    If Err.Number <> 0 Then
        ParseISODate = Now() ' Default to current date if parsing fails
    End If
    On Error GoTo 0
End Function

'@Description("Example usage with ADODB recordset")
Public Sub TestJSONToRecordset()
    Dim jsonString As String
    jsonString = "{" & _
        """fields"": [" & _
        "  {""name"": ""CustomerID"", ""type"": ""string""}," & _
        "  {""name"": ""CompanyName"", ""type"": ""string""}," & _
        "  {""name"": ""OrderCount"", ""type"": ""integer""}," & _
        "  {""name"": ""IsActive"", ""type"": ""boolean""}" & _
        "]," & _
        """data"": [" & _
        "  {""CustomerID"": ""ALFKI"", ""CompanyName"": ""Alfreds Futterkiste"", ""OrderCount"": 6, ""IsActive"": true}," & _
        "  {""CustomerID"": ""ANATR"", ""CompanyName"": ""Ana Trujillo"", ""OrderCount"": 4, ""IsActive"": false}" & _
        "]" & _
        "}"
    
    Dim rs As ADODB.Recordset
    Set rs = JSONToRecordset(jsonString)
    
    ' Display the results
    Do While Not rs.EOF
        Debug.Print rs!CustomerID, rs!CompanyName, rs!OrderCount, rs!IsActive
        rs.MoveNext
    Loop
    
    ' Can also use it like a DAO recordset for most operations
    rs.MoveFirst
    Debug.Print "Record count: " & rs.RecordCount
    
    ' Can update values
    rs!OrderCount = 10
    rs.Update
    
    ' Can filter
    rs.Filter = "IsActive = true"
    
    ' Can sort
    rs.Sort = "CompanyName ASC"
    
    rs.Close
    Set rs = Nothing
End Sub

'@Description("Helper function to convert ADODB recordset to array for binding")
Public Function RecordsetToArray(rs As ADODB.Recordset) As Variant
    If rs.RecordCount = 0 Then
        RecordsetToArray = Empty
        Exit Function
    End If
    
    rs.MoveFirst
    RecordsetToArray = rs.GetRows()
End Function

'@Description("Saves a disconnected ADODB recordset to a permanent Access table")
Public Function SaveRecordsetToTable(ByVal rs As ADODB.Recordset, ByVal tableName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Validate inputs
    If rs Is Nothing Then
        SaveRecordsetToTable = False
        Exit Function
    End If
    
    If rs.State <> adStateOpen Then
        SaveRecordsetToTable = False
        Exit Function
    End If
    
    If Len(Trim(tableName)) = 0 Then
        SaveRecordsetToTable = False
        Exit Function
    End If
    
    Dim db As DAO.Database
    Set db = CurrentDb()
    
    ' Check if table exists
    Dim lvTableExists As Boolean
    lvTableExists = TableExists(tableName, db)
    
    ' If table doesn't exist, create it
    If Not lvTableExists Then
        If Not CreateTableFromRecordset(rs, tableName, db) Then
            SaveRecordsetToTable = False
            GoTo Cleanup
        End If
    Else
        ' If table exists, verify structure compatibility
        If Not VerifyTableStructure(rs, tableName, db) Then
            ' Structure doesn't match - could either fail or try to adapt
            ' For safety, we'll fail rather than potentially lose data
            SaveRecordsetToTable = False
            GoTo Cleanup
        End If
    End If
    
    ' Save records to the table
    If Not AppendRecordsToTable(rs, tableName, db) Then
        SaveRecordsetToTable = False
        GoTo Cleanup
    End If
    
    ' Success
    SaveRecordsetToTable = True
    
Cleanup:
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    SaveRecordsetToTable = False
    Set db = Nothing
End Function

'@Description("Checks if a table exists in the database")
Private Function TableExists(ByVal tableName As String, ByVal db As DAO.Database) As Boolean
    On Error Resume Next
    Dim tdf As DAO.TableDef
    Set tdf = db.TableDefs(tableName)
    TableExists = (Err.Number = 0)
    Set tdf = Nothing
    Err.Clear
End Function

'@Description("Creates a new table based on ADODB recordset structure")
Private Function CreateTableFromRecordset(ByVal rs As ADODB.Recordset, _
                                         ByVal tableName As String, _
                                         ByVal db As DAO.Database) As Boolean
    On Error GoTo ErrorHandler
    
    Dim tdf As DAO.TableDef
    Set tdf = db.CreateTableDef(tableName)
    
    Dim adoField As ADODB.field
    Dim daoField As DAO.field
    
    ' Create DAO fields based on ADODB fields
    For Each adoField In rs.fields
        Set daoField = CreateDAOField(adoField, tdf)
        If Not daoField Is Nothing Then
            tdf.fields.Append daoField
        End If
    Next adoField
    
    ' Append the table to the database
    db.TableDefs.Append tdf
    
    CreateTableFromRecordset = True
    Exit Function
    
ErrorHandler:
    CreateTableFromRecordset = False
End Function

'@Description("Creates a DAO field based on an ADODB field")
Private Function CreateDAOField(ByVal adoField As ADODB.field, _
                               ByVal tdf As DAO.TableDef) As DAO.field
    On Error GoTo ErrorHandler
    
    Dim daoField As DAO.field
    Dim daoType As Long
    Dim fieldSize As Long
    
    ' Map ADODB data type to DAO data type
    Select Case adoField.Type
        Case adBoolean
            daoType = dbBoolean
            fieldSize = 0
            
        Case adTinyInt, adSmallInt, adInteger
            daoType = dbLong
            fieldSize = 0
            
        Case adBigInt
            daoType = dbLong
            fieldSize = 0
            
        Case adSingle
            daoType = dbSingle
            fieldSize = 0
            
        Case adDouble
            daoType = dbDouble
            fieldSize = 0
            
        Case adCurrency
            daoType = dbCurrency
            fieldSize = 0
            
        Case adDecimal, adNumeric
            daoType = dbDouble
            fieldSize = 0
            
        Case adDate, adDBDate, adDBTime, adDBTimeStamp
            daoType = dbDate
            fieldSize = 0
            
        Case adChar, adVarChar, adWChar, adVarWChar
            daoType = dbText
            fieldSize = IIf(adoField.DefinedSize > 255, 255, adoField.DefinedSize)
            If fieldSize = 0 Then fieldSize = 255
            
        Case adLongVarChar, adLongVarWChar
            daoType = dbMemo
            fieldSize = 0
            
        Case adBinary, adVarBinary
            daoType = dbBinary
            fieldSize = IIf(adoField.DefinedSize > 510, 510, adoField.DefinedSize)
            
        Case adLongVarBinary
            daoType = dbLongBinary
            fieldSize = 0
            
        Case adGUID
            daoType = dbText
            fieldSize = 38
            
        Case Else
            ' Default to text
            daoType = dbText
            fieldSize = 255
    End Select
    
    ' Create the field
    If fieldSize > 0 Then
        Set daoField = tdf.CreateField(adoField.Name, daoType, fieldSize)
    Else
        Set daoField = tdf.CreateField(adoField.Name, daoType)
    End If
    
    ' Set field properties
    daoField.Required = False  ' Allow nulls for flexibility
    
    If daoType = dbText Then
        daoField.AllowZeroLength = True
    End If
    
    Set CreateDAOField = daoField
    Exit Function
    
ErrorHandler:
    Set CreateDAOField = Nothing
End Function

'@Description("Verifies that table structure is compatible with recordset")
Private Function VerifyTableStructure(ByVal rs As ADODB.Recordset, _
                                     ByVal tableName As String, _
                                     ByVal db As DAO.Database) As Boolean
    On Error GoTo ErrorHandler
    
    Dim tdf As DAO.TableDef
    Set tdf = db.TableDefs(tableName)
    
    Dim adoField As ADODB.field
    Dim daoField As DAO.field
    
    ' Check if all recordset fields exist in the table
    For Each adoField In rs.fields
        On Error Resume Next
        Set daoField = tdf.fields(adoField.Name)
        If Err.Number <> 0 Then
            ' Field doesn't exist in table
            VerifyTableStructure = False
            Exit Function
        End If
        On Error GoTo ErrorHandler
    Next adoField
    
    VerifyTableStructure = True
    Exit Function
    
ErrorHandler:
    VerifyTableStructure = False
End Function

'@Description("Appends records from ADODB recordset to Access table")
Private Function AppendRecordsToTable(ByVal rs As ADODB.Recordset, _
                                     ByVal tableName As String, _
                                     ByVal db As DAO.Database) As Boolean
    On Error GoTo ErrorHandler
    
    ' If recordset is empty, nothing to do but still successful
    If rs.RecordCount = 0 Then
        AppendRecordsToTable = True
        Exit Function
    End If
    
    ' Open DAO recordset for the table
    Dim daoRS As DAO.Recordset
    Set daoRS = db.OpenRecordset(tableName, dbOpenDynaset)
    
    ' Move to first record in ADODB recordset
    rs.MoveFirst
    
    ' Copy records
    Dim adoField As ADODB.field
    Dim fieldName As String
    Dim fieldValue As Variant
    
    Do While Not rs.EOF
        daoRS.AddNew
        
        For Each adoField In rs.fields
            fieldName = adoField.Name
            fieldValue = adoField.value
            
            ' Check if field exists in DAO recordset
            On Error Resume Next
            If Not IsNull(fieldValue) Then
                daoRS.fields(fieldName).value = fieldValue
            Else
                daoRS.fields(fieldName).value = Null
            End If
            On Error GoTo ErrorHandler
        Next adoField
        
        daoRS.Update
        rs.MoveNext
    Loop
    
    ' Clean up
    daoRS.Close
    Set daoRS = Nothing
    
    AppendRecordsToTable = True
    Exit Function
    
ErrorHandler:
    If Not daoRS Is Nothing Then
        If daoRS.EditMode <> dbEditNone Then
            daoRS.CancelUpdate
        End If
        daoRS.Close
        Set daoRS = Nothing
    End If
    AppendRecordsToTable = False
End Function

'@Description("Alternative version that optionally clears existing data before appending")
Public Function SaveRecordsetToTableWithOptions(ByVal rs As ADODB.Recordset, _
                                               ByVal tableName As String, _
                                               Optional ByVal clearExistingData As Boolean = False, _
                                               Optional ByVal createBackup As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    
    ' Validate inputs
    If rs Is Nothing Or rs.State <> adStateOpen Or Len(Trim(tableName)) = 0 Then
        SaveRecordsetToTableWithOptions = False
        Exit Function
    End If
    
    Dim db As DAO.Database
    Set db = CurrentDb()
    
    ' Create backup if requested and table exists
    If createBackup And TableExists(tableName, db) Then
        Dim backupName As String
        backupName = tableName & "_Backup_" & Format(Now(), "yyyymmdd_hhnnss")
        db.Execute "SELECT * INTO [" & backupName & "] FROM [" & tableName & "]", dbFailOnError
    End If
    
    ' Clear existing data if requested and table exists
    If clearExistingData And TableExists(tableName, db) Then
        db.Execute "DELETE FROM [" & tableName & "]", dbFailOnError
    End If
    
    ' Use the main function to save the data
    SaveRecordsetToTableWithOptions = SaveRecordsetToTable(rs, tableName)
    
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    SaveRecordsetToTableWithOptions = False
    Set db = Nothing
End Function


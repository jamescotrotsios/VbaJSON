Attribute VB_Name = "TestJsonHelper"

'@Description("Example usage function")
Public Sub ExampleUsage()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim jsonOutput As String, lvSQL As String
    
    
    lvSQL = "SELECT UserInfo.Name,UserInfo.Account,UserInfo.[E-Mail] FROM UserInfo WHERE (((UserInfo.[E-Mail]) IS NOT NULL));"
    
    ' Open database and recordset
    Set db = CurrentDb() ' or OpenDatabase("path\to\database.mdb")
    Set rs = db.OpenRecordset(lvSQL)
    
    
    ' Convert to JSON
    jsonOutput = RecordsetToJSON(rs)
    
    ' Output to immediate window
    Debug.Print jsonOutput
    
    ' Or save to file
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim textFile As Object
    Set textFile = fso.CreateTextFile("C:\data\output.json", True)
    textFile.Write jsonOutput
    textFile.Close
    
    ' Cleanup
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub

'@Description("Example usage")
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
    
    Dim rs As DAO.Recordset
    Set rs = JSONToRecordset(jsonString)
    
    ' Display the results
    Do While Not rs.EOF
        Debug.Print rs!CustomerID, rs!CompanyName, rs!OrderCount, rs!IsActive
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub

'@Description("Example usage of SaveRecordsetToTable")
Public Sub TestSaveRecordsetToTable()
    ' Create a sample JSON string
    Dim jsonString As String
    jsonString = "{" & _
        """fields"": [" & _
        "  {""name"": ""CustomerID"", ""type"": ""string""}," & _
        "  {""name"": ""CompanyName"", ""type"": ""string""}," & _
        "  {""name"": ""OrderCount"", ""type"": ""integer""}" & _
        "]," & _
        """data"": [" & _
        "  {""CustomerID"": ""ALFKI"", ""CompanyName"": ""Alfreds Futterkiste"", ""OrderCount"": 6}," & _
        "  {""CustomerID"": ""ANATR"", ""CompanyName"": ""Ana Trujillo"", ""OrderCount"": 4}" & _
        "]" & _
        "}"
    
    ' Convert JSON to recordset
    Dim rs As ADODB.Recordset
    Set rs = JSONToRecordset(jsonString)
    
    ' Save to table
    Dim success As Boolean
    success = SaveRecordsetToTable(rs, "ImportedCompanyData")
    
    If success Then
        Debug.Print "Data successfully saved to table!"
    Else
        Debug.Print "Failed to save data to table."
    End If
    
    ' Clean up
    rs.Close
    Set rs = Nothing
End Sub

'@Description("Example usage of validation")
Public Sub TestJSONValidation()
    Dim jsonString As String
    Dim errorMsg As String
    Dim isValid As Boolean
    
    ' Valid JSON
    jsonString = "{" & _
        """fields"": [" & _
        "  {""name"": ""ID"", ""type"": ""integer""}," & _
        "  {""name"": ""Name"", ""type"": ""string""}" & _
        "]," & _
        """data"": [" & _
        "  {""ID"": 1, ""Name"": ""John""}" & _
        "]" & _
        "}"
    
    isValid = ValidateJSON(jsonString, errorMsg)
    Debug.Print "Valid JSON test: " & isValid
    If Not isValid Then Debug.Print "Error: " & errorMsg
    
    ' Invalid JSON - missing fields
    jsonString = "{""data"": []}"
    
    isValid = ValidateJSON(jsonString, errorMsg)
    Debug.Print "Invalid JSON test: " & isValid
    If Not isValid Then Debug.Print "Error: " & errorMsg
    
    ' Invalid JSON - bad field type
    jsonString = "{" & _
        """fields"": [" & _
        "  {""name"": ""ID"", ""type"": ""invalid_type""}" & _
        "]," & _
        """data"": []" & _
        "}"
    
    isValid = ValidateJSON(jsonString, errorMsg)
    Debug.Print "Invalid type test: " & isValid
    If Not isValid Then Debug.Print "Error: " & errorMsg
    
    ' Invalid JSON - undefined field in data
    jsonString = "{" & _
        """fields"": [" & _
        "  {""name"": ""ID"", ""type"": ""integer""}" & _
        "]," & _
        """data"": [" & _
        "  {""ID"": 1, ""UndefinedField"": ""value""}" & _
        "]" & _
        "}"
    
    isValid = ValidateJSON(jsonString, errorMsg)
    Debug.Print "Undefined field test: " & isValid
    If Not isValid Then Debug.Print "Error: " & errorMsg
End Sub



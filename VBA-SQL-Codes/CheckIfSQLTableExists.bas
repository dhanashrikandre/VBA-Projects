Attribute VB_Name = "CheckIfSQLTableExists"
Sub CheckIfSQLTableExists() 'check if abc table exist in code
    
    Dim conn As Object
    Dim rs As Object
    Dim connectionString As String
    
    tableName = newabcname
    
    connectionString = "Provider=SQLOLEDB;Data Source=server;Initial Catalog=Data;User ID=Enter_Username;Password=pwd"
       
    ' Create a connection object
    Set conn = CreateObject("ADODB.Connection")
    
    ' Open the connection
    conn.Open connectionString
    
    ' Create a recordset object
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Use the recordset to execute a query to check if the table exists
    rs.Open "SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '" & tableName & "'", conn
    
    ' Check if any records were returned
    If Not rs.EOF Then
        MsgBox "Table '" & tableName & "' exists in the database.", vbInformation
    Else
        MsgBox "Table '" & tableName & "' does not exist in the database.", vbExclamation
    End If
    
    ' Close the connections
    rs.Close
    conn.Close
End Sub


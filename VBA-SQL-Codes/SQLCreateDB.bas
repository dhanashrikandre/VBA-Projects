Attribute VB_Name = "SQLCreateDB"
Sub SQLUpdate_DateTime()  'sql create dummy database
    Dim conn As Object
    Dim cmd As Object
    Dim connString As String
    Dim dbName As String
    Dim currentDate As Date
    
    ' Replace these with your SQL Server details
    connString = "Provider=SQLOLEDB;Data Source=1.1.11.1,11;Initial Catalog=DataName;User ID=sa;Password=PasswordName;"
    dbName = "DummyDatabase"
    
    ' Get the current date and time
    currentDate = Now
    
    ' Initialize ADO objects
    Set conn = CreateObject("ADODB.Connection")
    Set cmd = CreateObject("ADODB.Command")
    
    ' Open the connection
    conn.Open connString
    
    ' Check if the database already exists, and if so, drop it
    On Error Resume Next
    conn.Execute "DROP DATABASE " & dbName
    On Error GoTo 0
    
    ' Create a new database
    conn.Execute "CREATE DATABASE " & dbName
    
    ' Switch to the new database
    conn.Execute "USE " & dbName
    
    ' Create a table if it doesn't exist
    conn.Execute "CREATE TABLE DummyTable (ID INT PRIMARY KEY, EntryDate DATETIME)"
    
    ' Insert the current date and time into the table
    conn.Execute "INSERT INTO DummyTable (ID, EntryDate) VALUES (1, '" & Format(currentDate, "yyyy-mm-dd hh:mm:ss") & "')"
    
    ' Close the connection
    conn.Close
    
    ' Clean up objects
    Set cmd = Nothing
    Set conn = Nothing
    
    MsgBox "SQL dummy database created successfully!", vbInformation
End Sub


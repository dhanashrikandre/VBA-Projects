Attribute VB_Name = "CheckSQLData"
Sub CheckSQLDataTime()  'check current and macro run date difference
    Dim conn As Object
    Dim rs As Object
    Dim connString As String
    Dim storedDate As Date
    Dim currentDate As Date
    Dim timeDifference As Double
    
    dbName = "DummyDatabase"
    ' Replace these with your SQL Server details
    connString = "Provider=SQLOLEDB;Data Source=1.1.11.1,11;Initial Catalog=DataName;User ID=sa;Password=PasswordName;"
    
    ' Get the current date and time
    currentDate = Now
    
    ' Initialize ADO objects
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    
    ' Open the connection
    conn.Open connString
    
    
    ' Switch to the new database
    conn.Execute "USE " & dbName
    ' Retrieve the stored date and time from the table
    rs.Open "SELECT TOP 1 EntryDate FROM DummyTable ORDER BY ID DESC", conn
    If Not rs.EOF Then
        storedDate = rs.Fields("EntryDate").value
    Else
        ' Handle the case where there are no records in the table
        MsgBox "No records found in the table.", vbExclamation
        rs.Close
        conn.Close
        Exit Sub
    End If
    rs.Close
    
    ' Calculate the time difference in hours
    timeDifference = DateDiff("h", storedDate, currentDate)
    
    ' Close the connection
    conn.Close
    
    ' Check if the time difference is more than 2 hours
    If timeDifference > 2 Then
        MsgBox "Time difference is more than 2 hours! Error.", vbCritical
        Exit Sub
    Else
        
        Dim wb As Workbook
        Dim cws As Worksheet, iws As Worksheet
        Set wb = ActiveWorkbook
        Set cws = wb.Sheets("MySheet")
        Set iws = wb.Sheets("CheckSheet")
        
        cwsdate = Split(cws.Range("A2").value, ":")
        iwsdate = Split(iws.Range("C4").value, ":")
        If CDate(cwsdate(1)) = CDate(iwsdate(1)) Then
        Else
            Debug.Print (cwsdate(1))
            Debug.Print (iwsdate(1))
            MsgBox "Dates for MySheet and CheckSheet sheets do not match", vbCritical
            Exit Sub
        End If
        
        
    End If
End Sub

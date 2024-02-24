Attribute VB_Name = "Module1"
Sub GetSystemInfo()
    Dim computerName As String
    Dim ipAddress As String
    Dim systemNumber As String
    
    ' Retrieve computer name
    computerName = Environ("COMPUTERNAME")
        
' Retrieve the IP address using WMI
    wmiQuery = "SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True"
    Set wmiObj = GetObject("winmgmts:\\.\root\cimv2")
    For Each objItem In wmiObj.ExecQuery(wmiQuery)
        ipAddress = objItem.ipAddress(0)
    Next objItem
    ' Retrieve system number (if available)
    systemNumber = Environ("SYSTEMNUMBER")
    
    ' Print the retrieved information
    MsgBox "Computer Name: " & computerName & vbCrLf & _
           "IP Address: " & ipAddress & vbCrLf & _
           "System Number: " & systemNumber
End Sub


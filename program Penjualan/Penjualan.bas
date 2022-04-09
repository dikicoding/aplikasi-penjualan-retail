Attribute VB_Name = "ModDB"
Global Con As New ADODB.Connection

Function KonekDB()
Dim sCon As String

sCon = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & "localhost" & ";DATABASE=" & "penjualan" & ";UID=" & "root" & ";PWD=" & "" & ";PORT=" & "3306" & ";OPTION=3"

If Con.State = adStateOpen Then
    Con.Close
End If

Con.ConnectionString = sCon
Con.Open
End Function



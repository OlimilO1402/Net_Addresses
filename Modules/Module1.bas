Attribute VB_Name = "MWMI"
Option Explicit

Public Function MACAddresses() As String()
Try: On Error GoTo Catch
    Dim WMI As Object: Set WMI = GetObject("winmgmts:")
    If WMI Is Nothing Then
        MsgBox "Error could not create WMI"
        Exit Function
    End If
    Dim sSQL As String: sSQL = "SELECT MACAddress FROM Win32_NetworkAdapter WHERE ((MACAddress Is Not NULL) AND (Manufacturer <> 'Microsoft'))"
    Dim jobs As Object: Set jobs = WMI.ExecQuery(sSQL)
    ReDim sa(0) As String
    Dim mo As Object
    For Each mo In jobs
        ReDim sa(UBound(sa) + 1)
        sa(UBound(sa)) = mo.MACAddress
    Next
    MACAddresses = sa
    Exit Function
Catch:
    MsgBox "Error could not create WMI"
End Function



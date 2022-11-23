Attribute VB_Name = "Module1"
Option Explicit
Private Type IPAddr
    s_b1    As Byte
    s_b2    As Byte
    s_b3    As Byte
    s_b4    As Byte
End Type

Private Type IPAddrCompat
    ul      As Long
End Type

Private Type MacAddress
    s_b1    As Byte
    s_b2    As Byte
    s_b3    As Byte
    s_b4    As Byte
    s_b5    As Byte
    s_b6    As Byte
    unused  As Integer
End Type

'possible return types and Error-Codes if SendARP-return is not 0
Private Const ERROR_GEN_FAILURE         As Long = 31&
Private Const ERROR_NOT_SUPPORTED       As Long = 50&
Private Const ERROR_BAD_NET_NAME        As Long = 67&
Private Const ERROR_INVALID_PARAMETER   As Long = 87&
Private Const ERROR_BUFFER_OVERFLOW     As Long = 111&
Private Const ERROR_NOT_FOUND           As Long = 1168&
Private Const ERROR_INVALID_USER_BUFFER As Long = 1784&

'https://learn.microsoft.com/en-us/windows/win32/api/iphlpapi/nf-iphlpapi-sendarp
Private Declare Function SendARP Lib "Iphlpapi.dll" (ByVal DestIP As Long, ByVal SrcIP As Long, ByRef pMacAddr As MacAddress, ByRef PhyAddrLen As Long) As Long

Public Function GetMac(ByVal strIP As String) As String
    
    'Dim nIndex As Long
    'Dim vasLocalIP As Variant
    Dim strIps() As String: strIps = Split(strIP, ".")
    
    Dim uIPAddr  As IPAddr
    With uIPAddr
        .s_b1 = CByte(strIps(0))
        .s_b2 = CByte(strIps(1))
        .s_b3 = CByte(strIps(2))
        .s_b4 = CByte(strIps(3))
    End With
    
    Dim uIPAddrCompat As IPAddrCompat: LSet uIPAddrCompat = uIPAddr
    Dim nMacAddrLen   As Long: nMacAddrLen = 8
    Dim uMacAddr As MacAddress
    If SendARP(uIPAddrCompat.ul, 0&, uMacAddr, nMacAddrLen) = 0 Then
        If nMacAddrLen = 6 Then
            GetMac = MacAddrString(uMacAddr, nMacAddrLen)
        End If
    End If
    
End Function


Private Function MacAddrString(ByRef the_uMacAddr As MacAddress, ByVal the_nMacAddrLen) As String
    With the_uMacAddr
        MacAddrString = Hex2(.s_b1) & ":" & Hex2(.s_b2) & ":" & Hex2(.s_b3) & ":" & Hex2(.s_b4) & ":" & Hex2(.s_b5) & ":" & Hex2(.s_b6)
    End With
End Function

Private Function Hex2(ByVal the_byt As Byte) As String
    Hex2 = Hex$(the_byt): If Len(Hex2) = 1 Then Hex2 = "0" & Hex2
End Function

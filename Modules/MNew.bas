Attribute VB_Name = "MNew"
Option Explicit

Public Function IPAddress(ByVal i1 As Integer, ByVal i2 As Integer, ByVal i3 As Integer, ByVal i4 As Integer, Optional i5, Optional i6, Optional i7, Optional i8) As IPAddress
    'Set IPAddress = New IPAddress: IPAddress.New_ i1, i2, i3, i4, i5, i6, i7, i8
    If IsMissing(i5) Then
        Set IPAddress = IPAddressV4(CByte(i1), CByte(i2), CByte(i3), CByte(i4))
    Else
        Set IPAddress = IPAddressV6(i1, i2, i3, i4, i5, i6, i7, i8)
    End If
End Function

Public Function IPAddressV(StrLngBytesNewAddress) As IPAddress
    Set IPAddressV = New IPAddress: IPAddressV.New_ StrLngBytesNewAddress
End Function

Public Function IPAddressV4(ByVal b1 As Byte, ByVal b2 As Byte, ByVal b3 As Byte, ByVal b4 As Byte, Optional Port As Integer) As IPAddress
    Set IPAddressV4 = New IPAddress: IPAddressV4.NewV4 b1, b2, b3, b4, Port
End Function

Public Function IPAddressV4Rnd() As IPAddress
    Randomize
    Set IPAddressV4Rnd = New IPAddress: IPAddressV4Rnd.NewV4 RndUInt8, RndUInt8, RndUInt8, RndUInt8
End Function

Public Function IPAddressV6(ByVal i1 As Integer, ByVal i2 As Integer, ByVal i3 As Integer, ByVal i4 As Integer, ByVal i5 As Integer, ByVal i6 As Integer, ByVal i7 As Integer, ByVal i8 As Integer) As IPAddress
    Set IPAddressV6 = New IPAddress: IPAddressV6.NewV6 i1, i2, i3, i4, i5, i6, i7, i8
End Function

Public Function IPAddressV6Rnd() As IPAddress
    Randomize
    Set IPAddressV6Rnd = New IPAddress: IPAddressV6Rnd.NewV6 RndUInt16, RndUInt16, RndUInt16, RndUInt16, RndUInt16, RndUInt16, RndUInt16, RndUInt16
End Function

Public Function MACAddress(ByVal b1 As Byte, ByVal b2 As Byte, ByVal b3 As Byte, ByVal b4 As Byte, ByVal b5 As Byte, ByVal b6 As Byte, Optional sep As String = "-") As MACAddress
    Set MACAddress = New MACAddress: MACAddress.New_ b1, b2, b3, b4, b5, b6, sep
End Function

Public Function MACAddressA(bytes05() As Byte, Optional sep As String = "-") As MACAddress
    Set MACAddress = New MACAddress: MACAddress.New_ bytes(0), bytes(1), bytes(2), bytes(3), bytes(4), bytes(5), sep
End Function


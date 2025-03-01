VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'https://de.wikipedia.org/wiki/IP-Adresse
'https://en.wikipedia.org/wiki/IP_address
Private Type TInt
    Value As Integer
End Type
Private Type TLng
    Value As Long
End Type
Private Type TCur
    Value As Currency
End Type
Private Type TByte2
    b1 As Byte
    b2 As Byte
End Type
Private Type TByte4
    b1 As Byte
    b2 As Byte
    b3 As Byte
    b4 As Byte
End Type
Private Type TInt2
    i1 As Integer
    i2 As Integer
End Type
Private Type BArr4
    Value(0 To 3) As Byte
End Type
Private Type IArr8
    Value(0 To 7) As Integer
End Type
Private Type TNetAddress
    Value1_b12 As TByte2
    Value2_b34 As TByte2
    Value3 As Integer
    Value4 As Integer
    Value5 As Integer
    Value6 As Integer
    Value7 As Integer
    Value8 As Integer
End Type
Private m_ip      As TNetAddress
Private m_Version As Byte
Private m_Port       'As Long 'by default it's missing

'IPAddress: 192.168.178.1 'unsinged-Byte1.unsinged-Byte2.unsinged-Byte3.unsinged-Byte4

'             ui21:ui22:ui23:ui24:ui25:ui26:ui27:ui28
'IPAddressV6: 1234:5678:90AB:CDEF:1234:5678:90AB:CDEF: UInt16_1:UInt16_2:UInt16_3:UInt16_4:UInt16_5:UInt16_6:UInt16_7:UInt16_8
'MACAddress:   12:34:56:78:AB:CD

Friend Sub New_(StrLngBytesNewAddress) ', Optional aName As String)
    Parse StrLngBytesNewAddress
End Sub

Friend Sub NewV4(ByVal b1 As Byte, ByVal b2 As Byte, ByVal b3 As Byte, ByVal b4 As Byte, Optional Port As Integer)
    With Me: .ValueB1 = b1: End With
    With m_ip: .Value1_b12.b2 = b2: .Value2_b34.b1 = b3: .Value2_b34.b2 = b4: End With
End Sub

Friend Sub NewV6(ByVal i1 As Integer, ByVal i2 As Integer, ByVal i3 As Integer, ByVal i4 As Integer, ByVal i5 As Integer, ByVal i6 As Integer, ByVal i7 As Integer, ByVal i8 As Integer)
    With Me: .ValueI1 = i1: .ValueI2 = i2: End With
    With m_ip: .Value3 = i3: .Value4 = i4: .Value5 = i5: .Value6 = i6: .Value7 = i7: .Value8 = i8: End With
End Sub
'
'Friend Sub NewD(ByVal d1 As Integer, ByVal d2 As Integer, ByVal d3 As Integer, ByVal d4 As Integer, Optional d5, Optional d6, Optional d7, Optional d8)
'    With m_Data: .Data1 = d1: .Data2 = d2: .Data3 = d3: .Data4 = d4: .Data5 = d5: .Data6 = d6: .Data7 = d7: .Data8 = d8: End With
'End Sub

Friend Sub NewC(other As IPAddress)
    With other
        m_Version = .Version
        If m_Version = 4 Then
            NewV4 .ValueB1, .ValueB2, .ValueB3, .ValueB4
        Else
            NewV6 .ValueI1, .ValueI2, .ValueI3, .ValueI4, .ValueI5, .ValueI6, .ValueI7, .ValueI8
        End If
    End With
    'Debug.Print other.IPToStr
    'Me.LAddress = other.LAddress
End Sub

Friend Function Clone() As IPAddress
    Set Clone = New IPAddress: Clone.NewC Me
End Function

'Copy to MNew:
'Public Function IPAddress(ByVal i1 As Integer, ByVal i2 As Integer, ByVal i3 As Integer, ByVal i4 As Integer, Optional i5, Optional i6, Optional i7, Optional i8) As IPAddress
'    'Set IPAddress = New IPAddress: IPAddress.New_ i1, i2, i3, i4, i5, i6, i7, i8
'    If IsMissing(i5) Then
'        Set IPAddress = IPAddressV4(CByte(i1), CByte(i2), CByte(i3), CByte(i4))
'    Else
'        Set IPAddress = IPAddressV6(i1, i2, i3, i4, i5, i6, i7, i8)
'    End If
'End Function
'
'Public Function IPAddressV(StrLngBytesNewAddress) As IPAddress
'    Set IPAddressV = New IPAddress: IPAddressV.New_ StrLngBytesNewAddress
'End Function
'
'Public Function IPAddressV4(ByVal b1 As Byte, ByVal b2 As Byte, ByVal b3 As Byte, ByVal b4 As Byte, Optional Port As Integer) As IPAddress
'    Set IPAddressV4 = New IPAddress: IPAddressV4.NewV4 b1, b2, b3, b4, Port
'End Function
'
'Public Function IPAddressV6(ByVal i1 As Integer, ByVal i2 As Integer, ByVal i3 As Integer, ByVal i4 As Integer, ByVal i5 As Integer, ByVal i6 As Integer, ByVal i7 As Integer, ByVal i8 As Integer) As IPAddress
'    Set IPAddressV6 = New IPAddress: IPAddressV6.NewV6 i1, i2, i3, i4, i5, i6, i7, i8
'End Function

'Friend Function IPToStr() As String
'    With m_ip
'        IPToStr = CStr(.b4) & "." & CStr(.b3) & "." & CStr(.b2) & "." & CStr(.b1)
'        If Not IsEmpty(m_Port) Then IPToStr = IPToStr & ":" & CStr(CLng(m_Port))
'    End With
'End Function

Public Function Parse(StrLngBytesNewAddress) As Boolean
    'StrLngBytesNewAddress: you can either give a string, or a bytearray or a long, or a currency
    'm_Name = aName
    Dim vt As VbVarType: vt = VarType(StrLngBytesNewAddress)
    Dim va, i As Long, u As Long
    Select Case vt
    Case vbCurrency
        Dim cc As TCur: cc.Value = CCur(StrLngBytesNewAddress) / CCur(10000): LSet m_ip = cc
        Exit Function
    Case vbLong
        Dim ll As TLng: ll.Value = CLng(StrLngBytesNewAddress): LSet m_ip = ll
        Exit Function
    Case (vbArray Or vbByte)
        va = StrLngBytesNewAddress
        u = UBound(va)
        m_Version = 4
    Case (vbArray Or vbInteger)
        va = StrLngBytesNewAddress
        u = UBound(va)
        m_Version = 6
    Case (vbArray Or vbVariant)
        va = StrLngBytesNewAddress
        u = UBound(va)
        m_Version = IIf(u < 4, 4, 6)
    Case vbString
        va = Split(StrLngBytesNewAddress, ".")
        u = UBound(va)
        If u = 3 Then
            m_Version = 4
        Else
            va = Split(StrLngBytesNewAddress, ":")
            u = UBound(va)
            m_Version = 6
            For i = 0 To u
                va(i) = "&H" & va(i)
            Next
        End If
    End Select
    Dim s As String, sa() As String
    If m_Version = 4 Then
        Dim ba As BArr4
        For i = 0 To u
            If i < u Then
                ba.Value(i) = CByte(va(i))
            Else
                'maybe the last contains a port
                s = va(i)
                If InStr(s, ":") = 0 Then
                    If IsNumeric(s) Then
                        ba.Value(i) = CByte(s)
                    End If
                Else
                    sa = Split(s, ":")
                    ba.Value(i) = CByte(sa(0))
                    m_Port = CLng(sa(1))
                End If
            End If
        Next
        LSet m_ip = ba
    Else
        Dim ia As IArr8
        For i = 0 To u
            If i < u Then
                ia.Value(i) = CInt(va(i))
            Else
                'maybe the last contains a port
                s = va(i)
                If InStr(s, "%") = 0 Then
                    If IsNumeric(s) Then
                        ia.Value(i) = CInt(s)
                    End If
                Else
                    sa = Split(s, "%")
                    ia.Value(i) = CInt(sa(0))
                    m_Port = CLng(sa(1))
                End If
            End If
        Next
        LSet m_ip = ia
    End If
End Function

Private Function Hex4(ByVal Value As Integer) As String
    Hex4 = Hex(Value): If Len(Hex4) < 4 Then Hex4 = String(4 - Len(Hex4), "0") & Hex4
End Function

Friend Function ToStr() As String
    'ToStr = IPToStr & IIf(Len(m_Name), " " & m_Name, "")
    Dim sep As String, i As Long, u As Long
    If m_Version = 4 Then
        sep = ".": u = 3
    Else
        sep = ":": u = 7
    End If
    ReDim sa(0 To u) As String
    'For i = 0 To u
    If m_Version = 4 Then
        sa(i) = CStr(m_ip.Value1_b12.b1): i = i + 1
        sa(i) = CStr(m_ip.Value1_b12.b2): i = i + 1
        sa(i) = CStr(m_ip.Value2_b34.b1): i = i + 1
        sa(i) = CStr(m_ip.Value2_b34.b2) ': i = i + 1
    Else
        Dim ii As TInt
        LSet ii = m_ip.Value1_b12
        sa(i) = Hex4(ii.Value):    i = i + 1
        LSet ii = m_ip.Value2_b34
        sa(i) = Hex4(ii.Value):    i = i + 1
        sa(i) = Hex4(m_ip.Value3): i = i + 1
        sa(i) = Hex4(m_ip.Value4): i = i + 1
        sa(i) = Hex4(m_ip.Value5): i = i + 1
        sa(i) = Hex4(m_ip.Value6): i = i + 1
        sa(i) = Hex4(m_ip.Value7): i = i + 1
        sa(i) = Hex4(m_ip.Value8) ': i = i + 1
    End If
'    Set p = IIf(m_Version = 4, ".", ":")
'    With m_ip
'        s = s & .Value1_b12.b1 & sep
'    End With
'    If m_Version = 4 Then sep = "." Else sep = ":"
    Dim portsep As String: portsep = IIf(m_Version = 4, ":", "%")
    ToStr = Join(sa, sep) & IIf(Not IsEmpty(m_Port), portsep & CStr(m_Port), "")
End Function

'Friend Function ToInfoStr() As String
''    ToInfoStr = Me.Name & vbCrLf & String(Len(m_Name), "=") & vbCrLf & _
''                IIf(Len(m_Server), "Server: " & Me.Server & vbCrLf, "") & _
''                Me.NsluResult & vbCrLf & _
''                Me.PingResult & vbCrLf & _
''                Me.StatResult
'End Function

Public Property Get Ptr() As LongPtr 'we could also remove the datatype and return only a variant instead, if maybe u do not need MPtr
    Ptr = VarPtr(m_ip)
End Property
Public Property Get Version() As Byte
    Version = m_Version
End Property

Public Property Get AddressL() As Long
    ' until now this works only for Version 4
    If m_Version = 4 Then
        Dim ll As TLng: LSet ll = m_ip: AddressL = ll.Value
    'Else
    End If
End Property
Public Property Let AddressL(ByVal Value As Long)
    Dim ll As TLng: ll.Value = Value
    LSet m_ip = ll
    m_Version = 4
End Property

Public Property Get Address() As Currency
    ' until now this works only for Version 4
    If m_Version = 4 Then
        Dim cc As TCur: LSet cc = m_ip: Address = cc.Value * 10000
    'Else
    End If
End Property
Public Property Let Address(ByVal Value As Currency)
    ' until now this works only for Version 4
    Dim cc As TCur: cc.Value = Value / 10000
    LSet m_ip = cc
    m_Version = 4
End Property

'Public Property Get Name() As String
'    Name = m_Name
'End Property
'Friend Property Let Name(Value As String)
'    m_Name = Value
'End Property

Public Property Get Port()
    Port = m_Port
End Property
Public Property Let Port(Value)
    m_Port = Value
End Property

'Public Property Get Server() As String
'    Server = m_Server
'End Property
'
'Public Property Get IsValid() As Boolean
'    IsValid = 0 < InStr(1, m_PingResult, "Antwort")
'    If Not IsValid Then IsValid = 0 < Len(m_Name)
'End Property
'Public Property Get Compare(other As IPAddress) As Integer
'    If other.b1 < Me.b1 Then Compare = 1:  Exit Property
'    If Me.b1 < other.b1 Then Compare = -1: Exit Property
'    If other.b2 < Me.b2 Then Compare = 1:  Exit Property
'    If Me.b2 < other.b2 Then Compare = -1: Exit Property
'    If other.b3 < Me.b3 Then Compare = 1:  Exit Property
'    If Me.b3 < other.b3 Then Compare = -1: Exit Property
'    If other.b4 < Me.b4 Then Compare = 1:  Exit Property
'    If Me.b4 < other.b4 Then Compare = -1: Exit Property
'End Property

Public Function Compare(other As IPAddress) As Integer
    If m_Version = 4 Then
        If other.ValueB1 < Me.ValueB1 Then Compare = 1:  Exit Function
        If Me.ValueB1 < other.ValueB1 Then Compare = -1: Exit Function
        If other.ValueB2 < Me.ValueB2 Then Compare = 1:  Exit Function
        If Me.ValueB2 < other.ValueB2 Then Compare = -1: Exit Function
        If other.ValueB3 < Me.ValueB3 Then Compare = 1:  Exit Function
        If Me.ValueB3 < other.ValueB3 Then Compare = -1: Exit Function
        If other.ValueB4 < Me.ValueB4 Then Compare = 1:  Exit Function
        If Me.ValueB4 < other.ValueB4 Then Compare = -1: Exit Function
    Else
        If other.ValueI1 < Me.ValueI1 Then Compare = 1:  Exit Function
        If Me.ValueI1 < other.ValueI1 Then Compare = -1: Exit Function
        If other.ValueI2 < Me.ValueI2 Then Compare = 1:  Exit Function
        If Me.ValueI2 < other.ValueI2 Then Compare = -1: Exit Function
        If other.ValueI3 < Me.ValueI3 Then Compare = 1:  Exit Function
        If Me.ValueI3 < other.ValueI3 Then Compare = -1: Exit Function
        If other.ValueI4 < Me.ValueI4 Then Compare = 1:  Exit Function
        If Me.ValueI4 < other.ValueI4 Then Compare = -1: Exit Function
        
        If other.ValueI5 < Me.ValueI5 Then Compare = 1:  Exit Function
        If Me.ValueI5 < other.ValueI5 Then Compare = -1: Exit Function
        If other.ValueI6 < Me.ValueI6 Then Compare = 1:  Exit Function
        If Me.ValueI6 < other.ValueI6 Then Compare = -1: Exit Function
        If other.ValueI7 < Me.ValueI7 Then Compare = 1:  Exit Function
        If Me.ValueI7 < other.ValueI7 Then Compare = -1: Exit Function
        If other.ValueI8 < Me.ValueI8 Then Compare = 1:  Exit Function
        If Me.ValueI8 < other.ValueI8 Then Compare = -1: Exit Function
    End If
End Function

Public Sub OneUp()
    Add 1
'    If m_Version = 4 Then
'        Dim bb As TByte4: LSet bb = m_ip
'        SwapByteOrderB4 bb
'        Dim cc As TCur:   LSet cc = bb
'        cc.Value = cc.Value + CCur(0.0001)
'        LSet bb = cc
'        SwapByteOrderB4 bb
'        LSet m_ip = bb
'    'Else
'    End If
End Sub

Public Sub Add(ByVal Value As Long)
    If m_Version = 4 Then
        Dim bb As TByte4: LSet bb = m_ip
        SwapByteOrderB4 bb
        Dim cc As TCur:   LSet cc = bb
        cc.Value = cc.Value + CCur(CCur(Value) / 10000)
        LSet bb = cc
        SwapByteOrderB4 bb
        LSet m_ip = bb
    'Else
    End If
End Sub

Private Sub SwapByteOrderB4(bb As TByte4)
    Dim tmp As Byte
    tmp = bb.b1: bb.b1 = bb.b4: bb.b4 = tmp
    tmp = bb.b2: bb.b2 = bb.b3: bb.b3 = tmp
End Sub

'Public Sub OneUp()
'    If m_ip.b1 < 255 Then
'        m_ip.b1 = m_ip.b1 + 1
'    ElseIf m_ip.b2 < 255 Then
'        m_ip.b1 = 0
'        m_ip.b2 = m_ip.b2 + 1
'    ElseIf m_ip.b3 < 255 Then
'        m_ip.b1 = 0: m_ip.b2 = 0
'        m_ip.b3 = m_ip.b3 + 1
'    ElseIf m_ip.b4 < 255 Then
'        m_ip.b1 = 0: m_ip.b2 = 0: m_ip.b3 = 0
'        m_ip.b4 = m_ip.b4 + 1
'    Else
'        m_ip.b1 = 0: m_ip.b2 = 0: m_ip.b3 = 0: m_ip.b4 = 0
'    End If
'End Sub
'
'Public Sub Add(ByVal Sum As Currency)
'    Dim cdSrc As BCur4: cdSrc.bytes = Sum / 10000
'    Dim curIP As BCur4: LSet curIP = m_ip
'    curIP.bytes = curIP.bytes + cdSrc.bytes
'    LSet m_ip = curIP
'End Sub

'Public Property Get PingResult() As String
'    PingResult = m_PingResult
'End Property
'Public Property Get StatResult() As String
'    StatResult = m_StatResult
'End Property
'Public Property Get NsluResult() As String
'    NsluResult = m_NsluResult
'End Property

' v ############################## v '    Public Property IPAddressV4    ' v ############################## v '
Public Property Get ValueB1() As Byte
    ValueB1 = m_ip.Value1_b12.b1
End Property
Public Property Let ValueB1(ByVal Value As Byte)
    m_ip.Value1_b12.b1 = Value
    m_Version = 4
End Property

Public Property Get ValueB2() As Byte
    ValueB2 = m_ip.Value1_b12.b2
End Property
Public Property Let ValueB2(ByVal Value As Byte)
    m_ip.Value1_b12.b2 = Value
    m_Version = 4
End Property

Public Property Get ValueB3() As Byte
    ValueB3 = m_ip.Value2_b34.b1
End Property
Public Property Let ValueB3(ByVal Value As Byte)
    m_ip.Value2_b34.b1 = Value
    m_Version = 4
End Property

Public Property Get ValueB4() As Byte
    ValueB4 = m_ip.Value2_b34.b2
End Property
Public Property Let ValueB4(ByVal Value As Byte)
    m_ip.Value2_b34.b2 = Value
    m_Version = 4
End Property
' ^ ############################## ^ '    Public Property IPAddressV4    ' ^ ############################## ^ '

' v ############################## v '    Public Property IPAddressV6    ' v ############################## v '
Public Property Get ValueI1() As Integer
    Dim ii As TInt: LSet ii = m_ip.Value1_b12
    ValueI1 = ii.Value
End Property
Public Property Let ValueI1(ByVal Value As Integer)
    Dim ii As TInt: ii.Value = Value
    LSet m_ip.Value1_b12 = ii
    m_Version = 6
End Property

Public Property Get ValueI2() As Integer
    Dim ii As TInt: LSet ii = m_ip.Value2_b34
    ValueI2 = ii.Value
End Property
Public Property Let ValueI2(ByVal Value As Integer)
    Dim ii As TInt: ii.Value = Value
    LSet m_ip.Value2_b34 = ii
    m_Version = 6
End Property

Public Property Get ValueI3() As Integer
    ValueI3 = m_ip.Value3
End Property
Public Property Let ValueI3(ByVal Value As Integer)
    m_ip.Value3 = Value
    m_Version = 6
End Property

Public Property Get ValueI4() As Integer
    ValueI4 = m_ip.Value4
End Property
Public Property Let ValueI4(ByVal Value As Integer)
    m_ip.Value4 = Value
    m_Version = 6
End Property

Public Property Get ValueI5() As Integer
    ValueI5 = m_ip.Value5
End Property
Public Property Let ValueI5(ByVal Value As Integer)
    m_ip.Value5 = Value
    m_Version = 6
End Property

Public Property Get ValueI6() As Integer
    ValueI6 = m_ip.Value6
End Property
Public Property Let ValueI6(ByVal Value As Integer)
    m_ip.Value6 = Value
    m_Version = 6
End Property

Public Property Get ValueI7() As Integer
    ValueI7 = m_ip.Value7
End Property
Public Property Let ValueI7(ByVal Value As Integer)
    m_ip.Value7 = Value
    m_Version = 6
End Property

Public Property Get ValueI8() As Integer
    ValueI8 = m_ip.Value8
End Property
Public Property Let ValueI8(ByVal Value As Integer)
    m_ip.Value8 = Value
    m_Version = 6
End Property
' ^ ############################## ^ '    Public Property IPAddressV6    ' ^ ############################## ^ '


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MACAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'https://de.wikipedia.org/wiki/MAC-Adresse
'https://en.wikipedia.org/wiki/MAC_address

Private Type TByte6
    Value1 As Byte
    Value2 As Byte
    Value3 As Byte
    Value4 As Byte
    Value5 As Byte
    Value6 As Byte
End Type

Private m_mac As TByte6
Private m_sep As String
    
Private Sub Class_Initialize()
    m_sep = "-"
End Sub

Friend Sub New_(ByVal b1 As Byte, ByVal b2 As Byte, ByVal b3 As Byte, ByVal b4 As Byte, ByVal b5 As Byte, ByVal b6 As Byte, Optional sep As String = "-")
    With m_mac
        .Value1 = b1
        .Value2 = b2
        .Value3 = b3
        .Value4 = b4
        .Value5 = b5
        .Value6 = b6
    End With
    m_sep = sep
End Sub

Friend Sub NewS(s As String)
    Parse s
End Sub

Friend Sub NewC(other As MACADDress)
    With other: Me.New_ .Value1, .Value2, .Value3, .Value4, .Value5, .Value6: End With
End Sub

Public Function Parse(s As String) As Boolean
Try: On Error GoTo Catch
    Dim sa() As String
    If Not TryOneOfSeparators(s, "-:", sa) Then Exit Function
    Dim i As Long
    
    Parse = True
    Exit Function
Catch:
End Function

Private Function TrySeparator(ByVal Value As String, sep As String, sa_out() As String) As Boolean
Try: On Error GoTo Catch
    sa_out = Split(Value, sep)
    TrySeparator = 0 < UBound(sa_out)
Catch:
End Function

Private Function TryOneOfSeparators(ByVal Value As String, multiseps As String, sa_out() As String) As Boolean
    Dim i As Long, sep As String
    For i = 1 To Len(multiseps)
        sep = Mid(multiseps, i, 1)
        If TrySeparator(Value, sep, sa_out) Then
            TryOneOfSeparators = True
            Exit Function
        End If
    Next
End Function


Friend Function Clone() As MACADDress
    Set Clone = New MACADDress: Clone.NewC Me
End Function

Friend Property Get Ptr() As LongPtr
    Ptr = VarPtr(m_mac)
End Property

Friend Function ToStr() As String
    Dim sep As String: sep = m_sep
    With m_mac
        ToStr = Hex2(.Value1) & sep & Hex2(.Value2) & sep & Hex2(.Value3) & sep & Hex2(.Value4) & sep & Hex2(.Value5) & sep & Hex2(.Value6)
    End With
End Function

Public Property Get Separator() As String
    Separator = m_sep
End Property
Public Property Let Separator(ByVal Value As String)
    m_sep = Value
End Property

Public Property Get Value1() As Byte
    Value1 = m_mac.Value1
End Property
Public Property Let Value1(ByVal Value As Byte)
    m_mac.Value1 = Value
End Property

Public Property Get Value2() As Byte
    Value2 = m_mac.Value2
End Property
Public Property Let Value2(ByVal Value As Byte)
    m_mac.Value2 = Value
End Property

Public Property Get Value3() As Byte
    Value3 = m_mac.Value3
End Property
Public Property Let Value3(ByVal Value As Byte)
    m_mac.Value3 = Value
End Property

Public Property Get Value4() As Byte
    Value4 = m_mac.Value4
End Property
Public Property Let Value4(ByVal Value As Byte)
    m_mac.Value4 = Value
End Property

Public Property Get Value5() As Byte
    Value5 = m_mac.Value5
End Property
Public Property Let Value5(ByVal Value As Byte)
    m_mac.Value5 = Value
End Property

Public Property Get Value6() As Byte
    Value6 = m_mac.Value6
End Property
Public Property Let Value6(ByVal Value As Byte)
    m_mac.Value6 = Value
End Property

Private Function Hex2(ByVal Value As Byte) As String
    Hex2 = Hex(Value): If Len(Hex2) < 2 Then Hex2 = "0" & Hex2
End Function


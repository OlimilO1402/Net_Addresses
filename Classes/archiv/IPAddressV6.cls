VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPAddressV6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Short8
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
    Data4 As Integer
    Data5 As Integer
    Data6 As Integer
    Data7 As Integer
    Data8 As Integer
End Type

Private m_Data As Short8

Friend Sub New_(ByVal d1 As Integer, ByVal d2 As Integer, ByVal d3 As Integer, ByVal d4 As Integer, ByVal d5 As Integer, ByVal d6 As Integer, ByVal d7 As Integer, ByVal d8 As Integer)
    With m_Data: .Data1 = d1: .Data2 = d2: .Data3 = d3: .Data4 = d4: .Data5 = d5: .Data6 = d6: .Data7 = d7: .Data8 = d8: End With
End Sub

Friend Sub NewC(other As IPAddressV6)
    With m_Data: .Data1 = other.Data1: .Data2 = other.Data2: .Data3 = other.Data3: .Data4 = other.Data4: .Data5 = other.Data5: .Data6 = other.Data6: .Data7 = other.Data7: .Data8 = other.Data8: End With
End Sub

Friend Function Clone() As IPAddressV6
    Set Clone = New IPAddressV6: Clone.NewC Me
End Function

Friend Function ToStr() As String
    With m_Data
        ToStr = Hex4(.Data1) & ":" & Hex4(.Data2) & ":" & Hex4(.Data3) & ":" & Hex4(.Data4) & ":" & Hex4(.Data5) & ":" & Hex4(.Data6) & ":" & Hex4(.Data7) & ":" & Hex4(.Data8)
    End With
End Function

Private Function Hex4(ByVal Value As Integer) As String
    Hex4 = Hex(Value): If Len(Hex4) < 4 Then Hex4 = String(4 - Len(Hex4), "0") & Hex4
End Function

Public Property Get Data1() As Integer
    Data1 = m_Data.Data1
End Property
'Public Property Let Data1(ByVal Value As Integer)
'    m_Data.Data1 = Value
'End Property

Public Property Get Data2() As Integer
    Data2 = m_Data.Data2
End Property
'Public Property Let Data2(ByVal Value As Integer)
'    m_Data.Data2 = Value
'End Property

Public Property Get Data3() As Integer
    Data3 = m_Data.Data3
End Property
'Public Property Let Data3(ByVal Value As Integer)
'    m_Data.Data3 = Value
'End Property

Public Property Get Data4() As Integer
    Data4 = m_Data.Data4
End Property
'Public Property Let Data4(ByVal Value As Integer)
'    m_Data.Data4 = Value
'End Property

Public Property Get Data5() As Integer
    Data5 = m_Data.Data5
End Property
'Public Property Let Data5(ByVal Value As Integer)
'    m_Data.Data5 = Value
'End Property

Public Property Get Data6() As Integer
    Data6 = m_Data.Data6
End Property
'Public Property Let Data6(ByVal Value As Integer)
'    m_Data.Data6 = Value
'End Property

Public Property Get Data7() As Integer
    Data7 = m_Data.Data7
End Property
'Public Property Let Data7(ByVal Value As Integer)
'    m_Data.Data7 = Value
'End Property

Public Property Get Data8() As Integer
    Data8 = m_Data.Data8
End Property
'Public Property Let Data8(ByVal Value As Integer)
'    m_Data.Data8 = Value
'End Property


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPAddresses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_List As Collection 'Of IPAddressV4

Private Sub Class_Initialize()
    Set m_List = New Collection
End Sub

Public Function Add(aIPV4 As IPAddressV4) As IPAddressV4
    m_List.Add aIPV4, aIPV4.IPToStr
End Function

Public Sub Clear()
    Class_Initialize
End Sub

Public Function Contains(aIPV4 As String) As Boolean
Try: On Error GoTo Catch
    Dim bIPV4 As IPAddressV4
    Set bIPV4 = m_List.Item(aIPV4)
    Contains = Not bIPV4 Is Nothing
    Exit Function
Catch:
End Function

Public Property Get Count() As Long
    Count = m_List.Count
End Property

Public Property Get Item(aIPV4 As String) As IPAddressV4
    Set Item = m_List.Item(aIPV4)
End Property

Public Property Get ItemI(Index As Long) As IPAddressV4
    Set ItemI = m_List.Item(Index)
End Property


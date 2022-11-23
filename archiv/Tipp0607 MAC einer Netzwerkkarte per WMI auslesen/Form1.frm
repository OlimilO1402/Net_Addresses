VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "www.ActiveVB.de"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source!

Option Explicit

Private Sub Form_Load()
    Dim MACs() As String
    Dim i As Long
    
    MACs = MACAddressWMI
    For i = 0 To UBound(MACs)
        If Len(MACs(i)) > 0 Then List1.AddItem MACs(i)
    Next i
End Sub

Public Function MACAddressWMI() As String()
Try: On Error GoTo Catch
    
    Dim WMIobj As Object: Set WMIobj = GetObject("winmgmts:")
    WMIobj.ExecQuery "SELECT MACAddress FROM Win32_NetworkAdapter WHERE ((MACAddress Is Not NULL) AND (Manufacturer <> 'Microsoft'))"
    
    ReDim sa(0) As String
    Dim MACobj As Object
    For Each MACobj In WMIobj
        ReDim sa(UBound(sa) + 1)
        sa(UBound(sa)) = MACobj.MACAddress
    Next
    
    MACAddressWMI = sa
    Exit Function
Catch:
    MsgBox "Fehler! WMI ist nicht vorhanden"
End Function


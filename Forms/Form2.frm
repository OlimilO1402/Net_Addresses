VERSION 5.00
Begin VB.Form FMACAddr 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox List1 
      Height          =   1845
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "FMACAddr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_LANAdapters As Collection

Private Sub Form_Load()
    Set m_LANAdapters = New Collection
Try: On Error GoTo Catch
    Dim sa() As String: sa = MWMI.MACAddresses
    Dim i As Long, s As String
    Dim la As LANAdapter
    For i = 0 To UBound(sa)
        s = sa(i)
        'Set la = MNew.LANAdapter(i, MNew.MACAddressA(s))
        List1.AddItem s 'la.MACADDress.ToStr
    Next
    Exit Sub
Catch: 'MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub Command1_Click()
    Dim mac As MACADDress
    Set mac = MNew.MACADDress(1, 2, 3, 4, 5, 6)
    MsgBox mac.ToStr
End Sub

VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox List1 
      Height          =   2010
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
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim sa() As String: sa = MWMI.MACAddresses
    Dim s
    For Each s In sa
        List1.AddItem
    Next
End Sub

Private Sub Command1_Click()
    Dim mac As MACAddress
    Set mac = MNew.MACAddress(1, 2, 3, 4, 5, 6)
    MsgBox mac.ToStr
End Sub

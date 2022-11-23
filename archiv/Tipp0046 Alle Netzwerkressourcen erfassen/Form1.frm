VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "www.ActiveVB.de"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox Liste 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4260
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   4320
      Width           =   3855
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
'Ansonsten viel Spaß und Erfolg mit diesem Source !

'Update by MGalpha
'WWW.grcs.de am 17.08.03
'eine Erklärung finden sie unter Module1

Option Explicit

Private FormAendern As Long

Private Sub Command1_Click()
    Label1.Caption = "Einen Moment bitte ..."
    Liste.Clear
    Netsuche1
    
    Label1.Caption = "Fertig  ..."
End Sub

Private Sub Form_Load()
    FormAendern = Form1.Width - Liste.Width
End Sub

Private Sub Form_Resize()
    Liste.Width = Form1.Width - FormAendern
End Sub

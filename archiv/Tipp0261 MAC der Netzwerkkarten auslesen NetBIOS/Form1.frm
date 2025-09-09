VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "ActiveVB Tipp - MAC-Adresse(n) auslesen"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox txtComputername 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Text            =   "Rechnername oder IP-Adresse"
      Top             =   1320
      Width           =   3015
   End
   Begin VB.OptionButton optOtherComputer 
      Caption         =   "Anderer Rechner"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.OptionButton optThisComputer 
      Caption         =   "Eigener Rechner"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   480
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Auslesen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   1800
      Width           =   3015
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0002
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "MAC-Adresse jedes gefundenen Netzwerkadapters"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3255
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

Option Explicit

Private Sub Command1_Click()
    Dim bLanAdapter() As Byte
    Dim i As Long
    Dim macaddr As String
    
    Me.MousePointer = vbHourglass
    
    ' Listbox leeren
    'List1.Clear
    
    'Alle Adapter auslesen
    Dim numAdapter As Long: numAdapter = MNetBios.EnumLanAdapter(bLanAdapter)
    
    'wurde mindestens ein aktiver Adapter gefunden
    If numAdapter > 0 Then
        'Für jeden Adapter die MAC-Adresse auslesen
        For i = 1 To numAdapter
            'diesen Adapter initalisieren
            MNetBios.ResetAdapter bLanAdapter(i), 20, 30
            
            If optThisComputer.Value = True Then
                ' MAC-Adresse dieses Adapters auslesen
                List1.AddItem MNetBios.GetMACAddress(bLanAdapter(i))
            Else
                ' Probieren die MAC-Adresse über diesen Adapter zu ermitteln
                macaddr = MNetBios.GetMACAddress(bLanAdapter(i), txtComputername.Text)
                ' Wenn eine MAC-Adresse über diesen Adapter ermittelt wurde,
                ' dann die MAC-Adresse anzeigen
                If Len(macaddr) > 0 Then List1.AddItem macaddr
            End If
        Next i
    Else
        List1.AddItem "Keine Netzwerkadapter gefunden!"
    End If
    
    Me.MousePointer = vbDefault
End Sub


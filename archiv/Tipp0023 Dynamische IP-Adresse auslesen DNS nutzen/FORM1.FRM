VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "www.ActiveVB.de"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command6 
      Caption         =   "Beenden"
      Height          =   375
      Left            =   7080
      TabIndex        =   17
      Top             =   3840
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ermittlung: Eigene Daten"
      Height          =   2295
      Left            =   0
      TabIndex        =   9
      Top             =   1440
      Width           =   7935
      Begin VB.Timer Timer1 
         Left            =   120
         Top             =   1200
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Auflegen"
         Height          =   375
         Left            =   1200
         TabIndex        =   15
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Einwählen"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   975
      End
      Begin VB.ListBox List1 
         Height          =   1425
         ItemData        =   "FORM1.frx":0000
         Left            =   2160
         List            =   "FORM1.frx":0002
         TabIndex        =   13
         Top             =   720
         Width           =   5655
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2160
         TabIndex        =   12
         Top             =   360
         Width           =   5655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Eigene Adresse ermitteln"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Dynamische IP-Adressen"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Abfrage: IP/Adresse"
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3960
         TabIndex        =   8
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   6360
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "IP->Name"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Name->Ip"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2160
         TabIndex        =   2
         Text            =   "193.7.255.22"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2160
         TabIndex        =   1
         Text            =   "www.altavista.de"
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label Label2 
         Caption         =   "="
         Height          =   135
         Left            =   6120
         TabIndex        =   6
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "="
         Height          =   135
         Left            =   3720
         TabIndex        =   5
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Label3"
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   4320
      Width           =   7935
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

Private Declare Function WSAGetLastError Lib "wsock32" () As Long
Private Declare Function WSAStartup Lib "wsock32" (ByVal wVersionRequired As Long, lpWSAData As WinSocketDataType) As Long
Private Declare Function WSACleanup Lib "wsock32" () As Long
Private Declare Function gethostname Lib "wsock32" (ByVal HostName As String, ByVal HostLen As Integer) As Long
Private Declare Function gethostbyname Lib "wsock32" (ByVal HostName As String) As Long
Private Declare Function gethostbyaddr Lib "wsock32" (ByVal addr As String, ByVal laenge As Integer, ByVal typ As Integer) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Const WS_VERSION_REQD    As Long = &H101&
Const SOCKET_ERROR       As Long = -1
Const WSADescription_Len As Long = 256
Const WSASYS_Status_Len  As Long = 128

Private Type HostDeType
    hName     As Long
    hAliases  As Long
    hAddrType As Integer
    hLength   As Integer
    hAddrList As Long
End Type

Private Type WinSocketDataType
    wversion        As Integer
    wHighVersion    As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets     As Integer
    iMaxUdpDg       As Integer
    lpszVendorInfo  As Long
End Type


'### Der folgende Abschnitt dient nur dazu, um festzustellen
'    ob eine Online-Verbindung besteht bzw. um diese herzu-
'    stellen und wieder abzubrechen.
'    Sie können diesen Block bei Verzicht dieser Funktionen
'    getrost löschen

Private Declare Function RasEnumConnections Lib "RasApi32" Alias "RasEnumConnectionsA" (lpRasCon As Any, lpcb As Long, lpcConnections As Long) As Long
Private Declare Function RasGetConnectStatus Lib "RasApi32" Alias "RasGetConnectStatusA" (ByVal hRasCon As Long, lpStatus As Any) As Long
Private Declare Function RasEnumEntries Lib "RasApi32" Alias "RasEnumEntriesA" (ByVal reserved As String, ByVal lpszPhonebook As String, lprasentryname As Long, lpcb As Long, lpcEntries As Long) As Long
Private Declare Function RasHangUp Lib "RasApi32" Alias "RasHangUpA" (ByVal hRasConn As Long) As Long

Const RAS_MaxEntryName  As Long = 256&
Const RAS_MaxDeviceType As Long = 16&
Const RAS_MaxDeviceName As Long = 32&

Private Type RASType
    dwSize   As Long
    hRasCon  As Long
    szEntryName(RAS_MaxEntryName)   As Byte
    szDeviceType(RAS_MaxDeviceType) As Byte
    szDeviceName(RAS_MaxDeviceName) As Byte
End Type

Private Type RASStatusType
    dwSize       As Long
    RasConnState As Long
    dwError      As Long
    szDeviceType(RAS_MaxDeviceType) As Byte
    szDeviceName(RAS_MaxDeviceName) As Byte
End Type

Private Type RASENTRYNAME95
    dwSize As Long
    szEntryName(RAS_MaxEntryName) As Byte
End Type

Dim DFUEName As String

Private Function DFUEStatus() As Boolean
    Dim RAS(255) As RASType, RASStatus As RASStatusType
    Dim lg As Long, lpcon As Long, Result As Long

    RAS(0).dwSize = 412
    lg = 256 * RAS(0).dwSize
    Result = RasEnumConnections(RAS(0), lg, lpcon)
    
    If lpcon = 0 Then
        DFUEStatus = False
        Label3.Caption = "Offline"
    Else
        RASStatus.dwSize = 160
        Result = RasGetConnectStatus(RAS(0).hRasCon, RASStatus)
        If RASStatus.RasConnState = &H2000 Then
            DFUEStatus = True
            Label3.Caption = "Online"
        Else
            DFUEStatus = False
            Label3.Caption = "Einwahl oder Trennen der Verbindung"
        End If
    End If
End Function

Private Function Online() As Boolean
    Dim Test As Boolean
    
    Test = DFUEStatus
    If Test = False Then Call MsgBox _
        ("Keine Online Verbindung vorhanden! Bitte einwählen!")
    
    Online = True ' Test
End Function

Private Function GetDFUE() As String
    Dim s As Long, ln As Long
    Dim i As Integer
    Dim r(255) As RASENTRYNAME95
    
    r(0).dwSize = 264
    s = 256 * r(0).dwSize
    Call RasEnumEntries(vbNullString, vbNullString, VarPtr(r(0)), s, ln)
    DFUEName = StrConv(r(i).szEntryName(), vbUnicode)
    DFUEName = Left$(DFUEName, InStr(DFUEName, vbNullChar) - 1)
    Shell "rundll32.exe rnaui.dll,RnaDial " & DFUEName
    Shell "rundll32.exe rnaui.dll,RnaDial " & DFUEName
End Function

Private Sub HangUp(ByVal Verbindung$)
    Dim s As Long, l As Long, ln As Long
    Dim rec As Long
    Dim aa As String
    ReDim r(255) As RASType

    r(0).dwSize = 412
    s = 256 * r(0).dwSize
    l = RasEnumConnections(r(0), s, ln)
    
    For l = 0 To ln - 1
        aa = StrConv(r(l).szEntryName(), vbUnicode)
        aa = Left$(aa, InStr(aa, Chr$(0)) - 1)
        If aa = Verbindung Then rec = RasHangUp(r(l).hRasCon)
    Next l
End Sub

Private Sub Form_Load()
    Command1.Caption = "Eigene Adresse ermitteln"
    Command2.Caption = "Dynamische IP-Adressen"
    Command3.Caption = "IP->Name"
    Command4.Caption = "Name->Ip"
    Command5.Caption = "Einwählen"
    Command6.Caption = "Beenden"
    Command7.Caption = "Auflegen"

    Timer1.Interval = 200
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    DFUEStatus
End Sub

Private Sub Command5_Click()
    Call GetDFUE
End Sub
Private Sub Command7_Click()
    If Len(DFUEName) > 0 Then Call HangUp(DFUEName)
End Sub
'### Ende des DFUE-Pfrüfungs Abschnitts


Private Sub Command1_Click()
    '### Eigene Adresse ermitteln
    InitSockets
    Text5.Text = MyHostName()
    CleanSockets
End Sub

Private Sub Command2_Click()
    '### Eigene IP-Adressen abfragen
    '    Diese Routine kann unteranderem dazu benutzt werden,
    '    dynamische durch einen Provider zugewiesene IP-Adressen
    '    auszulesen.
    '    Da hier alle eigenen IP ausgelesen werden müssen die statio-
    '    nären [Localhost (127.0.0.1), Netzwerk (192.168.xxx.xxx)]
    '    eleminiert werden. Entfernen Sie dann aber auch folgende
    '    Steuerelemente aus dem Form:
    '       Timer1
    '       Label3
    '       Command5
    '       Command7
 
    Dim X As Integer
    Dim IP As String, DNS As String, HOST As String
  
    If Not Online Then Exit Sub
    
    MousePointer = vbHourglass
    Call InitSockets
    HOST = MyHostName
    List1.Clear
     
    Do
        IP = HostByName(HOST, X)
        If Len(IP) = 0 Then Exit Do
          
        DNS = HostByAddress(IP)
        List1.AddItem "DNS: " & DNS & "  " & "IP: " & IP
        X = X + 1
    Loop
     
    Call CleanSockets
    MousePointer = vbDefault
End Sub

Private Sub Command3_Click()
    Dim aa As String
    '### DNS-Abfrage nach Domäne (gibt IP zurück)
    
    If Not Online Then Exit Sub
    
    MousePointer = vbHourglass
    Call InitSockets
    aa = HostByAddress(Text1.Text)
    If Len(aa) = 0 Then Call MsgBox("Nicht gefunden")
    Text4.Text = aa
    CleanSockets
    MousePointer = vbDefault
End Sub

Private Sub Command4_Click()
    Dim aa As String
    
    '### DNS-Abfrage nach IP (gibt Domäne zurück)
    If Not Online Then Exit Sub
    
    MousePointer = vbHourglass
    InitSockets
    aa = HostByName$(Text2.Text)
    If Len(aa) = 0 Then Call MsgBox("Nicht gefunden")
    Text3.Text = aa
    CleanSockets
    MousePointer = vbDefault
End Sub

Private Sub Command6_Click()
    Unload Me
End Sub

Private Function HostByAddress(ByVal Addresse$) As String
    Dim X As Integer
    Dim HostDeAddress As Long
    Dim aa As String, BB As String * 5
    Dim HOST As HostDeType
  
    aa = Chr$(Val(NextChar(Addresse, ".")))
    aa = aa + Chr$(Val(NextChar(Addresse, ".")))
    aa = aa + Chr$(Val(NextChar(Addresse, ".")))
    aa = aa + Chr$(Val(Addresse))
    
    HostDeAddress = gethostbyaddr(aa, Len(aa), 2)
    If HostDeAddress = 0 Then
        HostByAddress = ""
        Exit Function
    End If
    
    Call RtlMoveMemory(HOST, HostDeAddress, LenB(HOST))
 
    aa = ""
    X = 0
    Do
       Call RtlMoveMemory(ByVal BB, HOST.hName + X, 1)
       If Left$(BB, 1) = Chr$(0) Then Exit Do
       aa = aa + Left$(BB, 1)
       X = X + 1
    Loop
    
    HostByAddress = aa
End Function

Private Function HostByName(Name As String, Optional X As Integer = 0) As String
    Dim MemIp() As Byte
    Dim Y As Integer
    Dim HostDeAddress As Long, HostIp As Long
    Dim IpAddress As String
    Dim HOST As HostDeType
  
    HostDeAddress = gethostbyname(Name)
    If HostDeAddress = 0 Then
        HostByName = ""
        Exit Function
    End If
    
    Call RtlMoveMemory(HOST, HostDeAddress, LenB(HOST))
    
    For Y = 0 To X
        Call RtlMoveMemory(HostIp, HOST.hAddrList + 4 * Y, 4)
        If HostIp = 0 Then
            HostByName = ""
            Exit Function
        End If
    Next Y
    
    ReDim MemIp(1 To HOST.hLength)
    Call RtlMoveMemory(MemIp(1), HostIp, HOST.hLength)
    
    IpAddress = ""
    
    For Y = 1 To HOST.hLength
        IpAddress = IpAddress & MemIp(Y) & "."
    Next Y
    
    IpAddress = Left$(IpAddress, Len(IpAddress) - 1)
    HostByName = IpAddress
End Function

Private Function MyHostName() As String
    Dim HostName As String * 256
    
    If gethostname(HostName, 256) = SOCKET_ERROR Then
        MsgBox "Windows Sockets error " & Str(WSAGetLastError())
        Exit Function
    Else
        MyHostName = NextChar(Trim$(HostName), Chr$(0))
    End If
End Function

Private Sub InitSockets()
    Dim Result As Integer
    Dim LoBy As Integer, HiBy As Integer
    Dim SocketData As WinSocketDataType
    
    Result = WSAStartup(WS_VERSION_REQD, SocketData)
    If Result <> 0 Then
        Call MsgBox("'winsock.dll' antwortet nicht!")
        End
    End If
End Sub

Private Sub CleanSockets()
    Dim Result As Long
    
    Result = WSACleanup()
    If Result <> 0 Then
        Call MsgBox("Socket Error " & Trim$(Str$(Result)) & _
                " in Prozedur 'CleanSockets' aufgetreten !")
        End
    End If
End Sub

Private Function NextChar(Text$, Char$) As String
    Dim POS As Integer
    
    POS = InStr(1, Text, Char)
    If POS = 0 Then
        NextChar = Text
        Text = ""
    Else
        NextChar = Left$(Text, POS - 1)
        Text = Mid$(Text, POS + Len(Char))
    End If
End Function

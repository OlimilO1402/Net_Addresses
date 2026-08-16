VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Net Addresses IP-V4, -V6, MAC"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11055
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "FMain"
   ScaleHeight     =   7830
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   7440
      TabIndex        =   14
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox TBTests 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   3
      Top             =   1560
      Width           =   9735
   End
   Begin VB.CommandButton BtnDoSomeTests 
      Caption         =   "Do Some Tests"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton BtnTestMACAddress 
      Caption         =   "MAC-Addr >>"
      Height          =   375
      Left            =   9120
      TabIndex        =   6
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton BtnIPV4AddRnd 
      Caption         =   "IPv4 Add Rnd"
      Height          =   375
      Left            =   7440
      TabIndex        =   5
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton BtnCreateIPAddr 
      Caption         =   "Create IP-Address"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton BtnIPV4Add1 
      Caption         =   "IPv4 Add 1"
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox TxtIP 
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
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   5535
   End
   Begin VB.CommandButton BtnBspIPv6nZoneID 
      Caption         =   "IPv6+zoneID"
      Height          =   375
      Left            =   4080
      TabIndex        =   13
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton BtnBspIPv6nPort 
      Caption         =   "IPv6+Port"
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton BtnBspIPv6incl4 
      Caption         =   "IPv6incl4"
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton BtnBspIPv6 
      Caption         =   "IPv6"
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton BtnBspIPv4 
      Caption         =   "IPv4"
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Bsps:"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command2_Click()
    TxtIP.Text = "2001:0db8:85a3:08d3:1319:8a2e:0370:7347"
    Dim ip As IPAddress: Set ip = MNew.IPAddressV(TxtIP.Text)
    DebugWriteLine ip.ToStr
    DebugWriteLine ip.ProviderRIRNet
    DebugWriteLine ip.EnduserNet
    DebugWriteLine ip.ProviderPrefix
    DebugWriteLine ip.InterfaceID
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    BtnBspIPv4_Click
End Sub

Private Sub Form_Resize()
    Dim L As Single, t As Single, W As Single, H As Single
    t = TBTests.Top
    W = Me.ScaleWidth
    H = Me.ScaleHeight - t
    If W > 0 And H > 0 Then TBTests.Move L, t, W, H
End Sub

Private Sub BtnBspIPv4_Click():        TxtIP.Text = "192.168.178.100":                         End Sub
Private Sub BtnBspIPv6_Click():        TxtIP.Text = "2001:db8:0:8d3:0:8a2e:70:7344":           End Sub
Private Sub BtnBspIPv6incl4_Click():   TxtIP.Text = "::ffff:127.0.0.1":                        End Sub
Private Sub BtnBspIPv6nPort_Click():   TxtIP.Text = "[::]:653061":                             End Sub
Private Sub BtnBspIPv6nZoneID_Click(): TxtIP.Text = "fe80::7645:6de2:ff:1%eth0":               End Sub

Private Sub BtnCreateIPAddr_Click()
    
    Dim ip As IPAddress: Set ip = MNew.IPAddressV(TxtIP.Text)
    DebugWriteLine ip.ToStr

End Sub

'Private Sub BtnCreateIPAddressV4_Click()
'End Sub
'Private Sub BtnCreateIPAddressV6_Click()
'End Sub

Private Sub BtnIPV4Add1_Click()
    Dim ip As IPAddress: Set ip = MNew.IPAddressV(TxtIP.Text)
    ip.OneUp
    TxtIP.Text = ip.ToStr
    DebugWriteLine ip.ToStr
End Sub

Private Sub BtnIPV4AddRnd_Click()
    Dim ip As IPAddress: Set ip = MNew.IPAddressV(TxtIP.Text)
    Dim s As String: s = ip.ToStr
    Randomize
    Dim b As Byte: b = Rnd * 255
    ip.Add b
    TxtIP.Text = ip.ToStr
    DebugWriteLine s & " + " & b & " = " & ip.ToStr
End Sub


Private Sub BtnDoSomeTests_Click()
    
    Dim ip As New IPAddress
    
    ip.ValueB1 = 192
    ip.ValueB2 = 168
    ip.ValueB3 = 178
    ip.ValueB4 = 100
                                                           '2147483647
                                                           '1689430208 = &H64B2A8C0
    DebugWriteLine "IP-V4.B1-4 = " & ip.ToStr & "; Value-Lng: " & ip.AddressL & " = &H" & Hex(ip.AddressL)
    
    ip.AddressL = &HFFFFFF
    
    DebugWriteLine "IP-V4.B1-4 = " & ip.ToStr & "; Value-Lng: " & ip.AddressL & " Value-Cur: " & ip.Address & " = &H" & Hex(ip.AddressL) '1689430208 &H64B2A8C0
    
    ip.ValueI1 = &H1234
    ip.ValueI2 = &H5678
    ip.ValueI3 = &H90AB
    ip.ValueI4 = &HCDEF
    ip.ValueI5 = &H1234
    ip.ValueI6 = &H5678
    ip.ValueI7 = &H90AB
    ip.ValueI8 = &HCDEF
    
    DebugWriteLine "IP-V6.I1-8 = " & ip.ToStr
    
    Set ip = MNew.IPAddress(192, 168, 178, 100)
    
    DebugWriteLine "New IPAddress = " & ip.ToStr
    
    Set ip = MNew.IPAddressV4Rnd
    
    DebugWriteLine "New IPAddress = " & ip.ToStr
    
    Set ip = MNew.IPAddress(&H1234, &H5678, &H80AB, &HCDEF, &H1234, &H5678, &H80AB, &HCDEF)
    
    DebugWriteLine "New IPAddress = " & ip.ToStr
    
    Set ip = MNew.IPAddressV(Array(192, 168, 178, 100))
    
    DebugWriteLine "New IPAddressV(Array(192, 168, 178, 100)) = " & ip.ToStr
        
    'Set ip = MNew.IPAddressV(Array(192, 168, 178, 100))
    
    'DebugWriteLine ip.ToStr
    
    Set ip = MNew.IPAddressV(Array(CByte(192), CByte(168), CByte(178), CByte(100)))
    
    DebugWriteLine "New IPAddressV(Array(CByte(192), CByte(168), CByte(178), CByte(100))) = " & ip.ToStr
    
    ReDim bb(0 To 3) As Byte: bb(0) = 192: bb(1) = 168: bb(2) = 178: bb(3) = 100
    Set ip = MNew.IPAddressV(bb)
    
    DebugWriteLine "ReDim bb(0 To 3) As Byte: bb(0) = 192: bb(1) = 168: bb(2) = 178: bb(3) = 100"
    DebugWriteLine "New IPAddressV(bb) = " & ip.ToStr
    
    ReDim ii(0 To 7) As Integer: ii(0) = &H1234: ii(1) = &H5678: ii(2) = &H90AB: ii(3) = &HCDEF: ii(4) = &H1234: ii(5) = &H5678: ii(6) = &H90AB: ii(7) = &HCDEF
    Set ip = MNew.IPAddressV(ii)
    
    DebugWriteLine "ReDim ii(0 To 7) As Integer: ii(0) = &H1234: ii(1) = &H5678: ii(2) = &H90AB: ii(3) = &HCDEF: ii(4) = &H1234: ii(5) = &H5678: ii(6) = &H90AB: ii(7) = &HCDEF"
    DebugWriteLine "New IPAddressV(ii) = " & ip.ToStr
    
    Set ip = MNew.IPAddress(192, 168, 178, 100)
    Dim ip2 As IPAddress: Set ip2 = ip.Clone
    If ip.compare(ip2) = 0 Then
        DebugWriteLine "IP-V4: " & ip.ToStr & " = " & ip2.ToStr
    End If
    ip2.OneUp
    If ip.compare(ip2) < 0 Then
        DebugWriteLine "IP-V4: " & ip.ToStr & " < " & ip2.ToStr
    End If
    ip.OneUp: ip.OneUp
    If ip.compare(ip2) > 0 Then
        DebugWriteLine "IP-V4: " & ip.ToStr & " > " & ip2.ToStr
    End If
    
End Sub

Private Sub Command1_Click()
    Dim nla As Long: nla = MNetBios.EnumLanAdapter
    MsgBox nla
    Dim ma As String
    ma = MNetBios.GetMACAddress(1)
    MsgBox ma
End Sub

Sub DebugWriteLine(s As String)
    TBTests.Text = TBTests.Text & s & vbCrLf
End Sub

Private Sub BtnTestMACAddress_Click()
    FMACAddr.Show
End Sub

Sub Set_Static()
     Dim objWMIService:   Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
     Dim colNetAdapters: Set colNetAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration " & "where IPEnabled=TRUE")
     Dim strIPAddress:         strIPAddress = Array("192.168.1.155")
     Dim strSubnetMask:       strSubnetMask = Array("255.255.255.0")
     Dim strGateway:             strGateway = Array("192.168.1.1")
     Dim strGatewaymetric: strGatewaymetric = Array(1)
     Dim strDNS
     Dim objNetAdapter
     For Each objNetAdapter In colNetAdapters
         Dim errEnable:     errEnable = objNetAdapter.EnableStatic(strIPAddress, strSubnetMask)
         Dim errGateways: errGateways = objNetAdapter.SetGateways(strGateway, strGatewaymetric)
         Dim errDNS:           errDNS = objNetAdapter.SetDNSServerSearchOrder(strDNS)
     Next
End Sub



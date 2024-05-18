VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnCreateIPAddressV6 
      Caption         =   "Create IP-AddressV6"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox TxtIPV6 
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
      Left            =   120
      TabIndex        =   2
      Text            =   "ABCD:EF01:2345:6789:ABCD:EF01:2345:6789"
      Top             =   720
      Width           =   5055
   End
   Begin VB.CommandButton BtnIPV4AddRnd 
      Caption         =   "IP-V4 Add Rnd"
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton BtnIPV4Add1 
      Caption         =   "IP-V4 Add 1"
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton BtnCreateIPAddressV4 
      Caption         =   "Create IP-AddressV4"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox TxtIPV4 
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
      Left            =   120
      TabIndex        =   0
      Text            =   "192.168.178.100"
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton BtnTestMACAddress 
      Caption         =   "MAC-Addr >>"
      Height          =   375
      Left            =   8400
      TabIndex        =   8
      Top             =   120
      Width           =   1335
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
      TabIndex        =   5
      Top             =   1800
      Width           =   9735
   End
   Begin VB.CommandButton BtnDoSomeTests 
      Caption         =   "Do Some Tests"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtnCreateIPAddressV4_Click()
    
    Dim ip As IPAddress: Set ip = MNew.IPAddressV(TxtIPV4.Text)
    DebugWriteLine ip.ToStr
    
End Sub

Private Sub BtnIPV4Add1_Click()
    Dim ip As IPAddress: Set ip = MNew.IPAddressV(TxtIPV4.Text)
    ip.OneUp
    TxtIPV4.Text = ip.ToStr
    DebugWriteLine ip.ToStr
End Sub

Private Sub BtnIPV4AddRnd_Click()
    Dim ip As IPAddress: Set ip = MNew.IPAddressV(TxtIPV4.Text)
    Dim s As String: s = ip.ToStr
    Randomize
    Dim b As Byte: b = Rnd * 255
    ip.Add b
    TxtIPV4.Text = ip.ToStr
    DebugWriteLine s & " + " & b & " = " & ip.ToStr
End Sub

Private Sub BtnCreateIPAddressV6_Click()
    
    Dim ip As IPAddress: Set ip = MNew.IPAddressV(TxtIPV6.Text)
    DebugWriteLine ip.ToStr
    
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
    If ip.Compare(ip2) = 0 Then
        DebugWriteLine "IP-V4: " & ip.ToStr & " = " & ip2.ToStr
    End If
    ip2.OneUp
    If ip.Compare(ip2) < 0 Then
        DebugWriteLine "IP-V4: " & ip.ToStr & " < " & ip2.ToStr
    End If
    ip.OneUp: ip.OneUp
    If ip.Compare(ip2) > 0 Then
        DebugWriteLine "IP-V4: " & ip.ToStr & " > " & ip2.ToStr
    End If
    
End Sub

Sub DebugWriteLine(s As String)
    TBTests.Text = TBTests.Text & s & vbCrLf
End Sub

Private Sub BtnTestMACAddress_Click()
    Form2.Show
End Sub

Private Sub Form_Resize()
    Dim L As Single, T As Single, W As Single, H As Single
    T = TBTests.Top
    W = Me.ScaleWidth
    H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then TBTests.Move L, T, W, H
End Sub

'Frage in AVB-Forum von tmsrtl am 29.03.24 um 18:25:08:
' "network Adresse einstellen":
'Ich benutze folgenden Code um die IP usw eizustellen, will es aber nur für einen (!) bestimmten Netwerkadapter tun.
'Wie muss ich das denn bitte tun !?
'Nickname: tmsrtl
'Vorname : thomas
'Nachname: Roethling
'E-Mail-Adresse  info@incom.ca

Sub Set_Static()
     Dim objWMIService:   Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
     Dim colNetAdapters: Set colNetAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration " & "where IPEnabled=TRUE")
     Dim strIPAddress:         strIPAddress = Array("192.168.1.155")
     Dim strSubnetMask:       strSubnetMask = Array("255.255.255.0")
     Dim strGateway:             strGateway = Array("192.168.1.1")
     Dim strGatewaymetric: strGatewaymetric = Array(1)
     'Dim strDNS: strDNS = Array("10.10.10.10", "10.10.10.11")
     Dim objNetAdapter
     For Each objNetAdapter In colNetAdapters
         Dim errEnable:     errEnable = objNetAdapter.EnableStatic(strIPAddress, strSubnetMask)
         Dim errGateways: errGateways = objNetAdapter.SetGateways(strGateway, strGatewaymetric)
         Dim errDNS:           errDNS = objNetAdapter.SetDNSServerSearchOrder(strDNS)
     Next
End Sub



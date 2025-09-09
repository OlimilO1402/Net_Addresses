Attribute VB_Name = "MNetBios"
'Modul: modNetBios
Option Explicit

Private Declare Function Netbios Lib "netapi32" (pncb As NCB) As Byte

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Private Declare Sub CopyMemory_ByRef Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Const NCBADDNAME   As Byte = &H30
Private Const NCBDELNAME   As Byte = &H31
Private Const NCBRESET     As Byte = &H32
Private Const NCBASTAT     As Byte = &H33
Private Const NCBADDGRNAME As Byte = &H36
Private Const NCBENUM      As Byte = &H37
Private Const NCBFINDNAME  As Byte = &H78
Private Const NCBNAMSZ     As Byte = 16

Public Const NRC_GOODRET     As Long = &H0&
Public Const NRC_BUFLEN      As Long = &H1&
Public Const NRC_ILLCMD      As Long = &H3&

Public Const NRC_CMDTMO      As Long = &H5&
Public Const NRC_INCOMP      As Long = &H6&
Public Const NRC_BADDR       As Long = &H7&
Public Const NRC_SNUMOUT     As Long = &H8&
Public Const NRC_NORES       As Long = &H9&

Public Const NRC_SCLOSED     As Long = &HA&
Public Const NRC_CMDCAN      As Long = &HB&
Public Const NRC_DUPNAME     As Long = &HD&
Public Const NRC_NAMTFUL     As Long = &HE&

Public Const NRC_ACTSES      As Long = &HF&

Public Const NRC_LOCTFUL     As Long = &H11
Public Const NRC_REMTFUL     As Long = &H12
Public Const NRC_ILLNN       As Long = &H13
Public Const NRC_NOCALL      As Long = &H14
Public Const NRC_NOWILD      As Long = &H15
Public Const NRC_INUSE       As Long = &H16
Public Const NRC_NAMERR      As Long = &H17
Public Const NRC_SABORT      As Long = &H18
Public Const NRC_NAMCONF     As Long = &H19

Public Const NRC_IFBUSY      As Long = &H21
Public Const NRC_TOOMANY     As Long = &H22
Public Const NRC_BRIDGE      As Long = &H23
Public Const NRC_CANOCCR     As Long = &H24
Public Const NRC_CANCEL      As Long = &H26
Public Const NRC_DUPENV      As Long = &H30
Public Const NRC_ENVNOTDEF   As Long = &H34
Public Const NRC_OSRESNOTAV  As Long = &H35
Public Const NRC_MAXAPPS     As Long = &H36
Public Const NRC_NOSAPS      As Long = &H37
Public Const NRC_NORESOURCES As Long = &H38
Public Const NRC_INVADDRESS  As Long = &H39
Public Const NRC_SYSTEM      As Long = &H40
Public Const NRC_INVDDID     As Long = &H3B
Public Const NRC_LOCKFAIL    As Long = &H3C
Public Const NRC_OPENERR     As Long = &H3F
Public Const NRC_PENDING     As Long = &HFF&

Private Type NCB
    ncb_Command    As Byte
    ncb_RetCode    As Byte
    ncb_LSN        As Byte
    ncb_Num        As Byte
    ncb_pBuffer    As Long
    ncb_Length     As Integer
    ncb_CallName   As String * NCBNAMSZ '16
    ncb_Name       As String * NCBNAMSZ '16
    ncb_RTO        As Byte
    ncb_STO        As Byte
    ncb_Post       As Long
    ncb_Lana_Num   As Byte
    ncb_Cmd_Cplt   As Byte
    ncb_Reserve(9) As Byte
    ncb_Event      As Long
End Type

Private Type ADAPTER_STATUS
    adapter_address(5) As Byte
    rev_major          As Byte
    reserved0          As Byte
    adapter_type       As Byte
    rev_minor          As Byte
    duration           As Integer
    frmr_recv          As Integer
    frmr_xmit          As Integer
    iframe_recv_err    As Integer
    xmit_aborts        As Integer
    xmit_success       As Long
    recv_success       As Long
    iframe_xmit_err    As Integer
    recv_buff_unavail  As Integer
    t1_timeouts        As Integer
    ti_timeouts        As Integer
    Reserved1          As Long
    free_ncbs          As Integer
    max_cfg_ncbs       As Integer
    max_ncbs           As Integer
    xmit_buf_unavail   As Integer
    max_dgram_size     As Integer
    pending_sess       As Integer
    max_cfg_sess       As Integer
    max_sess           As Integer
    max_sess_pkt_size  As Integer
    name_count         As Integer
End Type

Private Type NAME_BUFFER
    Name       As String * NCBNAMSZ
    name_num   As Integer
    name_flags As Integer
End Type

Private Type ASTAT
    adapt        As ADAPTER_STATUS
    NameBuff(30) As NAME_BUFFER
End Type

Private Type ENUM_LANA
    bCount     As Byte
    bLana(300) As Byte
End Type

Public Function EnumLanAdapter(bLanArray() As Byte) As Long
    Dim bRetEnum As ENUM_LANA
    Dim myNcb As NCB
    With myNcb
        .ncb_Command = NCBENUM           'NetBios Command Enum setzen
        .ncb_pBuffer = VarPtr(bRetEnum)  'Bufferpointer eintragen
        .ncb_Length = Len(bRetEnum)      'Größe des Buffers angeben
    End With
    'Alle aktiven Netzwerkkarten enumerieren
    If Netbios(myNcb) = NRC_GOODRET Then
        'Anzahl der aktiven Netzwerkkarten auslesen
        If bRetEnum.bCount Then
            EnumLanAdapter = CLng(bRetEnum.bCount)
            'Nur auslesen, wenn mindestens 1 Netzwerkkarte gefunden wurde
            'Return Array anpassen
            ReDim bLanArray(1 To bRetEnum.bCount)
            'Daten ins Array kopieren
            CopyMemory_ByRef bLanArray(1), bRetEnum.bLana(0), bRetEnum.bCount
        End If
    End If
End Function

Public Function ResetAdapter(lLanNumber As Byte, lSessions As Long, lMaxNames As Long) As Long
    Dim myNcb As NCB
    With myNcb
        .ncb_Lana_Num = lLanNumber                  'Welche Netzwerkkarte soll resettet werden
        .ncb_Command = NCBRESET                     'NetBios Command setzen
        .ncb_LSN = 0
        Mid$(.ncb_CallName, 1, 1) = Chr$(lSessions) 'Maximale Anzahl an Sessions seztzen
        Mid$(.ncb_CallName, 3, 1) = Chr$(lMaxNames) 'Maximale Anzahl an Namen setzen
        If Netbios(myNcb) = NRC_GOODRET Then ResetAdapter = 1 'Netzwerkkarte resetten
    End With
End Function

Public Function GetMACAddress(ByVal lLanNumber As Byte, Optional Server As String = "*") As String
    'Dim bRet    As Byte
    Dim myNcb   As NCB
    Dim myASTAT As ASTAT
    With myNcb
        .ncb_Command = NCBASTAT        'NetBios Command setzen
        .ncb_Lana_Num = lLanNumber     'Welche Netzwerkkarte soll benutzt werden
        .ncb_CallName = Server         'Server setzen, dies kann auch ein RemoteHost sein!
        .ncb_Length = Len(myASTAT)     'Größe des Speichers setzen
        .ncb_pBuffer = VarPtr(myASTAT) 'pASTAT 'Buffer eintragen
    End With
    
    'Karte auslesen
    Dim ret As Long
    ret = Netbios(myNcb)
    If ret = NRC_GOODRET Then
        With myASTAT.adapt
            'Daten in die neue
            GetMACAddress = Hex2(.adapter_address(0)) & "-" & Hex2(.adapter_address(1)) & "-" & Hex2(.adapter_address(2)) & "-" & Hex2(.adapter_address(3)) & "-" & Hex2(.adapter_address(4)) & "-" & Hex2(.adapter_address(5))
        End With
    End If
End Function

Private Function Hex2(ByVal Value As Byte) As String
    Hex2 = Hex(Value): If Len(Hex2) = 1 Then Hex2 = "0" & Hex2
End Function

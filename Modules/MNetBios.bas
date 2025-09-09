Attribute VB_Name = "MNetBios"
'Modul: modNetBios
Option Explicit

'https://learn.microsoft.com/en-us/windows/win32/api/nb30/nf-nb30-netbios
Private Declare Function Netbios Lib "netapi32" (pncb As NCB) As Byte

'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
'Private Declare Sub CopyMemory_ByRef Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal bytlen As Long)


Private Const NCBCALL        As Long = &H10& 'Opens a session with another name.
Private Const NCBLISTEN      As Long = &H11& 'Enables a session to be opened with another name (local or remote).
Private Const NCBHANGUP      As Long = &H12& 'Closes a specified session.
Private Const NCBSEND        As Long = &H14& 'Sends data to the specified session partner.
Private Const NCBRECV        As Long = &H15& 'Receives data from the specified session partner.
Private Const NCBRECVANY     As Long = &H16& 'Receives data from any session corresponding to a specified name.
Private Const NCBCHAINSEND   As Long = &H17& 'Sends the contents of two data buffers to the specified session partner.
Private Const NCBDGSEND      As Long = &H20& 'Sends datagram to a specified name.
Private Const NCBDGRECV      As Long = &H21& 'Receives a datagram from any name.
Private Const NCBDGSENDBC    As Long = &H22& 'Sends a broadcast datagram to every host on the local area network (LAN).
Private Const NCBDGRECVBC    As Long = &H23& 'Receives a broadcast datagram from any name.

Private Const NCBADDNAME     As Long = &H30& 'Adds a unique name to the local name table. The TDI driver ensures that the name is unique across the network.
Private Const NCBDELNAME     As Long = &H31& 'Deletes a name from the local name table.
Private Const NCBRESET       As Long = &H32& 'Resets a LAN adapter. An adapter must be reset before it can accept any other NCB command that specifies the same number in the ncb_lana_num member.
Private Const NCBASTAT       As Long = &H33& 'Retrieves the status of either a local or remote adapter. When this code is specified, the ncb_buffer member points to a buffer to be filled with an ADAPTER_STATUS structure, followed by an array of NAME_BUFFER structures.
Private Const NCBSSTAT       As Long = &H34& 'Retrieves the status of the session. When this value is specified, the ncb_buffer member points to a buffer to be filled with a SESSION_HEADER structure, followed by one or more SESSION_BUFFER structures.
Private Const NCBCANCEL      As Long = &H35& 'Cancels a previous pending command.
Private Const NCBADDGRNAME   As Long = &H36& 'Adds a group name to the local name table. This name cannot be used by another process on the network as a unique name, but it can be added by anyone as a group name.
Private Const NCBENUM        As Long = &H37& ' Windows Server 2003, Windows XP, Windows 2000, and Windows NT:  Enumerates LAN adapter (LANA) numbers. When this code is specified, the ncb_buffer member points to a buffer to be filled with a LANA_ENUM structure. NCBENUM is not a standard NetBIOS 3.0 command.

Private Const NCBUNLINK      As Long = &H70& 'Unlinks the adapter.  This command is provided for compatibility with earlier versions of NetBIOS. It has no effect in Windows.
Private Const NCBCHAINSENDNA As Long = &H72& 'Sends the contents of two data buffers to the specified session partner and does not wait for acknowledgment.
Private Const NCBLANSTALERT  As Long = &H73& 'Windows Server 2003, Windows XP, Windows 2000, and Windows NT:  Notifies the user of LAN failures that last for more than one minute.
Private Const NCBACTION      As Long = &H77& 'Windows Server 2003, Windows XP, Windows 2000, and Windows NT:  Enables extensions to the transport interface. NCBACTION is mapped to TdiAction. When this code is specified, the ncb_buffer member points to a buffer to be filled with an ACTION_HEADER structure, which is optionally followed by data. NCBACTION commands cannot be canceled by using NCBCANCEL. NCBACTION is not a standard NetBIOS 3.0 command.
Private Const NCBFINDNAME    As Long = &H78& '    Determines the location of a name on the network. When this code is specified, the ncb_buffer member points to a buffer to be filled with a FIND_NAME_HEADER structure followed by one or more FIND_NAME_BUFFER structures.
Private Const NCBTRACE       As Long = &H79& 'Activates or deactivates NCB tracing. This command is not supported.

Private Const NCBNAMSZ       As Long = 16
'Use the following values to specify how resources are to be freed:
'    If ncb_lsn is not 0x00, all resources associated with ncb_lana_num are to be freed.
'    If ncb_lsn is 0x00, all resources associated with ncb_lana_num are to be freed, and new resources are to be allocated. The ncb_callname[0] byte specifies the maximum number of sessions, and the ncb_callname[2] byte specifies the maximum number of names. A nonzero value for the ncb_callname[3] byte requests that the application use NAME_NUMBER_1.
'Private Const NCBSENDNA      As Long = &H0& 'Sends data to specified session partner and does not wait for acknowledgment.


Public Const NRC_GOODRET     As Long = &H0&
'2
Public Const NRC_BUFLEN      As Long = &H1&
Public Const NRC_ILLCMD      As Long = &H3&

Public Const NRC_CMDTMO      As Long = &H5&
Public Const NRC_INCOMP      As Long = &H6&
Public Const NRC_BADDR       As Long = &H7&
Public Const NRC_SNUMOUT     As Long = &H8&
Public Const NRC_NORES       As Long = &H9&
Public Const NRC_SCLOSED     As Long = &HA&
Public Const NRC_CMDCAN      As Long = &HB&
'C
Public Const NRC_DUPNAME     As Long = &HD&
Public Const NRC_NAMTFUL     As Long = &HE&
Public Const NRC_ACTSES      As Long = &HF&
'10
Public Const NRC_LOCTFUL     As Long = &H11&
Public Const NRC_REMTFUL     As Long = &H12&
Public Const NRC_ILLNN       As Long = &H13&
Public Const NRC_NOCALL      As Long = &H14&
Public Const NRC_NOWILD      As Long = &H15&
Public Const NRC_INUSE       As Long = &H16&
Public Const NRC_NAMERR      As Long = &H17&
Public Const NRC_SABORT      As Long = &H18&
Public Const NRC_NAMCONF     As Long = &H19&
'20
Public Const NRC_IFBUSY      As Long = &H21&
Public Const NRC_TOOMANY     As Long = &H22&
Public Const NRC_BRIDGE      As Long = &H23&
Public Const NRC_CANOCCR     As Long = &H24&
'25
Public Const NRC_CANCEL      As Long = &H26&
'27, 28, 29
Public Const NRC_DUPENV      As Long = &H30&
'31, 32, 33
Public Const NRC_ENVNOTDEF   As Long = &H34&
Public Const NRC_OSRESNOTAV  As Long = &H35&
Public Const NRC_MAXAPPS     As Long = &H36&
Public Const NRC_NOSAPS      As Long = &H37&
Public Const NRC_NORESOURCES As Long = &H38&
Public Const NRC_INVADDRESS  As Long = &H39&
Public Const NRC_INVDDID     As Long = &H3B&
Public Const NRC_LOCKFAIL    As Long = &H3C&
Public Const NRC_OPENERR     As Long = &H3F&
Public Const NRC_SYSTEM      As Long = &H40&
Public Const NRC_PENDING     As Long = &HFF&

'typedef struct _NCB {
'  UCHAR  ncb_command;
'  UCHAR  ncb_retcode;
'  UCHAR  ncb_lsn;
'  UCHAR  ncb_num;
'  PUCHAR ncb_buffer;
'  WORD   ncb_length;
'  UCHAR  ncb_callname[NCBNAMSZ];
'  UCHAR  ncb_name[NCBNAMSZ];
'  UCHAR  ncb_rto;
'  UCHAR  ncb_sto;
'  void()(_NCB *)  * ncb_post;
'  UCHAR  ncb_lana_num;
'  UCHAR  ncb_cmd_cplt;
'#if ...
'  UCHAR  ncb_reserve[18];
'#Else
'  UCHAR  ncb_reserve[10];
'#End If
'  HANDLE ncb_event;
'} NCB, *PNCB;

Private Type NCB
    ncb_Command    As Byte
    ncb_RetCode    As Byte
    ncb_LSN        As Byte
    ncb_Num        As Byte
    ncb_Buffer     As LongPtr
    ncb_Length     As Integer
    ncb_CallName(0 To NCBNAMSZ - 1) As Byte 'String * NCBNAMSZ '16
    ncb_Name(0 To NCBNAMSZ - 1)     As Byte 'String * NCBNAMSZ '16
    ncb_rto        As Byte
    ncb_sto        As Byte
    ncb_Post       As LongPtr
    ncb_Lana_Num   As Byte
    ncb_Cmd_Cplt   As Byte
    ncb_Reserve(0 To 9) As Byte
    'ncb_Reserve(0 to 17) As Byte
    ncb_hEvent     As LongPtr
End Type

Private Type ENUM_LANA
    bCount     As Byte
    bLana(300) As Byte
End Type
Private m_RetEnum As ENUM_LANA
Private m_NCB     As NCB
Private m_bLanArray() As Byte

Public Function EnumLanAdapter() As Long
    With m_NCB
        .ncb_Command = NCBENUM           'NetBios Command Enum setzen
        .ncb_pBuffer = VarPtr(m_RetEnum)   'Bufferpointer eintragen
        .ncb_Length = Len(m_RetEnum)       'Größe des Buffers angeben
    End With
    'Alle aktiven Netzwerkkarten enumerieren
    Dim ret As Long: ret = Netbios(m_NCB)
    If ret <> NRC_GOODRET Then
        'Fehler ermitteln
        Exit Function
    End If
    'Anzahl der aktiven Netzwerkkarten auslesen
    If m_RetEnum.bCount Then
        EnumLanAdapter = CLng(m_RetEnum.bCount)
        'Nur auslesen, wenn mindestens 1 Netzwerkkarte gefunden wurde
        'Return Array anpassen
        ReDim m_bLanArray(1 To bRetEnum.bCount)
        'Daten ins Array kopieren
        CopyMemory_ByRef m_bLanArray(1), bRetEnum.bLana(0), bRetEnum.bCount
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


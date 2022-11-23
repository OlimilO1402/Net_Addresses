Attribute VB_Name = "Module1"
Option Explicit

'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source!

'halllöle ich habe mir mal erlaubt das Ding hier zu Vereinfachen !
'hoffe jetzt finden sich hier einige mehr durch
'ausserdem habe ich die Intger durch Long ersetzt!
'Und die Max hochgesetzt auf 1024 geändert falls das
'nicht ausreicht dann 2048 einsetzen
'Euer Mgalpha = MGalpha@gmx.de
'
'Tut mir aber bitte ein Gefallen und macht nicht die selben Fehler
'wie dieser Programierer
'Einfach zu Koplieziert !
'und zuviele überflüßige Programm zeilen !
'das klaut leistung


Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias _
        "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal _
        lpPassword As String, ByVal lpUserName As String, ByVal _
        dwFlags As Long) As Long
        
Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias _
        "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As _
        Long, ByVal dwUsage As Long, lpNetResource As Any, lphEnum _
        As Long) As Long
        
Private Declare Function WNetEnumResource Lib "mpr.dll" Alias _
        "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, _
        lpBuffer As NETRESOURCE, lpBufferSize As Long) As Long
        
Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum _
        As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (lpTo As Any, lpFrom As Any, ByVal lLen As Long)
        
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" _
        (ByVal lpString As Any) As Long
        
Private Type NETRESOURCE
  dwScope As Long
  dwType As Long
  dwDisplayType As Long
  dwUsage As Long
  pLocalName As Long
  pRemoteName As Long
  pComment As Long
  pProvider As Long
End Type

Private Type NETRESOURCE_REAL
  dwScope As Long
  dwType As Long
  dwDisplayType As Long
  dwUsage As Long
  sLocalName As String
  sRemoteName As String
  sComment As String
  sProvider As String
End Type

Private Const RESOURCE_CONNECTED As Long = &H1&
Private Const RESOURCE_GLOBALNET As Long = &H2&
Private Const RESOURCE_REMEMBERED As Long = &H3&

Private Const RESDTYPE_DIRECTORY As Long = &H9&
Private Const RESDTYPE_DOMAIN As Long = &H1&
Private Const RESDTYPE_FILE As Long = &H4&
Private Const RESDTYPE_GENERIC As Long = &H0&
Private Const RESDTYPE_GROUP As Long = &H5&
Private Const RESDTYPE_NETWORK As Long = &H6&
Private Const RESDTYPE_ROOT As Long = &H7&
Private Const RESDTYPE_SERVER As Long = &H2&
Private Const RESDTYPE_SHARE As Long = &H3&
Private Const RESDTYPE_SHAREADMIN As Long = &H8&

Private Const RESOURCETYPE_ANY As Long = &H0&
Private Const RESOURCETYPE_DISK As Long = &H1&
Private Const RESOURCETYPE_PRINT As Long = &H2&
Private Const RESOURCETYPE_UNKNOWN As Long = &HFFFF&

Private Const RESOURCEUSAGE_ALL As Long = &H0&
Private Const RESOURCEUSAGE_CONNECTABLE As Long = &H1&
Private Const RESOURCEUSAGE_CONTAINER As Long = &H2&
Private Const RESOURCEUSAGE_RESERVED As Long = &H80000000

Private Const NO_ERROR As Long = 0&
Private Const ERROR_MORE_DATA As Long = 234&
Private Const RESOURCE_ENUM_ALL As Long = &HFFFF&


Public Function Netsuche1()
    Const MAX_RESOURCES = 2048
    Const NOT_A_CONTAINER = -1
    
    'VARS
    Dim bFirstTime As Boolean, Läufer As Long, lRet As Long, hEnum As Long
    Dim lCnt As Long, lMin As Long, lLen As Long, lBufSize As Long
    Dim lLastIx As Long, l As Long, NetAusgabe As String
    
    'samel Vars
    Dim uNetApi(0 To MAX_RESOURCES) As NETRESOURCE
    Dim uNet() As NETRESOURCE_REAL
    
    bFirstTime = True
    
    '### Ressourcen Auslesen
    Do
        If bFirstTime Then
            lRet = WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, _
                                RESOURCEUSAGE_ALL, ByVal 0&, hEnum)
            bFirstTime = False
        Else
            If uNet(lLastIx).dwUsage And RESOURCEUSAGE_CONTAINER Then
                lRet = WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, _
                                RESOURCEUSAGE_ALL, uNet(lLastIx), hEnum)
        Else
            lRet = NOT_A_CONTAINER
            hEnum = 0
        End If
        
            lLastIx = lLastIx + 1
        End If
        
        If lRet = NO_ERROR Then
            lCnt = RESOURCE_ENUM_ALL
            
            Do
                lBufSize = UBound(uNetApi) * Len(uNetApi(0)) / 2
                lRet = WNetEnumResource(hEnum, lCnt, uNetApi(0), lBufSize)
                
                If lCnt > 0 Then
                
                    ReDim Preserve uNet(0 To lMin + lCnt - 1) _
                                   As NETRESOURCE_REAL
                    
                    For l = 0 To lCnt - 1
                        uNet(lMin + l).dwScope = uNetApi(l).dwScope
                        uNet(lMin + l).dwType = uNetApi(l).dwType
                        uNet(lMin + l).dwDisplayType = uNetApi(l).dwDisplayType
                        uNet(lMin + l).dwUsage = uNetApi(l).dwUsage
                        
                        If uNetApi(l).pLocalName Then
                            lLen = lstrlen(uNetApi(l).pLocalName)
                            uNet(lMin + l).sLocalName = Space$(lLen)
                            CopyMemory ByVal uNet(lMin + l).sLocalName, _
                                ByVal uNetApi(l).pLocalName, lLen
                        End If
                        
                        If uNetApi(l).pRemoteName Then
                            lLen = lstrlen(uNetApi(l).pRemoteName)
                            uNet(lMin + l).sRemoteName = Space$(lLen)
                            
                            CopyMemory ByVal uNet(lMin + l).sRemoteName, _
                                ByVal uNetApi(l).pRemoteName, lLen
                        End If
                        
                        If uNetApi(l).pComment Then
                            lLen = lstrlen(uNetApi(l).pComment)
                            uNet(lMin + l).sComment = Space$(lLen)
                            
                            CopyMemory ByVal uNet(lMin + l).sComment, _
                                ByVal uNetApi(l).pComment, lLen
                        End If
                        
                        If uNetApi(l).pProvider Then
                            lLen = lstrlen(uNetApi(l).pProvider)
                            uNet(lMin + l).sProvider = Space$(lLen)
                            
                            CopyMemory ByVal uNet(lMin + l).sProvider, _
                                ByVal uNetApi(l).pProvider, lLen
                        End If
                        DoEvents
                    Next l
                End If
            
                lMin = lMin + lCnt
                DoEvents
            Loop While lRet = ERROR_MORE_DATA
        End If
        If hEnum Then l = WNetCloseEnum(hEnum)
        DoEvents
    Loop While lLastIx < lMin

    '### Auswerten wenn nichts da dann exit sub
    If UBound(uNet) + 1 = 0 Then Exit Function
    
    For Läufer = 0 To UBound(uNet)
        Select Case uNet(Läufer).dwDisplayType
          Case RESDTYPE_DIRECTORY:  '= "Ordner"
              
              NetAusgabe = "ordner " & _
                  "A1 " & uNet(Läufer).dwDisplayType & _
                  "A2 " & uNet(Läufer).dwScope & _
                  "A3 " & uNet(Läufer).dwType & _
                  "A4 " & uNet(Läufer).dwUsage & _
                  "A5 " & uNet(Läufer).sComment & _
                  "A6 " & uNet(Läufer).sLocalName & _
                  "A7 " & uNet(Läufer).sProvider & _
                  "A8 " & uNet(Läufer).sRemoteName
              
          Case RESDTYPE_DOMAIN:     '= "Domäne"
              NetAusgabe = "Domäne: " & Trim(uNet(Läufer).sRemoteName)
              
          Case RESDTYPE_FILE:       '= "Datei"
              NetAusgabe = "datei " & _
                  "A1 " & uNet(Läufer).dwDisplayType & _
                  "A2 " & uNet(Läufer).dwScope & _
                  "A3 " & uNet(Läufer).dwType & _
                  "A4 " & uNet(Läufer).dwUsage & _
                  "A5 " & uNet(Läufer).sComment & _
                  "A6 " & uNet(Läufer).sLocalName & _
                  "A7 " & uNet(Läufer).sProvider & _
                  "A8 " & uNet(Läufer).sRemoteName
                  
          Case RESDTYPE_GENERIC:    '= "Generic"
              NetAusgabe = "datei " & _
                  "A1 " & uNet(Läufer).dwDisplayType & _
                  "A2 " & uNet(Läufer).dwScope & _
                  "A3 " & uNet(Läufer).dwType & _
                  "A4 " & uNet(Läufer).dwUsage & _
                  "A5 " & uNet(Läufer).sComment & _
                  "A6 " & uNet(Läufer).sLocalName & _
                  "A7 " & uNet(Läufer).sProvider & _
                  "A8 " & uNet(Läufer).sRemoteName
                  
          Case RESDTYPE_GROUP:      '= "Gruppe"
              NetAusgabe = "Gruppe " & _
                  "A1 " & uNet(Läufer).dwDisplayType & _
                  "A2 " & uNet(Läufer).dwScope & _
                  "A3 " & uNet(Läufer).dwType & _
                  "A4 " & uNet(Läufer).dwUsage & _
                  "A5 " & uNet(Läufer).sComment & _
                  "A6 " & uNet(Läufer).sLocalName & _
                  "A7 " & uNet(Läufer).sProvider & _
                  "A8 " & uNet(Läufer).sRemoteName
                  
          Case RESDTYPE_NETWORK:    '= "Netzwerk"
              NetAusgabe = "datei " & _
                  "A1 " & uNet(Läufer).dwDisplayType & _
                  "A2 " & uNet(Läufer).dwScope & _
                  "A3 " & uNet(Läufer).dwType & _
                  "A4 " & uNet(Läufer).dwUsage & _
                  "A5 " & uNet(Läufer).sComment & _
                  "A6 " & uNet(Läufer).sLocalName & _
                  "A7 " & uNet(Läufer).sProvider & _
                  "A8 " & uNet(Läufer).sRemoteName
                  
          Case RESDTYPE_ROOT:       '= "Root"
              NetAusgabe = "root " & _
                  "A1 " & uNet(Läufer).dwDisplayType & _
                  "A2 " & uNet(Läufer).dwScope & _
                  "A3 " & uNet(Läufer).dwType & _
                  "A4 " & uNet(Läufer).dwUsage & _
                  "A5 " & uNet(Läufer).sComment & _
                  "A6 " & uNet(Läufer).sLocalName & _
                  "A7 " & uNet(Läufer).sProvider & _
                  "A8 " & uNet(Läufer).sRemoteName
                  
          Case RESDTYPE_SERVER:     '= "Rechner"
              NetAusgabe = "Rechner " & uNet(Läufer).sComment & " " & _
                  uNet(Läufer).sRemoteName
    
          Case RESDTYPE_SHARE:      '= "Freigabe"
              NetAusgabe = "freigabe " & uNet(Läufer).sRemoteName
              
          Case RESDTYPE_SHAREADMIN: '= "Freigaben Admin"
              NetAusgabe = "Freigaben Admin " & _
                  "A1 " & uNet(Läufer).dwDisplayType & _
                  "A2 " & uNet(Läufer).dwScope & _
                  "A3 " & uNet(Läufer).dwType & _
                  "A4 " & uNet(Läufer).dwUsage & _
                  "A5 " & uNet(Läufer).sComment & _
                  "A6 " & uNet(Läufer).sLocalName & _
                  "A7 " & uNet(Läufer).sProvider & _
                  "A8 " & uNet(Läufer).sRemoteName
          End Select
          
          'in Liste einsetzen
          If Len(NetAusgabe) > 0 Then
              Form1.Liste.AddItem NetAusgabe
              NetAusgabe = ""
          End If
      Next Läufer
End Function



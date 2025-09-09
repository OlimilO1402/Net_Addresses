VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NetAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type ADAPTER_STATUS
    adapter_address(0 To 5) As Byte 'MAC-Adress
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
    NName(0 To 15) As Byte 'String * NCBNAMSZ
    name_num   As Integer
    name_flags As Integer
End Type

Private Type ASTAT
    adapt        As ADAPTER_STATUS
    NameBuff(30) As NAME_BUFFER
End Type

Private m_AdapterStatus As ASTAT
Private m_MACAddress    As MACAddress

Public Property Get Ptr() As LongPtr
    Ptr = VarPtr(m_AdapterStatus)
End Property

Public Property Get Name() As String
    Name = StrConv(m_AdapterStatus.NameBuff(0).NName, vbNarrow)
End Property

Public Sub Reset()
    
End Sub

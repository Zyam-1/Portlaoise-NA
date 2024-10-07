Attribute VB_Name = "basPing"
Option Explicit

Public Const SOCKET_ERROR = 0

Public Type WSAdata
    wVersion As Long
    wHighVersion As Long
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Long
    iMaxUdpDg As Long
    lpVendorInfo As Long
End Type

Public Type Hostent
    h_name As Long
    h_aliases As Long
    h_addrtype As Long
    h_length As Long
    h_addr_list As Long
End Type

Public Type IP_OPTION_INFORMATION
    TTL As Byte
    Tos As Byte
    Flags As Byte
    OptionsSize As Long
    OptionsData As String * 128
End Type

Public Type IP_ECHO_REPLY
    Address(0 To 3) As Byte
    Status As Long
    RoundTripTime As Long
    DataSize As Long
    Reserved As Long
    data As Long
    Options As IP_OPTION_INFORMATION
End Type

Public Declare Function GetHostByName Lib "wsock32.dll" Alias "gethostbyname" (ByVal HostName As String) As Long
Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired&, lpWSAdata As WSAdata) As Long
Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Public Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal HANDLE As Long) As Boolean
Public Declare Function IcmpSendEcho Lib "ICMP" (ByVal IcmpHandle As Long, ByVal DestAddress As Long, ByVal RequestData As String, ByVal RequestSize As Long, RequestOptns As IP_OPTION_INFORMATION, ReplyBuffer As IP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Boolean

Public Function Ping(ByVal HostName As String) As Boolean
      'Returns True if success

          Dim hFile As Long, lpWSAdata As WSAdata
          Dim hHostent As Hostent, AddrList As Long
          Dim Address As Long, rIP As String
          Dim OptInfo As IP_OPTION_INFORMATION
          Dim EchoReply As IP_ECHO_REPLY

10        On Error GoTo Ping_Error

20        Ping = False

30        Call WSAStartup(&H101, lpWSAdata)

40        If GetHostByName(HostName + String(64 - Len(HostName), 0)) <> SOCKET_ERROR Then
50            CopyMemory hHostent.h_name, ByVal GetHostByName(HostName + String(64 - Len(HostName), 0)), Len(hHostent)
60            CopyMemory AddrList, ByVal hHostent.h_addr_list, 4
70            CopyMemory Address, ByVal AddrList, 4
80        End If

90        hFile = IcmpCreateFile()
100       If hFile = 0 Then
              '    MsgBox "Unable to Create File Handle"
110           Exit Function
120       End If

130       OptInfo.TTL = 255
140       If IcmpSendEcho(hFile, Address, String(32, "A"), 32, OptInfo, EchoReply, Len(EchoReply) + 8, 2000) Then
150           rIP = CStr(EchoReply.Address(0)) + "." + CStr(EchoReply.Address(1)) + "." + CStr(EchoReply.Address(2)) + "." + CStr(EchoReply.Address(3))
160       Else
              '  MsgBox "Timeout"
170           Call IcmpCloseHandle(hFile)
180           Call WSACleanup
190           Exit Function
200       End If
210       If EchoReply.Status = 0 Then
              '    MsgBox "Reply from " + HostName + " (" + rIP + ") recieved after " + Trim$(CStr(EchoReply.RoundTripTime)) + "ms"
220       Else
              '    MsgBox "Failure ..."
230       End If

240       Call IcmpCloseHandle(hFile)
250       Call WSACleanup

260       Ping = True

270       Exit Function

Ping_Error:

          Dim strES As String
          Dim intEL As Integer



280       intEL = Erl
290       strES = Err.Description
300       LogError "basPing", "Ping", intEL, strES


End Function



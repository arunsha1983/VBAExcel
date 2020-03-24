' ###################################################################################
' ### Constants and Type Definitions
' ###
Private Const MIN_SOCKETS_REQD As Long = 1
Private Const WS_VERSION_REQD As Long = &H101
Private Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
Private Const SOCKET_ERROR As Long = -1
Private Const WSADESCRIPTION_LEN = 257
Private Const WSASYS_STATUS_LEN = 129
Private Const MAX_WSADescription = 256
Private Const MAX_WSASYSStatus = 128

Private Const IP_SUCCESS As Long = 0
Private Const IP_BUF_TOO_SMALL As Long = 11001
Private Const IP_DEST_NET_UNREACHABLE As Long = 11002
Private Const IP_DEST_HOST_UNREACHABLE As Long = 11003
Private Const IP_DEST_PROT_UNREACHABLE As Long = 11004
Private Const IP_DEST_PORT_UNREACHABLE As Long = 11005
Private Const IP_NO_RESOURCES As Long = 11006
Private Const IP_BAD_OPTION  As Long = 11007
Private Const IP_HW_ERROR As Long = 11008
Private Const IP_PACKET_TOO_BIG As Long = 11009
Private Const IP_REQ_TIMED_OUT As Long = 11010
Private Const IP_BAD_REQ As Long = 11011
Private Const IP_BAD_ROUTE As Long = 11012
Private Const IP_TTL_EXPIRED_TRANSIT As Long = 11013
Private Const IP_TTL_EXPIRED_REASSEM As Long = 11014
Private Const IP_PARAM_PROBLEM As Long = 11015
Private Const IP_SOURCE_QUENCH As Long = 11016
Private Const IP_OPTION_TOO_BIG As Long = 11017
Private Const IP_BAD_DESTINATION As Long = 11018
Private Const IP_GENERAL_FAILURE As Long = 11050

Private Const IP_FLAG_REVERSE As Long = 1
Private Const IP_FLAG_DF As Long = 2
Private Const IP_FLAG_REVERSE_DF As Long = 3

Private Const AF_UNSPEC = 0
Private Const AF_INET = 2
Private Const AF_NETBIOS = 17
Private Const AF_INET6 = 23

Private Type IN_ADDR
    s_b1 As Byte
    s_b2 As Byte
    s_b3 As Byte
    s_b4 As Byte
End Type

Private Type Hostent
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Private Type IP_OPTION_INFORMATION
    TTL As Byte
    Tos As Byte
    Flags As Byte
    OptionsSize As Long
    OptionsData As String * 128
End Type

Private Type IP_ECHO_REPLY
    Address As IN_ADDR
    Status As Long
    RoundTripTime As Long
    DataSize As Integer
    Reserved As Integer
    data As Long
    Options As IP_OPTION_INFORMATION
End Type

Private Type WSAdata
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To MAX_WSADescription) As Byte
    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
    wMaxSockets As Integer
    wMaxUDPDG As Integer
    dwVendorInfo As Long
End Type

' ###################################################################################
' ### WINSOCK Native Function Imports
' ###
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub GetPointer Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As Any, ByVal Length As Long)
Private Declare Sub GetValue Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function lstrlenA Lib "kernel32" (ByVal Ptr As Any) As Long
Private Declare Function WSACleanup Lib "WSOCK32" () As Long
Private Declare Function WSAStartup Lib "WSOCK32" (ByVal wVersionRequired As Long, lpWSAdata As WSAdata) As Long
Private Declare Function inet_addr Lib "WSOCK32" (ByVal cp As String) As Long
Private Declare Function GetHostByAddr Lib "WSOCK32" Alias "gethostbyaddr" (haddr As Long, ByVal hnlen As Long, ByVal addrtype As Long) As Long
Private Declare Function GetHostByName Lib "WSOCK32" Alias "gethostbyname" (ByVal Hostname As String) As Long
Private Declare Function IcmpCreateFile Lib "icmp" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp" (ByVal HANDLE As Long) As Boolean
Private Declare Function IcmpSendEcho Lib "icmp" ( _
    ByVal IcmpHandle As Long, _
    ByVal DestAddress As Long, ByVal RequestData As String, _
    ByVal RequestSize As Integer, _
    RequestOptns As IP_OPTION_INFORMATION, _
    ReplyBuffer As IP_ECHO_REPLY, _
    ByVal ReplySize As Long, _
    ByVal Timeout As Long _
    ) As Boolean

' ###################################################################################
' ### Private Utility Functions
' ###
Private Function HiByte(ByVal wParam As Integer)
    HiByte = wParam \ &H100 And &HFF&
End Function

Private Function LoByte(ByVal wParam As Integer)
    LoByte = wParam And &HFF&
End Function

Private Sub SocketsCleanup()
    If WSACleanup() <> ERROR_SUCCESS Then
        'Failed to cleanup sockets.
    End If
End Sub

Private Function SocketsInitialize() As Boolean
    Dim WSAD As WSAdata
    Dim sLoByte As String
    Dim sHiByte As String

    If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
        'The 32-bit Windows Socket is not responding.
        SocketsInitialize = False
        Exit Function
    End If

    If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        SocketsInitialize = False
        Exit Function
    End If

    If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
        sHiByte = CStr(HiByte(WSAD.wVersion))
        sLoByte = CStr(LoByte(WSAD.wVersion))
        SocketsInitialize = False
        Exit Function
    End If
    'must be OK, so lets do it
    SocketsInitialize = True
End Function

' ###################################################################################
' ### Exposed Excel Worksheet Functions
' ###
Public Function GetHostName(ByVal Address As String) As String
    Dim lLength As Long
    Dim lRet As Long

    If Not SocketsInitialize() Then
        GetHostName = "WINSOCK_FAILURE"
        Exit Function
    End If

    lRet = GetHostByAddr(inet_addr(Address), 4, AF_INET)
    If lRet <> 0 Then
        CopyMemory lRet, ByVal lRet, 4
        lLength = lstrlenA(lRet)
        If lLength > 0 Then
            GetHostName = Space$(lLength + 1)
            CopyMemory ByVal GetHostName, ByVal lRet, lLength
        End If
    Else
        GetHostName = ""
    End If

    SocketsCleanup
End Function

Public Function GetIpAddress(ByVal Hostname As String) As String
    Dim hFile As Long
    Dim hHostent As Hostent
    Dim lRet As Long
    Dim AddrList As Long
    Dim ptrAddr As Long
    Dim ptrStrIp As Long
    Dim strIpAddress As String
    Dim Addr As IN_ADDR

    If Not SocketsInitialize() Then
        GetIpAddress = "WINSOCK_FAILURE"
        Exit Function
    End If

    'Get Host IP Address by Name
    lRet = GetHostByName(Hostname + String(64 - Len(Hostname), 0))
    If lRet <> SOCKET_ERROR Then

        GetValue hHostent.h_name, GetHostByName(Hostname + String(64 - Len(Hostname), 0)), Len(hHostent)
        GetValue ptrAddr, ByVal hHostent.h_addr_list, 4

        For ptrAddr = ptrAddr To (ptrAddr + lstrlenA(hHostent.h_addr_list) - 1) Step 4
            GetValue Addr, ptrAddr, 4 'Debug: cast as IP Address
            GetIpAddress = GetIpAddress + " " + CStr(Addr.s_b1) + "." _
                                              + CStr(Addr.s_b2) + "." _
                                              + CStr(Addr.s_b3) + "." _
                                              + CStr(Addr.s_b4)
        Next

        GetIpAddress = LTrim(GetIpAddress)
    End If
    SocketsCleanup
End Function

'Returns Error code as String or Round-Trip Response Time in msecs.
Public Function Ping(ByVal Hostname As String, _
                     Optional TTL As Long = 255, _
                     Optional msTimeout As Long = 2000, _
                     Optional packetLength As Long = 32, _
                     Optional DoNotFragment As Boolean = True) As Variant
    Dim hFile As Long
    Dim hHostent As Hostent
    Dim AddrList As Long
    Dim Address As Long
    Dim strIpAddress As String
    Dim OptInfo As IP_OPTION_INFORMATION
    Dim EchoReply As IP_ECHO_REPLY
    Dim StatusString As String
    Dim lRet As Long

    If Not SocketsInitialize() Then
        Ping = "WINSOCK_INIT_FAIL"
        Exit Function
    End If

    'Get Host IP Address by Name
    lRet = GetHostByName(Hostname + String(64 - Len(Hostname), 0))
    If lRet <> SOCKET_ERROR Then
        GetValue hHostent.h_name, lRet, Len(hHostent)
        GetValue AddrList, hHostent.h_addr_list, 4
        GetValue Address, AddrList, 4
    Else
        Ping = "NS_SOCKET_ERROR"
        Exit Function
    End If

    'Attempt to Create File Handle to store response in.
    hFile = IcmpCreateFile()
    If hFile = 0 Then
        Ping = "FILE_HANDLE_FAILURE"
        Call WSACleanup 'Terminate WinSock
        Exit Function
    Else
        'Set Options
        OptInfo.TTL = TTL 'Sets Time-to-Live (Limits Router Hops between host and destination)
        OptInfo.Tos = 0 'Silently ignored by WinSock. See: http://support.microsoft.com/kb/248611
        If DoNotFragment Then
            OptInfo.Flags = IP_FLAG_DF
        Else
            OptInfo.Flags = 0
        End If

        'Send ICMP Echo
        If IcmpSendEcho(hFile, Address, String(packetLength, "A"), packetLength, OptInfo, EchoReply, Len(EchoReply) + 8, msTimeout) Then
            strIpAddress = CStr(EchoReply.Address.s_b1) + "." _
                         + CStr(EchoReply.Address.s_b2) + "." _
                         + CStr(EchoReply.Address.s_b3) + "." _
                         + CStr(EchoReply.Address.s_b4)
        End If

        Select Case EchoReply.Status
                    Case IP_SUCCESS
                'The ICMP echo was delivered successfully with the proper response.
                Ping = EchoReply.RoundTripTime

            Case IP_BUF_TOO_SMALL
                'The reply buffer was too small.
                Ping = strIpAddress & ":BUF_TOO_SMALL"

            Case IP_DEST_NET_UNREACHABLE
                'The destination network was unreachable.
                Ping = strIpAddress & ":DEST_NET_UNREACHABLE"

            Case IP_DEST_HOST_UNREACHABLE
                'The destination host was unreachable.
                Ping = strIpAddress & ":DEST_HOST_UNREACHABLE"

            Case IP_DEST_PROT_UNREACHABLE
                'The destination protocol was unreachable.
                Ping = strIpAddress & ":DEST_PROT_UNREACHABLE"

            Case IP_DEST_PORT_UNREACHABLE
                'The destination port was unreachable.
                Ping = strIpAddress & ":DEST_PORT_UNREACHABLE"

            Case IP_NO_RESOURCES
                'Insufficient IP resources were available.
                Ping = strIpAddress & ":NO_RESOURCES"

            Case IP_BAD_OPTION
                'A bad IP option was specified.
                Ping = strIpAddress & ":BAD_OPTION"

            Case IP_HW_ERROR
                'A hardware error occurred.
                Ping = strIpAddress & ":HW_ERROR"

            Case IP_PACKET_TOO_BIG
                'The packet was too big.
                Ping = strIpAddress & ":PACKET_TOO_BIG"

            Case IP_REQ_TIMED_OUT
                'The request timed out.
                Ping = strIpAddress & ":REQ_TIMED_OUT"

            Case IP_BAD_REQ
                'A bad request.
                Ping = strIpAddress & ":BAD_REQ"

            Case IP_BAD_ROUTE
                'A bad route.
                Ping = strIpAddress & ":BAD_ROUTE"

            Case IP_TTL_EXPIRED_TRANSIT
                'The time to live (TTL) expired in transit.
                Ping = strIpAddress & ":TTL_EXPIRED_TRANSIT"

            Case IP_TTL_EXPIRED_REASSEM
                'The time to live expired during fragment reassembly.
                Ping = strIpAddress & ":TTL_EXPIRED_REASSEM"

            Case IP_PARAM_PROBLEM
                'A parameter problem.
                Ping = strIpAddress & ":PARAM_PROBLEM"

            Case IP_SOURCE_QUENCH
                'Datagrams are arriving too fast to be processed and datagrams may have been discarded.
                Ping = strIpAddress & ":SOURCE_QUENCH"

            Case IP_OPTION_TOO_BIG
                'An IP option was too big.
                Ping = strIpAddress & ":OPTION_TOO_BIG"

            Case IP_BAD_DESTINATION
                'A bad destination.
                Ping = strIpAddress & ":BAD_DESTINATION"

            Case IP_GENERAL_FAILURE
                'A general failure. This error can be returned for some malformed ICMP packets.
                Ping = "GENERAL_FAILURE"

            Case Else
                'An unknown error occured.
                Ping = "UNKNOWN_FAILURE"
        End Select
    End If

    Call IcmpCloseHandle(hFile) 'Close File Handle

    SocketsCleanup
End Function

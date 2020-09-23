VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum enmSoxState
    ' Used
    soxDisconnected = 0&
    soxListening = 1&
    soxConnecting = 2&
    soxConnected = 3&
    soxDataRecv = 4&
    soxDataSend = 5&
    soxClosing = 6&
    soxBound = 10& ' The socket has been bound to its current Port and Address
    soxERROR = -1& ' This is here so the outside calling functions can test function return values eg. If Sox.CloseIt(123) = soxERROR Then ... but is not used within the Class module
    ' Unused
'    soxWaitingAnswer = 7&
'    soxCommandAnswered = 8&
'    soxCommandNotAnswered = 9&
'    soxWaitingBinary = 11&
End Enum

' API States
'    soxDisconnected = 0&
'    soxConnecting = 1&
'    soxConnected = 2&
'    soxWaitingAnswer = 3&
'    soxCommandAnswered = 4&
'    soxCommandNotAnswered = 5&
'    soxDataReceived = 6&
'    soxListening = 7&
'    soxWaitingBinary = 8&
'    soxERROR = -1&

''VB WinSock OCX Defined
'Public Enum enmVBState
'    sckClosed = 0
'    sckOpen = 1
'    sckListening = 2
'    sckConnected = 7
'    sckError = 9

    'sckClosed 0 Default. Closed
    'sckOpen 1 Open
    'sckListening 2 Listening
    'sckConnectionPending  3 Connection pending
    'sckResolvingHost  4 Resolving host
    'sckHostResolved  5 Host resolved
    'sckConnecting  6 Connecting
    'sckConnected  7 Connected
    'sckClosing  8 Peer is closing the connection
    'sckError  9 Error
'End Enum

Public Enum enmSoxOptions
    ' Set & Get Compatible Options
    soxSO_BROADCAST = &H20& 'BOOL Allow transmission of broadcast messages on the socket.
    soxSO_DEBUG = &H1& 'BOOL Record debugging information.
'    soxSO_SO_DONTLINGER = Not soxSO_LINGER 'BOOL Do not block close waiting for unsent data to be sent. Setting this option is equivalent to setting SO_LINGER with l_onoff set to zero.
    soxSO_DONTROUTE = &H10& 'BOOL Do not route: send directly to interface.
    soxSO_KEEPALIVE = &H8& 'BOOL Send keepalives
    soxSO_LINGER = &H80& 'struct LINGER  Linger on close if unsent data is present.
    soxSO_OOBINLINE = &H100& 'BOOL Receive out-of-band data in the normal data stream. (See section DECnet Out-Of-band data for a discussion of this topic.)
    soxSO_RCVBUF = &H1002& 'int Specify the total per-socket buffer space reserved for receives. This is unrelated to SO_MAX_MSG_SIZE or the size of a TCP window.
    soxSO_REUSEADDR = &H4& 'BOOL Allow the socket to be bound to an address that is already in use. (See bind.)
    soxSO_SNDBUF = &H1001& 'int Specify the total per-socket buffer space reserved for sends. This is unrelated to SO_MAX_MSG_SIZE or the size of a TCP window.
    ' Inverting TCP_NODELAY value to create unique value of -2
    soxSO_TCP_NODELAY = Not &H1& 'BOOL Disables the Nagle algorithm for send coalescing.
    ' Get ONLY Compatible Options
    soxSO_USELOOPBACK = &H40& 'bypass hardware when possible
    soxSO_ACCEPTCONN = &H2& 'BOOL Socket is listening.
    soxSO_ERROR = &H1007& 'int Retrieve error status and clear.
    soxSO_TYPE = &H1008& 'Get Socket Type (From FTP - Experimental) (Seems to always returns 1 for a valid TCP socket, -1 for a closed socket)
    ' Unsupported options by either SOX or WinSock
'    soxSO_GROUP_ID 'GROUP The identifier of the group to which this socket belongs.
'    soxSO_GROUP_PRIORITY 'int The relative priority for sockets that are part of a socket group.
'    soxSO_MAX_MSG_SIZE 'unsigned int Maximum size of a message for message-oriented socket types (for example, SOCK_DGRAM). Has no meaning for stream oriented sockets.
'    soxSO_PROTOCOL_INFO 'WSAPROTOCOL_INFO Description of protocol info for protocol that is bound to this socket.
'    soxPVD_CONFIG 'Service Provider Dependent An "opaque" data structure object from the service provider associated with socket s. This object stores the current configuration information of the service provider. The exact format of this data structure is service provider specific.
End Enum

Private Type typPortal 'Class specific variables
    hWnd As Long 'The handle to the window we create on initialization that will receive WinSock messages
    WndProc As Long 'Pointer to the origional WindowProc of our window (We need to give control of ALL messages back to it before we destroy it)
    Sockets As Long 'How many Sockets are comming through the Portal, Actually hold the Socket array count. NB - MUST change with Redim of Sockets
End Type

'API Defined
Private Type typSocketAddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero(7) As Byte
End Type
Private Const SOCKADDR_SIZE = 16

'Class module Defined
Private Type typSocket
    Socket As Long ' The actual WinSock API socket number
    SocketAddr As typSocketAddr ' Info about the connection
    State As enmSoxState ' Not FULLY implemented
    uMsg As Long ' Server (-1) / Client (0) Socket (Server = A Socket that has a connection to the Server / Client = A Socket that was created in Accept that connected to us)
    DataLen As Long
    Data() As Byte
End Type

Private Const WSADESCRIPTION_LEN As Long = 256 '(Confirmed)
Private Const WSASYS_STATUS_LEN As Long = 128 '(Confirmed)

'API Defined
'Contains information about our current WinSock implementation
Private Type typWSAData
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADESCRIPTION_LEN) As Byte
    szSystemStatus(0 To WSASYS_STATUS_LEN) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Type typLinger
    l_onoff As Integer
    l_linger As Integer
End Type

'VB WinSock OCX Defined Error codes
Public Enum enmError
    sckOutOfMemory = 7                'Out of memory
    sckInvalidPropertyValue = 380     'The property value is invalid
    sckGetNotSupported = 394          'The property cannot be read
    sckSetNotSupported = 383          'The property is read-only
    sckBadState = 40006               'Wrong protocol or connection state for the requested transaction or request
    sckInvalidArg = 40014             'The argument passed to a function was not in the correct format or in the specified range
    sckSuccess = 40017                'Successful
    sckUnsupported = 40018            'Unsupported variant type
    sckInvalidOp = 40020              'Invalid operation at current state
    sckOutOfRange = 40021             'Argument is out of range
    sckWrongProtocol = 40026          'Wrong protocol for the requested transaction or request
    sckOpCanceled = 1004              'The operation was canceled
    sckInvalidArgument = 10014        'The requested address is a broadcast address, but flag is not set
    sckWouldBlock = 10035             'Socket is non-blocking and the specified operation will block
    sckInProgress = 10036             'A blocking Winsock operation in progress
    sckAlreadyComplete = 10037        'The operation is completed. No blocking operation in progress
    sckNotSocket = 10038              'The descriptor is not a socket
    sckMsgTooBig = 10040              'The datagram is too large to fit into the buffer and is truncated
    sckPortNotSupported = 10043       'The specified port is not supported
    sckAddressInUse = 10048           'Address in use
    sckAddressNotAvailable = 10049    'Address not available from the local machine
    sckNetworkSubsystemFailed = 10050 'Network subsystem failed
    sckNetworkUnreachable = 10051     'The network cannot be reached from this host at this time
    sckNetReset = 10052               'Connection has timed out when SO_KEEPALIVE is set
    sckConnectAborted = 10053         'Connection is aborted due to timeout or other failure
    sckConnectionReset = 10054        'The connection is reset by remote side
    sckNoBufferSpace = 10055          'No buffer space is available
    sckAlreadyConnected = 10056       'Socket is already connected
    sckNotConnected = 10057           'Socket is not connected
    sckSocketShutdown = 10058         'Socket has been shut down
    sckTimedout = 10060               'Socket has been shut down
    sckConnectionRefused = 10061      'Connection is forcefully rejected
    sckNotInitialized = 10093         'WinsockInit should be called first
    sckHostNotFound = 11001           'Authoritative answer: Host not found
    sckHostNotFoundTryAgain = 11002   'Non-Authoritative answer: Host not found
    sckNonRecoverableError = 11003    'Non-recoverable errors
    sckNoData = 11004                 'Valid name, no data record of requested type
End Enum

'All WinSock error constants are based on WSABASEERR
Private Const WSABASEERR = 10000
'WinSock definitions of regular Microsoft C error constants
Private Const WSAEINTR = (WSABASEERR + 4) 'Interrupted function call
Private Const WSAEBADF = (WSABASEERR + 9)
Private Const WSAEACCES = (WSABASEERR + 13) 'Permission Denied
Private Const WSAEFAULT = (WSABASEERR + 14) 'Bad address
Private Const WSAEINVAL = (WSABASEERR + 22) 'Invalid argument
Private Const WSAEMFILE = (WSABASEERR + 24) 'Too many open files
'Windows Sockets definitions of regular Berkeley error constants
Private Const WSAEWOULDBLOCK = (WSABASEERR + 35) 'Resource temporarily unavailable
Private Const WSAEINPROGRESS = (WSABASEERR + 36) 'Operation now in progress
Private Const WSAEALREADY = (WSABASEERR + 37) 'Operation already in progress
Private Const WSAENOTSOCK = (WSABASEERR + 38) 'Socket operation on non-socket
Private Const WSAEDESTADDRREQ = (WSABASEERR + 39) 'Destination address required
Private Const WSAEMSGSIZE = (WSABASEERR + 40) 'Message too long
Private Const WSAEPROTOTYPE = (WSABASEERR + 41) 'Protocol wrong type for socket
Private Const WSAENOPROTOOPT = (WSABASEERR + 42) 'Bad protocol option
Private Const WSAEPROTONOSUPPORT = (WSABASEERR + 43) 'Protocol not supported
Private Const WSAESOCKTNOSUPPORT = (WSABASEERR + 44) 'Socket type not supported
Private Const WSAEOPNOTSUPP = (WSABASEERR + 45) 'Operation not supported
Private Const WSAEPFNOSUPPORT = (WSABASEERR + 46) 'Protocol family not supported
Private Const WSAEAFNOSUPPORT = (WSABASEERR + 47) 'Address family not supported by protocol family
Private Const WSAEADDRINUSE = (WSABASEERR + 48) 'Address already in use
Private Const WSAEADDRNOTAVAIL = (WSABASEERR + 49) 'Cannot assign requested address
Private Const WSAENETDOWN = (WSABASEERR + 50) 'Network is down
Private Const WSAENETUNREACH = (WSABASEERR + 51) 'Network is unreachable
Private Const WSAENETRESET = (WSABASEERR + 52) 'Network dropped connection on reset
Private Const WSAECONNABORTED = (WSABASEERR + 53) 'Software caused connection abort
Private Const WSAECONNRESET = (WSABASEERR + 54) 'Connection reset by peer
Private Const WSAENOBUFS = (WSABASEERR + 55) 'No buffer space available
Private Const WSAEISCONN = (WSABASEERR + 56) 'Socket is already connected
Private Const WSAENOTCONN = (WSABASEERR + 57) 'Socket is not connected
Private Const WSAESHUTDOWN = (WSABASEERR + 58) 'Cannot send after socket shutdown
Private Const WSAETOOMANYREFS = (WSABASEERR + 59) 'Too many references: can't splice (UnConfirmed Description)
Private Const WSAETIMEDOUT = (WSABASEERR + 60) 'Connection timed out
Private Const WSAECONNREFUSED = (WSABASEERR + 61) 'Connection refused
Private Const WSAELOOP = (WSABASEERR + 62) 'Too many levels of symbolic links (UnConfirmed Description)
Private Const WSAENAMETOOLONG = (WSABASEERR + 63) 'File name too long (UnConfirmed Description)
Private Const WSAEHOSTDOWN = (WSABASEERR + 64) 'Host is down
Private Const WSAEHOSTUNREACH = (WSABASEERR + 65) 'No route to host
Private Const WSAENOTEMPTY = (WSABASEERR + 66) 'Directory not empty (UnConfirmed Description)
Private Const WSAEPROCLIM = (WSABASEERR + 67) 'Too many processes
Private Const WSAEUSERS = (WSABASEERR + 68) 'Too many users (UnConfirmed Description)
Private Const WSAEDQUOT = (WSABASEERR + 69) 'Disk quota exceeded (UnConfirmed Description)
Private Const WSAESTALE = (WSABASEERR + 70) 'Stale NFS file handle (UnConfirmed Description)
Private Const WSAEREMOTE = (WSABASEERR + 71) 'Too many levels of remote in path (UnConfirmed Description)
'Extended Windows Sockets error constant definitions
Private Const WSASYSNOTREADY = (WSABASEERR + 91) 'Network subsystem is unavailable
Private Const WSAVERNOTSUPPORTED = (WSABASEERR + 92) 'WINSOCK.DLL version out of range
Private Const WSANOTINITIALISED = (WSABASEERR + 93) 'Successful WSAStartup not yet performed
Private Const WSAEDISCON1 = (WSABASEERR + 94) 'Graceful shutdown in progress
'Private Const WSA_INVALID_HANDLE    '(OS Dependent) Specified event object handle is invalid
'Private Const WSA_INVALID_PARAMETER '(OS Dependent) One or more parameters are invalid
'Private Const WSAINVALIDPROCTABLE   '(OS Dependent) Invalid procedure table from service provider
'Private Const WSAINVALIDPROVIDER    '(OS Dependent) Invalid service provider version number
'Private Const WSA_IO_INCOMPLETE     '(OS Dependent) Overlapped I/O event object not in signaled state
'Private Const WSA_IO_PENDING        '(OS Dependent) Overlapped operations will complete later
'Private Const WSA_NOT_ENOUGH_MEMORY '(OS Dependent) Insufficient memory available
'Private Const WSAPROVIDERFAILEDINIT '(OS Dependent) Unable to initialize a service provider
'Private Const WSA_OPERATION_ABORTED '(OS Dependent) Overlapped operation aborted
Private Const WSAEDISCON2 = (WSABASEERR + 101) 'Graceful shutdown in progress
Private Const WSAENOMORE = (WSABASEERR + 102)
Private Const WSAECANCELLED = (WSABASEERR + 103)
Private Const WSAEINVALIDPROCTABLE = (WSABASEERR + 104)
Private Const WSAEINVALIDPROVIDER = (WSABASEERR + 105)
Private Const WSAEPROVIDERFAILEDINIT = (WSABASEERR + 106)
Private Const WSASYSCALLFAILURE = (WSABASEERR + 107) '(OS Dependent) System call failure
Private Const WSASERVICE_NOT_FOUND = (WSABASEERR + 108)
Private Const WSATYPE_NOT_FOUND = (WSABASEERR + 109) 'Class type not found
Private Const WSA_E_NO_MORE = (WSABASEERR + 110)
Private Const WSA_E_CANCELLED = (WSABASEERR + 111)
Private Const WSAEREFUSED = (WSABASEERR + 112)
'Authoritative Answer: Host not found
Private Const WSAHOST_NOT_FOUND = (WSABASEERR + 1001) 'Host not found
'Non-Authoritative: Host not found, or SERVERFAIL
Private Const WSATRY_AGAIN = (WSABASEERR + 1002) 'Non-authoritative host not found
'Non recoverable errors, FORMERR, REFUSED, NOTIMP
Private Const WSANO_RECOVERY = (WSABASEERR + 1003) 'This is a non-recoverable error
'Valid name, no data record of requested type
Private Const WSANO_DATA = (WSABASEERR + 1004) 'Valid name, no data record of requested type
'Code snippet of 2 other error codes ... but I don't need them cause they'll be logged under code 11004
'/* no address, look for MX record */
'#define WSANO_ADDRESS           WSANO_DATA
'#define NO_ADDRESS              WSANO_ADDRESS



Private Const INVALID_SOCKET = -1& ' Indication of an Invalid Socket
Private Const SOCKET_ERROR = -1&

Private Const INADDR_ANY = &H0 'Used for auto bind of a socket / selects an unused socket (Confirmed)
Private Const INADDR_NONE = &HFFFFFFFF 'Was FFFF (Confirmed) ... Returned address is an error

'Confirmed Address Families
'Currently all AF values translate directly into their coresponding Protocol family
'eg. AF_INET = PF_INET ... I never added the Protocol Family values as this is confusing enough
'Just remember, if you see PF_INIT in other code ... it's the same as AF_INIT !!!
Private Const AF_UNSPEC = 0 'unspecified
Private Const AF_UNIX = 1 'local to host (pipes, portals)
Private Const AF_INET = 2 'internetwork: UDP, TCP, etc
Private Const AF_IMPLINK = 3 'arpanet imp addresses
Private Const AF_PUP = 4 'pup protocols: e.g. BSP
Private Const AF_CHAOS = 5 'mit CHAOS protocols
Private Const AF_NS = 6 'XEROX NS protocols
Private Const AF_ISO = 7 'ISO protocols
Private Const AF_OSI = AF_ISO 'OSI is ISO
Private Const AF_ECMA = 8 'european computer manufacturers
Private Const AF_DATAKIT = 9 'datakit protocols
Private Const AF_CCITT = 10 'CCITT protocols, X.25 etc
Private Const AF_SNA = 11 'IBM SNA
Private Const AF_DECnet = 12 'DECnet
Private Const AF_DLI = 13 'Direct data link interface
Private Const AF_LAT = 14 'LAT
Private Const AF_HYLINK = 15 'NSC Hyperchannel
Private Const AF_APPLETALK = 16 'AppleTalk
Private Const AF_NETBIOS = 17 'NetBios-style addresses
Private Const AF_MAX = 18

'Confirmed - Maximum queue length specifiable by listen
Private Const SOMAXCONN = 5
'Confirmed flags for recv
'To extract the data from recv - use 0 for flags value (default)
Private Const MSG_OOB = &H1                     'Process out-of-band data
Private Const MSG_PEEK = &H2                    'Peek at incoming message (Probably the only one used !!!)
Private Const MSG_DONTROUTE = &H4               'Send without using routing tables

'Confirmed types
Private Const SOCK_STREAM = 1    'stream socket
Private Const SOCK_DGRAM = 2     'datagram socket
Private Const SOCK_RAW = 3       'raw-protocol interface
Private Const SOCK_RDM = 4       'reliably-delivered message
Private Const SOCK_SEQPACKET = 5 'sequenced packet stream

'Confirmed option level flags (per-socket)
Private Const SOL_SOCKET = &HFFFF& 'Officially the only option for socket level
'Confirmed option flags (per-socket)
Private Const SO_DEBUG = &H1&          'turn on debugging info recording
Private Const SO_ACCEPTCONN = &H2&     'socket has had listen()
Private Const SO_REUSEADDR = &H4&     'allow local address reuse
Private Const SO_KEEPALIVE = &H8&    'keep connections alive (VERY important for future use)
Private Const SO_DONTROUTE = &H10&     'just use interface addresses
Private Const SO_BROADCAST = &H20&    'permit sending of broadcast msgs
Private Const SO_USELOOPBACK = &H40& 'bypass hardware when possible
Private Const SO_LINGER = &H80&      'linger on close if data present
'Private Const SO_DONTLINGER (u_int)(~SO_LINGER)'??? What does this C code mean ??? I believe that this is not really an option for write ... but displays the inverse of SO_LINGER
Private Const SO_OOBINLINE = &H100&    'leave received OOB data in line

'Confirmed additional sock options used by getsockopt API (SO = SockOption)
Private Const SO_SNDBUF = &H1001&    'send buffer size
Private Const SO_RCVBUF = &H1002&    'receive buffer size
Private Const SO_SNDLOWAT = &H1003& 'send low-water mark
Private Const SO_RCVLOWAT = &H1004&  'receive low-water mark
Private Const SO_SNDTIMEO = &H1005&  'send timeout
Private Const SO_RCVTIMEO = &H1006&  'receive timeout
Private Const SO_ERROR = &H1007&    'get error status and clear (Use THIS in stead of WSAGetLastError to return the Socket specific error)
Private Const SO_TYPE = &H1008&     'get socket type

'Confirmed TCP Options
Private Const TCP_NODELAY = &H1

'Confirmed flags to be used with the WSAAsyncSelect() call and on Msg arrival
Private Const FD_READ = &H1
Private Const FD_WRITE = &H2
Private Const FD_OOB = &H4
Private Const FD_ACCEPT = &H8
Private Const FD_CONNECT = &H10
Private Const FD_CLOSE = &H20

' Confirmed ShutDown options
Private Const SD_RECEIVE = &H0
Private Const SD_SEND = &H1
Private Const SD_BOTH = &H2


'Confirmed list of Protocols for use by Socket API call
Private Const IPPROTO_IP = 0 'dummy for IP
Private Const IPPROTO_ICMP = 1 'control message protocol
Private Const IPPROTO_GGP = 2 'gateway^2 (deprecated)
Private Const IPPROTO_TCP = 6 'tcp
Private Const IPPROTO_PUP = 12 'pup
Private Const IPPROTO_UDP = 17 'user datagram protocol
Private Const IPPROTO_IDP = 22 'xns idp
Private Const IPPROTO_ND = 77 'UNOFFICIAL net disk proto
Private Const IPPROTO_RAW = 255 'raw IP packet
Private Const IPPROTO_MAX = 256



'Private Const GWL_EXSTYLE = (-20)
'Private Const GWL_STYLE = (-16)
Private Const GWL_WNDPROC = (-4)
'Private Const GWL_HINSTANCE = (-6)
'Private Const GWL_ID = (-12)
'Private Const GWL_USERDATA = (-21)
'Private Const DWL_DLGPROC = 4 'Used for dialog boxes
'Private Const DWL_MSGRESULT = 0 'Used for dialog boxes
'Private Const DWL_USER = 8 'Used for dialog boxes
'Public Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long


'Public Const WINSOCK_MESSAGE = 4025 'The only Message type currently used ... not anymore :)
Private Const soxSERVER = 4026& ' This indicates that the Socket is either a Listening Socket, or was created from a Listening Socket, either way, our machine is acting as a Sox Server
Private Const soxCLIENT = 4027& ' This indicates that the Socket is a connection we established to another computer/server, therefore our machine is acting as a Sox Client on this Socket

'Limitations - Only allows one hook in App for use by WinSock
'Solution - Create an array to store the various hooks in our app (Why you would want this anyway is beyond me)

Private Declare Function apiWSAStartup Lib "WS2_32" Alias "WSAStartup" (ByVal wVersionRequired As Long, lpWSADATA As typWSAData) As Long
Private Declare Function apiWSACleanup Lib "WS2_32" Alias "WSACleanup" () As Long
Private Declare Function apiSocket Lib "WS2_32" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
Private Declare Function apiCloseSocket Lib "WS2_32" Alias "closesocket" (ByVal s As Long) As Long
Private Declare Function apiBind Lib "WS2_32" Alias "bind" (ByVal s As Long, addr As typSocketAddr, ByVal namelen As Long) As Long
Private Declare Function apiListen Lib "WS2_32" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Private Declare Function apiConnect Lib "WS2_32" Alias "connect" (ByVal s As Long, name As typSocketAddr, ByVal namelen As Long) As Long
Private Declare Function apiAccept Lib "WS2_32" Alias "accept" (ByVal s As Long, addr As typSocketAddr, addrlen As Long) As Long
Private Declare Function apiWSAAsyncSelect Lib "WS2_32" Alias "WSAAsyncSelect" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Private Declare Function apiRecv Lib "WS2_32" Alias "recv" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function apiSend Lib "WS2_32" Alias "send" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function apiGetSockOpt Lib "WS2_32" Alias "getsockopt" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Private Declare Function apiSetSockOpt Lib "WS2_32" Alias "setsockopt" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare Function apiHToNL Lib "WS2_32" Alias "htonl" (ByVal hostlong As Long) As Long 'Host To Network Long
Private Declare Function apiHToNS Lib "WS2_32" Alias "htons" (ByVal hostshort As Long) As Integer 'Host To Network Short
Private Declare Function apiNToHL Lib "WS2_32" Alias "ntohl" (ByVal netlong As Long) As Long 'Network To Host Long
Private Declare Function apiNToHS Lib "WS2_32" Alias "ntohs" (ByVal netshort As Long) As Integer 'Network To Host Short
Private Declare Function apiIPToNL Lib "WS2_32" Alias "inet_addr" (ByVal cp As String) As Long
Private Declare Function apiNLToIP Lib "WS2_32" Alias "inet_ntoa" (ByVal inn As Long) As Long
Private Declare Function apiGetHostName Lib "WS2_32" Alias "gethostname" (ByVal name As String, ByVal namelen As Long) As Long
Private Declare Function apiShutDown Lib "WS2_32" Alias "shutdown" (ByVal s As Long, ByVal how As Long) As Long


Private Declare Function apiCreateWindowEx Lib "User32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function apiDestroyWindow Lib "User32" Alias "DestroyWindow" (ByVal hWnd As Long) As Long
Private Declare Function apiCallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function apiSetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function apiLStrLen Lib "Kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function apiLstrCpy Lib "Kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Sub apiCopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)


Private WSAData As typWSAData 'Stores WinSock data on initialization of WinSock 2
Private Portal As typPortal
Private Sockets() As typSocket


Private Sub Class_Initialize()
    'Let Server.Port = 1234 'Could force a specific port as default other than 0
    If apiWSAStartup(&H101, WSAData) = SOCKET_ERROR Then
        Call MsgBox("WinSock failed to initialize properly - Error#: " & Err.LastDllError, vbApplicationModal + vbCritical, "Critical Error")  'Creates an 'application instance' and memory space in the WinSock DLL (MUST be cleaned up later)
    Else
        Let Portal.hWnd = apiCreateWindowEx(0&, "STATIC", "Portal", 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&) 'Create a hidden object to accept our WinSock messages
        If Portal.hWnd = 0 Then
            Call MsgBox("Error: " & Err.LastDllError & " on Portal creation.", vbApplicationModal + vbCritical, "Critical Error")   'If cleanup failed, does not / cannot raise errors
        Else
            Let Portal.WndProc = apiSetWindowLong(Portal.hWnd, GWL_WNDPROC, AddressOf WindowProc)
            Let Portal.Sockets = -1 ' Initialize our socket count ... NB - WE HAVE NONE, used primarily to Redim the Sockets Array
        End If
    End If
End Sub

Private Sub Class_Terminate() 'Scaled down version of Terminate code, used purely as failsafe, normally just call Terminate Sub above
    Dim tmpSox As Long
    For tmpSox = 0 To Portal.Sockets
        Call apiShutDown(Sockets(tmpSox).Socket, SD_BOTH)
        Call apiCloseSocket(Sockets(tmpSox).Socket)
    Next tmpSox
    'Correctly replaces/reattaches the origional WindowProc procedure to our 'hidden' handle
    Call apiSetWindowLong(Portal.hWnd, GWL_WNDPROC, Portal.WndProc)
    'This will destroy our hidden object that received all our WinSock API messages
    If apiDestroyWindow(Portal.hWnd) = 0 Then Call MsgBox("Error#: " & Err.LastDllError & " on Portal destruction", vbApplicationModal + vbCritical, "Critical Error")   'If cleanup failed, does not / cannot raise errors
    If apiWSACleanup = SOCKET_ERROR Then Call MsgBox("WinSock failed to terminate properly, memory leak imminent - Error#: " & Err.LastDllError, vbApplicationModal + vbCritical, "Critical Error")   'If cleanup failed, does not / cannot raise errors
End Sub

Private Function Accept(inSocket As Long) As Long 'Returns: New Sox Number -- inSocket is the listening WinSocket ...
    Const Procedure As String = "Accept"
    Dim tmpSocket As Long
    Dim tmpSocketAddr As typSocketAddr 'This stores the details of our new socket/client, including the client IP address
    Let tmpSocket = apiAccept(inSocket, tmpSocketAddr, SOCKADDR_SIZE) 'Accept API returns a valid, random, unused socket for us to use for the new client
    If tmpSocket = INVALID_SOCKET Then 'Accept API may not give us a valid socket eg. when all sockets are full, you may have to add additional error trapping if you believe you will use over 32,767 sockets
        'Since a socket was not commited for the new Connection ... we don't have to close it (Since the socket was never even created)
        Call RaiseError(INVALID_SOCKET, Err.LastDllError, Procedure, "On Accept")
        Let Accept = INVALID_SOCKET
    Else ' Success, A new connection ... Accept now contains the new Socket number
        For Accept = 0 To Portal.Sockets ' First search to see if the socket already exists
            If Sockets(Accept).Socket = tmpSocket Then Exit For
        Next Accept
        If Accept = Portal.Sockets + 1 Then ' If we haven't found an address (Hopefully the only case), Search for an open slot in the array
            For Accept = 0 To Portal.Sockets ' First search to see if the socket already exists
                If Sockets(Accept).Socket = soxDisconnected Then Exit For ' Found an open Socket
            Next Accept
            If Accept = Portal.Sockets + 1 Then ' If we haven't found an address (Hopefully the only case), Search for an open slot in the array
                ReDim Preserve Sockets(Accept) As typSocket
                Let Portal.Sockets = Accept
            End If
        End If
        Let Sockets(Accept).Socket = tmpSocket
        Let Sockets(Accept).SocketAddr = tmpSocketAddr 'Set the details of the new socket/client
        Let Sockets(Accept).uMsg = soxSERVER  'This is a Client Socket - It has connected to US
        Let Sockets(Accept).DataLen = -1
        Erase Sockets(Accept).Data
        Call RaiseState(Accept, soxConnecting) ' Could possibly leave this on soxDisconnected, and on Select Case State, thurn it on and set it ready to send data (Or set it to connecting)
        Call RaiseStatus(Accept, Procedure, "New Connection")
        Call Sox_Connect(Accept)
    End If
End Function

Public Function Bind(LocalPort As Integer, LocalIP As String) As Long
    Const Procedure As String = "Bind"
    Dim tmpSocket As Long
    Dim tmpSocketAddr As typSocketAddr
    If LocalPort = 0 Or LocalIP = vbNullString Then
        Call RaiseError(INVALID_SOCKET, 0, Procedure, "Invalid Port or IP address")
        Let Bind = INVALID_SOCKET
    Else
        Let tmpSocketAddr.sin_family = AF_INET
        Let tmpSocketAddr.sin_port = apiHToNS(LocalPort)
        If tmpSocketAddr.sin_port = INVALID_SOCKET Then
            Call RaiseError(INVALID_SOCKET, Err.LastDllError, Procedure, "Invalid Port: " & tmpSocketAddr.sin_port)
            Let Bind = INVALID_SOCKET
        Else
            Let tmpSocketAddr.sin_addr = apiIPToNL(LocalIP) 'If this is Zero, it will assign 0.0.0.0 !!!
            If tmpSocketAddr.sin_addr = INADDR_NONE Then 'If 255.255.255.255 is returned ... we have a problem ... I think :)
                Call RaiseError(INVALID_SOCKET, Err.LastDllError, Procedure, "Invalid NL Address: " & tmpSocketAddr.sin_addr) ' NL = Network Long
                Let Bind = INVALID_SOCKET
            Else
                Let tmpSocket = apiSocket(AF_INET, SOCK_STREAM, IPPROTO_TCP) 'This is where you specify what type of protocol to use and what type of Streaming to use, returns a new socket number 4 us (NB - From here, if any further steps fail after this one succeeds, we must close the socket)
                If tmpSocket = INVALID_SOCKET Then
                    Call RaiseError(INVALID_SOCKET, Err.LastDllError, Procedure, "On Creation")
                    Let Bind = INVALID_SOCKET
                Else
                    If apiBind(tmpSocket, tmpSocketAddr, SOCKADDR_SIZE) = SOCKET_ERROR Then 'Socket Number, Socket Address space / Name, Name Length ...
                        Call RaiseError(SOCKET_ERROR, Err.LastDllError, Procedure, "On Bind")
                        If apiCloseSocket(tmpSocket) = SOCKET_ERROR Then Call RaiseError(SOCKET_ERROR, Err.LastDllError, Procedure, "On Close") ' We MUST start closing the Socket handle from this point (Unless we store the number and force WinSock to use it later ... Nah ... too much codeing :)))
                        Let Bind = SOCKET_ERROR
                    Else
                        For Bind = 0 To Portal.Sockets ' First search to see if the socket already exists
                            If Sockets(Bind).Socket = tmpSocket Then Exit For
                        Next Bind
                        If Bind = Portal.Sockets + 1 Then ' If we haven't found an address (Hopefully the only case), Search for an open slot in the array
                            For Bind = 0 To Portal.Sockets ' First search to see if the socket already exists
                                If Sockets(Bind).Socket = soxDisconnected Then Exit For ' Found an open Socket
                            Next Bind
                            If Bind = Portal.Sockets + 1 Then ' If we haven't found an address (Hopefully the only case), Search for an open slot in the array
                                ReDim Preserve Sockets(Bind) As typSocket
                                Let Portal.Sockets = Bind
                            End If
                        End If
                        Let Sockets(Bind).Socket = tmpSocket
                        Let Sockets(Bind).SocketAddr = tmpSocketAddr 'Set the details of the new socket/client
                        Call RaiseState(Bind, soxBound)
                        Call RaiseStatus(Bind, Procedure, "Bound")
                    End If
                End If
            End If
        End If
    End If
End Function

'At the moment, we are closing all the sockets in their respectful areas, eg. Connect closes it's own sockets, we may in future need to close all sockets here (Cannot do this as Async will fail and the socket will not be closed properly, Create another private Function to do this)
Public Function CloseIt(inSox As Long) As Long 'OCX Returns # of errors in collection, so should we :)))
    Const Procedure As String = "CloseIt"
    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Call RaiseError(INVALID_SOCKET, 0, Procedure, "Invalid Sox: " & inSox)
        Let CloseIt = INVALID_SOCKET
    Else
        If apiGetSockOpt(Sockets(inSox).Socket, SOL_SOCKET, soxSO_ERROR, CloseIt, 4) = SOCKET_ERROR Then
            Call RaiseError(inSox, Err.LastDllError, Procedure, "On GetSockOpt")
            Let CloseIt = SOCKET_ERROR
        Else
            While Sockets(inSox).State = soxDataRecv Or Sockets(inSox).State = soxDataSend
                DoEvents
            Wend
            If apiShutDown(Sockets(inSox).Socket, SD_SEND) = SOCKET_ERROR Then
                Call RaiseError(inSox, Err.LastDllError, Procedure, "On Shutdown")
                Let CloseIt = SOCKET_ERROR
            Else
                Call RaiseState(inSox, soxClosing)
            End If
        End If
    End If
End Function

Public Function Connect(Optional RemoteHost As String, Optional RemotePort As Integer) As Long 'Returns the new Sox Number / SOCKET_ERROR On Error
    Const Procedure As String = "Connect"
    Dim tmpSocket As Long
    Dim tmpSocketAddr As typSocketAddr
    Let tmpSocketAddr.sin_family = AF_INET
    Let tmpSocketAddr.sin_port = apiHToNS(RemotePort) ' apiHToNS(RemotePort)
    If tmpSocketAddr.sin_port = INVALID_SOCKET Then
        Call RaiseError(INVALID_SOCKET, Err.LastDllError, Procedure, "Port: " & tmpSocketAddr.sin_port)
        Let Connect = INVALID_SOCKET
    Else
        Let tmpSocketAddr.sin_addr = apiIPToNL(RemoteHost) 'If this is Zero, it will assign 0.0.0.0 !!!
        If tmpSocketAddr.sin_addr = INADDR_NONE Then 'If 255.255.255.255 is returned ... we have a problem ... I think :)
            Call RaiseError(INVALID_SOCKET, Err.LastDllError, Procedure, "NL Address: " & tmpSocketAddr.sin_addr) ' NL = Network Long
            Let Connect = INVALID_SOCKET
        Else
            Let tmpSocket = apiSocket(AF_INET, SOCK_STREAM, IPPROTO_TCP) 'This is where you specify what type of protocol to use and what type of Streaming to use, returns a new socket number 4 us (NB - From here, if any further steps fail after this one succeeds, we must close the socket)
            If tmpSocket = INVALID_SOCKET Then
                Call RaiseError(INVALID_SOCKET, Err.LastDllError, Procedure, "On Socket Creation")
                Let Connect = INVALID_SOCKET
            Else
                If apiConnect(tmpSocket, tmpSocketAddr, SOCKADDR_SIZE) = SOCKET_ERROR Then
                    Call RaiseError(SOCKET_ERROR, Err.LastDllError, Procedure, "On Connect")
                    If apiCloseSocket(tmpSocket) = SOCKET_ERROR Then Call RaiseError(SOCKET_ERROR, Err.LastDllError, Procedure, "On Close") ' We MUST start closing the Socket handle from this point (Unless we store the number and force WinSock to use it later ... Nah ... too much codeing :)))
                    Let Connect = SOCKET_ERROR
                Else
                    If apiWSAAsyncSelect(tmpSocket, Portal.hWnd, ByVal soxCLIENT, ByVal FD_ACCEPT Or FD_CLOSE Or FD_CONNECT Or FD_READ Or FD_WRITE) = SOCKET_ERROR Then ' Reassign this Socket to Send and Receive on the DATA channel
                        Call RaiseError(SOCKET_ERROR, Err.LastDllError, Procedure, "On AsyncSelect")
                        If apiCloseSocket(tmpSocket) = SOCKET_ERROR Then Call RaiseError(SOCKET_ERROR, Err.LastDllError, Procedure, "On Close")
                        Let Connect = SOCKET_ERROR
                    Else
                        For Connect = 0 To Portal.Sockets ' First search to see if the socket already exists
                            If Sockets(Connect).Socket = tmpSocket Then Exit For
                        Next Connect
                        If Connect = Portal.Sockets + 1 Then ' If we haven't found an address (Hopefully the only case), Search for an open slot in the array
                            For Connect = 0 To Portal.Sockets ' First search to see if the socket already exists
                                If Sockets(Connect).Socket = soxDisconnected Then Exit For ' Found an open Socket
                            Next Connect
                            If Connect = Portal.Sockets + 1 Then ' If we haven't found an address (Hopefully the only case), Search for an open slot in the array
                                ReDim Preserve Sockets(Connect) As typSocket
                                Let Portal.Sockets = Connect
                            End If
                        End If
                        Let Sockets(Connect).Socket = tmpSocket
                        Let Sockets(Connect).SocketAddr = tmpSocketAddr 'Set the details of the new socket/client
                        Let Sockets(Connect).uMsg = soxCLIENT ' This is a Server connection - We have connected to it (Could even be another Client computer but the fact is we connected to it)
                        Let Sockets(Connect).DataLen = -1
                        Erase Sockets(Connect).Data
                        Call RaiseState(Connect, soxConnecting)
                        Call RaiseStatus(Connect, Procedure, "Connecting")
                    End If
                End If
            End If
        End If
    End If
End Function

Private Sub GetData(inSox As Long) ' Extracts data from the WinSock Recv buffers and places it in our local buffer (data() array)
    Const Procedure As String = "GetData"
    Dim tmpRecv As Long 'Holds how much data we actually received
    Dim tmpBuffer(0 To 8191) As Byte 'This buffer could be optimized for small data, eg. A chat program, if you set it's size, to say 255 (256 in total), it could retrieve data faster
'    Dim tmpLen As Long ' Represents the total amount of Data in the tmpData array!
'    Dim tmpData() As Byte ' Holds the actual data we will return to the client
    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Call RaiseError(INVALID_SOCKET, 0, Procedure, "Invalid Sox: " & inSox)
    Else
        ' First we will disable further notification of FD_READ, because if we extract data with the Recv function, WinSock API posts ANOTHER FD_READ notification to say there's more ...
        ' This is a valid (dare I say recommended) procedure according to WinSock API documentation on MSDN
        If apiWSAAsyncSelect(Sockets(inSox).Socket, Portal.hWnd, ByVal Sockets(inSox).uMsg, 0&) = SOCKET_ERROR Then  ' Reassign this Socket to Send and Receive on the DATA channel
            Call RaiseError(inSox, Err.LastDllError, Procedure, "On AsyncSelect")
        Else
            If Sockets(inSox).State = soxConnected Then Call RaiseState(inSox, soxDataRecv)
            Do
                DoEvents
                Let tmpRecv = apiRecv(Sockets(inSox).Socket, tmpBuffer(0), 8192, 0) ' NB: If U changed the size of tmpBuffer, you will have to make the relative adjustment to the 256 value, also the value at the end of the Do Loop
                Select Case tmpRecv
                    Case SOCKET_ERROR ' Houston, we have a problem :)))
                        Call RaiseError(inSox, Err.LastDllError, Procedure, "On Recv") ' If we send the exact same amount of data as the buffer, we get error 10035 = Resource temporarily unavailable ... not 0
                    Case 0 ' The Socket was Gracefully closed
        '                Call RaiseState(inSox, soxDisconnected) ' Should I set it to disconnected here or not ???
                        Call RaiseStatus(inSox, Procedure, "Gracefully Closed")
                    Case Else
                        ReDim Preserve Sockets(inSox).Data(Sockets(inSox).DataLen + tmpRecv)
                        Call apiCopyMemory(Sockets(inSox).Data(Sockets(inSox).DataLen + 1), tmpBuffer(0), tmpRecv)
                        Let Sockets(inSox).DataLen = Sockets(inSox).DataLen + tmpRecv
                End Select
            Loop While tmpRecv = 8192 ' If tmpRecv = 8192, the buffer was full, and we potentially have more data, if only 8192 bytes were sent, then we on Recv, we will have Error 10035 'Resource temporarily unavailable'
            Do While Sockets(inSox).DataLen > 2
                If Sockets(inSox).DataLen - 3 > DataSize(inSox) Then
                    Call Sox_DataArrival(inSox, ExtractData(inSox))
                Else
                    Exit Do
                End If
            Loop
            If Sockets(inSox).State = soxDataRecv Then Call RaiseState(inSox, soxConnected) ' Could also be closing
            If apiWSAAsyncSelect(Sockets(inSox).Socket, Portal.hWnd, ByVal Sockets(inSox).uMsg, ByVal FD_CLOSE Or FD_READ Or FD_WRITE) = SOCKET_ERROR Then   ' Reassign this Socket to Send and Receive on the DATA channel
                Call RaiseError(inSox, Err.LastDllError, Procedure, "On AsyncSelect")
            End If
        End If
    End If
End Sub

'Creates a socket and sets it in listen mode. This method works only for TCP connections
Public Function Listen(inAddress As String, inPort As Integer) As Long   'Returns Sox number / SOCKET_ERROR On Error
    Const Procedure As String = "Listen"
    Dim tmpSocket As Long
    Dim tmpSocketAddr As typSocketAddr
    Let tmpSocketAddr.sin_family = AF_INET
    Let tmpSocketAddr.sin_port = apiHToNS(inPort)
    If tmpSocketAddr.sin_port = INVALID_SOCKET Then
        Call RaiseError(INVALID_SOCKET, Err.LastDllError, Procedure, "Invalid Port: " & tmpSocketAddr.sin_port)
        Let Listen = INVALID_SOCKET
    Else
        Let tmpSocketAddr.sin_addr = apiIPToNL(inAddress) 'If this is Zero, it will assign 0.0.0.0 !!!
        If tmpSocketAddr.sin_addr = INADDR_NONE Then 'If 255.255.255.255 is returned ... we have a problem ... I think :)
            Call RaiseError(INVALID_SOCKET, Err.LastDllError, Procedure, "Invalid NL Address: " & tmpSocketAddr.sin_addr) ' NL = Network Long
            Let Listen = INVALID_SOCKET
        Else
            Let tmpSocket = apiSocket(AF_INET, SOCK_STREAM, IPPROTO_TCP) 'This is where you specify what type of protocol to use and what type of Streaming to use, returns a new socket number 4 us (NB - From here, if any further steps fail after this one succeeds, we must close the socket)
            If tmpSocket = INVALID_SOCKET Then
                Call RaiseError(INVALID_SOCKET, Err.LastDllError, Procedure, "On Creation")
                Let Listen = INVALID_SOCKET
            Else
                If apiBind(tmpSocket, tmpSocketAddr, SOCKADDR_SIZE) = SOCKET_ERROR Then 'Socket Number, Socket Address space / Name, Name Length ...
                    Call RaiseError(SOCKET_ERROR, Err.LastDllError, Procedure, "On Bind")
                    If apiCloseSocket(tmpSocket) = SOCKET_ERROR Then Call RaiseError(SOCKET_ERROR, Err.LastDllError, Procedure, "On Close") ' We MUST start closing the Socket handle from this point (Unless we store the number and force WinSock to use it later ... Nah ... too much codeing :)))
                    Let Listen = SOCKET_ERROR
                Else
                    If apiListen(ByVal tmpSocket, ByVal 5) = SOCKET_ERROR Then ' 5 = Maximum connections
                        Call RaiseError(SOCKET_ERROR, Err.LastDllError, Procedure, "On Listen")
                        If apiCloseSocket(tmpSocket) = SOCKET_ERROR Then Call RaiseError(SOCKET_ERROR, Err.LastDllError, Procedure, "On Close") ' We MUST start closing the Socket handle from this point (Unless we store the number and force WinSock to use it later ... Nah ... too much codeing :)))
                        Let Listen = SOCKET_ERROR
                    Else
                        If apiWSAAsyncSelect(tmpSocket, Portal.hWnd, ByVal soxSERVER, ByVal FD_ACCEPT Or FD_CLOSE Or FD_CONNECT Or FD_READ Or FD_WRITE) = SOCKET_ERROR Then ' Reassign this Socket to Send and Receive on the DATA channel
                            Call RaiseError(SOCKET_ERROR, Err.LastDllError, Procedure, "On AsyncSelect")
                            If apiCloseSocket(tmpSocket) = SOCKET_ERROR Then Call RaiseError(SOCKET_ERROR, Err.LastDllError, Procedure, "On Close")
                            Let Listen = SOCKET_ERROR
                        Else
                            For Listen = 0 To Portal.Sockets ' First search to see if the socket already exists
                                If Sockets(Listen).Socket = tmpSocket Then Exit For
                            Next Listen
                            If Listen = Portal.Sockets + 1 Then ' If we haven't found an address (Hopefully the only case), Search for an open slot in the array
                                For Listen = 0 To Portal.Sockets ' First search to see if the socket already exists
                                    If Sockets(Listen).Socket = soxDisconnected Then Exit For ' Found an open Socket
                                Next Listen
                                If Listen = Portal.Sockets + 1 Then ' If we haven't found an address (Hopefully the only case), Search for an open slot in the array
                                    ReDim Preserve Sockets(Listen) As typSocket
                                    Let Portal.Sockets = Listen
                                End If
                            End If
                            Let Sockets(Listen).Socket = tmpSocket
                            Let Sockets(Listen).SocketAddr = tmpSocketAddr 'Set the details of the new socket/client
                            Let Sockets(Listen).uMsg = soxSERVER
                            Let Sockets(Listen).DataLen = -1
                            Erase Sockets(Listen).Data
                            Call RaiseState(Listen, soxListening)
                            Call RaiseStatus(Listen, Procedure, "Listening")
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

'When a UNICODE string is passed in, it is converted to an ANSI string before being sent out on the network (Hint - All data eventually get converted to Byte Arrays before being sent, therefore this is the most efficient data type)
Public Function SendData(inSox As Long, inData As Variant) As Long  'Data to be sent. For binary data, byte array should be used (for optimal performace, change inData to a byte array and only allow that datatype to be sent)
    Const Procedure As String = "SendData"
    Dim tmpData() As Byte
    Dim tmpFinal() As Byte
    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Call RaiseError(INVALID_SOCKET, 0, Procedure, "Invalid Sox: " & inSox)
        Let SendData = INVALID_SOCKET
    Else
        While Sockets(inSox).State = soxDataSend
            DoEvents
        Wend
        If Sockets(inSox).State = soxConnected Then Call RaiseState(inSox, soxDataSend)
        Select Case VarType(inData)
            Case 8209   'byte array - Optimal Data Type
                Let tmpData = inData
            Case 8      'string, if we recieve a string, its assumed we are linemode
                Let tmpData = StrConv(inData, vbFromUnicode)
            Case Else
                Let tmpData = StrConv(CStr(inData), vbFromUnicode)
        End Select
        ReDim tmpFinal(3) As Byte
        Call Long2Byte2(UBound(tmpData), tmpFinal)
        ReDim Preserve tmpFinal(UBound(tmpData) + 4) As Byte
        Call apiCopyMemory(tmpFinal(4), tmpData(0), UBound(tmpData) + 1)
        Select Case Sockets(inSox).State
            Case soxDataSend
                DoEvents
                If apiSend(Sockets(inSox).Socket, tmpFinal(0), UBound(tmpFinal) + 1, 0) = SOCKET_ERROR Then
                    DoEvents
                    Call RaiseError(inSox, Err.LastDllError, Procedure, "On Send")
                    Let SendData = SOCKET_ERROR
                Else
                    DoEvents
                    Call Sox_SendComplete(inSox)
                End If
                If apiWSAAsyncSelect(Sockets(inSox).Socket, Portal.hWnd, ByVal Sockets(inSox).uMsg, ByVal FD_CLOSE Or FD_READ Or FD_WRITE) = SOCKET_ERROR Then
                    Call RaiseError(inSox, Err.LastDllError, Procedure, "On AsyncSelect")
                End If
            Case soxDisconnected
                Call RaiseError(inSox, 0, Procedure, "Unconnected")
                Let SendData = SOCKET_ERROR
            Case soxClosing
                Call RaiseError(inSox, 0, Procedure, "Closing")
                Let SendData = SOCKET_ERROR
            Case Else
                Call RaiseError(inSox, 0, Procedure, "Invalid State: " & Sockets(inSox).State)
                Let SendData = SOCKET_ERROR
        End Select
    End If
End Function

Private Function WinSockEvent(ByVal lParam As Long) As Integer 'WSAGETSELECTEVENT
    If (lParam And &HFFFF&) > &H7FFF Then
        Let WinSockEvent = (lParam And &HFFFF&) - &H10000
    Else
        Let WinSockEvent = lParam And &HFFFF&
    End If
End Function
Private Function WinSockError(ByVal lParam As Long) As Integer 'WSAGETSELECTERROR
    Let WinSockError = (lParam And &HFFFF0000) \ &H10000
End Function

Private Sub RaiseError(inSox As Long, inCode As Long, inProcedure As String, inSnipet As String)   'Returns EXACTLY the same value as inError but raises the corresponding event if this is an error
    Select Case inCode
        Case WSABASEERR: Call Sox_Error(inSox, inCode, "General WinSock subsystem failure", inProcedure, inSnipet) 'Just sounds cool :)))
        Case WSAEINTR: Call Sox_Error(inSox, inCode, "Interrupted function call", inProcedure, inSnipet)
        Case WSAEBADF: Call Sox_Error(inSox, inCode, "Unknown", inProcedure, inSnipet) ' Unknown
        Case WSAEACCES: Call Sox_Error(inSox, inCode, "Permission Denied", inProcedure, inSnipet)
        Case WSAEFAULT: Call Sox_Error(inSox, inCode, "Bad address", inProcedure, inSnipet)
        Case WSAEINVAL: Call Sox_Error(inSox, inCode, "Invalid argument", inProcedure, inSnipet)
        Case WSAEMFILE: Call Sox_Error(inSox, inCode, "Too many open files", inProcedure, inSnipet)
        Case WSAEWOULDBLOCK: Call Sox_Error(inSox, inCode, "Resource temporarily unavailable", inProcedure, inSnipet)
        Case WSAEINPROGRESS: Call Sox_Error(inSox, inCode, "Operation now in progress", inProcedure, inSnipet)
        Case WSAEALREADY: Call Sox_Error(inSox, inCode, "Operation already in progress", inProcedure, inSnipet)
        Case WSAENOTSOCK: Call Sox_Error(inSox, inCode, "Socket operation on non-socket", inProcedure, inSnipet)
        Case WSAEDESTADDRREQ: Call Sox_Error(inSox, inCode, "Destination address required", inProcedure, inSnipet)
        Case WSAEMSGSIZE: Call Sox_Error(inSox, inCode, "Message too long", inProcedure, inSnipet)
        Case WSAEPROTOTYPE: Call Sox_Error(inSox, inCode, "Protocol wrong type for socket", inProcedure, inSnipet)
        Case WSAENOPROTOOPT: Call Sox_Error(inSox, inCode, "Bad protocol option", inProcedure, inSnipet)
        Case WSAEPROTONOSUPPORT: Call Sox_Error(inSox, inCode, "Protocol not supported", inProcedure, inSnipet)
        Case WSAESOCKTNOSUPPORT: Call Sox_Error(inSox, inCode, "Socket type not supported", inProcedure, inSnipet)
        Case WSAEOPNOTSUPP: Call Sox_Error(inSox, inCode, "Operation not supported", inProcedure, inSnipet)
        Case WSAEPFNOSUPPORT: Call Sox_Error(inSox, inCode, "Protocol family not supported", inProcedure, inSnipet)
        Case WSAEAFNOSUPPORT: Call Sox_Error(inSox, inCode, "Address family not supported by protocol family", inProcedure, inSnipet)
        Case WSAEADDRINUSE: Call Sox_Error(inSox, inCode, "Address already in use", inProcedure, inSnipet)
        Case WSAEADDRNOTAVAIL: Call Sox_Error(inSox, inCode, "Cannot assign requested address", inProcedure, inSnipet)
        Case WSAENETDOWN: Call Sox_Error(inSox, inCode, "Network is down", inProcedure, inSnipet)
        Case WSAENETUNREACH: Call Sox_Error(inSox, inCode, "Network is unreachable", inProcedure, inSnipet)
        Case WSAENETRESET: Call Sox_Error(inSox, inCode, "Network dropped connection on reset", inProcedure, inSnipet)
        Case WSAECONNABORTED: Call Sox_Error(inSox, inCode, "Software caused connection abort", inProcedure, inSnipet)
        Case WSAECONNRESET: Call Sox_Error(inSox, inCode, "Connection reset by peer", inProcedure, inSnipet)
        Case WSAENOBUFS: Call Sox_Error(inSox, inCode, "No buffer space available", inProcedure, inSnipet)
        Case WSAEISCONN: Call Sox_Error(inSox, inCode, "Socket is already connected", inProcedure, inSnipet)
        Case WSAENOTCONN: Call Sox_Error(inSox, inCode, "Socket is not connected", inProcedure, inSnipet)
        Case WSAESHUTDOWN: Call Sox_Error(inSox, inCode, "Cannot send after socket shutdown", inProcedure, inSnipet)
        Case WSAETOOMANYREFS: Call Sox_Error(inSox, inCode, "Too many references: can't splice", inProcedure, inSnipet)  ' UnConfirmed Description
        Case WSAETIMEDOUT: Call Sox_Error(inSox, inCode, "Connection timed out", inProcedure, inSnipet)
        Case WSAECONNREFUSED: Call Sox_Error(inSox, inCode, "Connection refused", inProcedure, inSnipet)
        Case WSAELOOP: Call Sox_Error(inSox, inCode, "Too many levels of symbolic links", inProcedure, inSnipet)  ' UnConfirmed Description
        Case WSAENAMETOOLONG: Call Sox_Error(inSox, inCode, "File name too long", inProcedure, inSnipet)  ' UnConfirmed Description
        Case WSAEHOSTDOWN: Call Sox_Error(inSox, inCode, "Host is down", inProcedure, inSnipet)
        Case WSAEHOSTUNREACH: Call Sox_Error(inSox, inCode, "No route to host", inProcedure, inSnipet)
        Case WSAENOTEMPTY: Call Sox_Error(inSox, inCode, "Directory not empty", inProcedure, inSnipet)  ' UnConfirmed Description
        Case WSAEPROCLIM: Call Sox_Error(inSox, inCode, "Too many processes", inProcedure, inSnipet)
        Case WSAEUSERS: Call Sox_Error(inSox, inCode, "Too many users", inProcedure, inSnipet)  ' UnConfirmed Description
        Case WSAEDQUOT: Call Sox_Error(inSox, inCode, "Disk quota exceeded", inProcedure, inSnipet)  ' UnConfirmed Description
        Case WSAESTALE: Call Sox_Error(inSox, inCode, "Stale NFS file handle", inProcedure, inSnipet) ' UnConfirmed Description
        Case WSAEREMOTE: Call Sox_Error(inSox, inCode, "Too many levels of remote in path", inProcedure, inSnipet)  ' UnConfirmed Description
        Case WSASYSNOTREADY: Call Sox_Error(inSox, inCode, "Network subsystem is unavailable", inProcedure, inSnipet)
        Case WSAVERNOTSUPPORTED: Call Sox_Error(inSox, inCode, "WinSock.DLL version out of range", inProcedure, inSnipet)
        Case WSANOTINITIALISED: Call Sox_Error(inSox, inCode, "Successful WSAStartup not yet performed", inProcedure, inSnipet)
        Case WSAEDISCON1: Call Sox_Error(inSox, inCode, "Graceful shutdown in progress", inProcedure, inSnipet)
        Case WSAEDISCON2: Call Sox_Error(inSox, inCode, "Graceful shutdown in progress", inProcedure, inSnipet)
        Case WSAENOMORE: Call Sox_Error(inSox, inCode, "Unknown", inProcedure, inSnipet) ' Unknown
        Case WSAECANCELLED: Call Sox_Error(inSox, inCode, "Unknown", inProcedure, inSnipet) ' Unknown
        Case WSAEINVALIDPROCTABLE: Call Sox_Error(inSox, inCode, "Unknown", inProcedure, inSnipet) ' Unknown
        Case WSAEINVALIDPROVIDER: Call Sox_Error(inSox, inCode, "Unknown", inProcedure, inSnipet) ' Unknown
        Case WSAEPROVIDERFAILEDINIT: Call Sox_Error(inSox, inCode, "Unknown", inProcedure, inSnipet) ' Unknown
        Case WSASYSCALLFAILURE: Call Sox_Error(inSox, inCode, "System call failure", inProcedure, inSnipet)
        Case WSASERVICE_NOT_FOUND: Call Sox_Error(inSox, inCode, "Unknown", inProcedure, inSnipet) ' Unknown
        Case WSATYPE_NOT_FOUND: Call Sox_Error(inSox, inCode, "Class type not found", inProcedure, inSnipet)
        Case WSA_E_NO_MORE: Call Sox_Error(inSox, inCode, "Unknown", inProcedure, inSnipet) ' Unknown
        Case WSA_E_CANCELLED: Call Sox_Error(inSox, inCode, "Unknown", inProcedure, inSnipet) ' Unknown
        Case WSAEREFUSED: Call Sox_Error(inSox, inCode, "Unknown", inProcedure, inSnipet) ' Unknown
        Case WSAHOST_NOT_FOUND: Call Sox_Error(inSox, inCode, "Host not found", inProcedure, inSnipet)
        Case WSATRY_AGAIN: Call Sox_Error(inSox, inCode, "Non-authoritative host not found", inProcedure, inSnipet)
        Case WSANO_RECOVERY: Call Sox_Error(inSox, inCode, "This is a non-recoverable error", inProcedure, inSnipet)
        Case WSANO_DATA: Call Sox_Error(inSox, inCode, "Valid name, no data record of requested type", inProcedure, inSnipet)
        Case Else: Call Sox_Error(inSox, inCode, "Unrecognized WinSock error", inProcedure, inSnipet)
    End Select
End Sub

Private Sub RaiseStatus(inSox As Long, inProcedure As String, inStatus As String)
    Call Sox_Status(inSox, inProcedure, inStatus)
End Sub

Private Sub RaiseState(inSox As Long, inState As enmSoxState)
    Let Sockets(inSox).State = inState
    Call Sox_State(inSox, inState)
End Sub

Public Sub Hook()
    If Portal.WndProc = 0 Then Let Portal.WndProc = apiSetWindowLong(Portal.hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub
Public Sub Unhook()
    Call apiSetWindowLong(Portal.hWnd, GWL_WNDPROC, Portal.WndProc)
    Let Portal.WndProc = 0
End Sub

Private Function StringFromPointer(ByVal lPointer As Long) As String
    Let StringFromPointer = Space$(apiLStrLen(ByVal lPointer))
    Call apiLstrCpy(ByVal StringFromPointer, ByVal lPointer)
End Function


Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Const Procedure As String = "WndProc"
    Select Case uMsg
        Case soxSERVER
            Select Case WinSockEvent(lParam)
                Case FD_ACCEPT: Call RaiseStatus(IsInSox(wParam), Procedure, "Server received: FD_ACCEPT")
                    Select Case WinSockError(lParam)
                        Case 0: Accept (wParam)
                        Case Else: Call RaiseError(IsInSox(wParam), WinSockError(lParam), Procedure, "On FD_ACCEPT -- lParam: " & lParam)
                    End Select
                Case FD_CLOSE: Call RaiseStatus(IsInSox(wParam), Procedure, "Server received: FD_CLOSE")
                    Select Case WinSockError(lParam)
                        Case 0
                            Select Case Sockets(IsInSox(wParam)).State
                                Case soxClosing: Call ShutDown(IsInSox(wParam))
                                Case Else
                                    Call CloseIt(IsInSox(wParam))
                                    Call ShutDown(IsInSox(wParam))
                            End Select
                        Case Else
                            Call ShutDown(IsInSox(wParam))
                            Call RaiseError(IsInSox(wParam), WinSockError(lParam), Procedure, "On FD_CLOSE -- lParam: " & lParam)
                    End Select
                Case FD_CONNECT: Call RaiseStatus(IsInSox(wParam), Procedure, "Server received: FD_CONNECT")
                Case FD_READ: Call RaiseStatus(IsInSox(wParam), Procedure, "Server received: FD_READ")
                    Call GetData(IsInSox(wParam))
                Case FD_WRITE: Call RaiseStatus(IsInSox(wParam), Procedure, "Server received: FD_WRITE") ' A Server Client is ready to Send
                    Select Case WinSockError(lParam)
                        Case 0
                            Select Case Sockets(IsInSox(wParam)).State
                                Case soxConnecting: Call RaiseState(IsInSox(wParam), soxConnected)
                                Case soxDataSend: Let Sockets(IsInSox(wParam)).State = soxConnected
                            End Select
                        Case Else: Call RaiseError(IsInSox(wParam), WinSockError(lParam), Procedure, "On FD_WRITE -- lParam: " & lParam)
                    End Select
                Case Else
            End Select
        Case soxCLIENT
            Select Case WinSockEvent(lParam)
                Case FD_ACCEPT: Call RaiseStatus(IsInSox(wParam), Procedure, "Client received: FD_ACCEPT")
                Case FD_CLOSE: Call RaiseStatus(IsInSox(wParam), Procedure, "Client received: FD_CLOSE")
                    Select Case WinSockError(lParam)
                        Case 0
                            Select Case Sockets(IsInSox(wParam)).State
                                Case soxClosing: Call ShutDown(IsInSox(wParam))
                                Case Else
                                    Call CloseIt(IsInSox(wParam))
                                    Call ShutDown(IsInSox(wParam))
                            End Select
                        Case Else
                            Call ShutDown(IsInSox(wParam))
                            Call RaiseError(IsInSox(wParam), WinSockError(lParam), Procedure, "On FD_CLOSE -- lParam: " & lParam)
                    End Select
                Case FD_CONNECT: Call RaiseStatus(IsInSox(wParam), Procedure, "Client received: FD_CONNECT")
                Case FD_READ: Call RaiseStatus(IsInSox(wParam), Procedure, "Client received: FD_READ")
                Case FD_WRITE: Call RaiseStatus(IsInSox(wParam), Procedure, "Client received: FD_WRITE")
                    Select Case WinSockError(lParam)
                        Case 0
                            Select Case Sockets(IsInSox(wParam)).State
                                Case soxConnecting: Call RaiseState(IsInSox(wParam), soxConnected)
                                Case soxDataSend: Let Sockets(IsInSox(wParam)).State = soxConnected
                            End Select
                        Case Else: Call RaiseError(IsInSox(wParam), WinSockError(lParam), Procedure, "On FD_WRITE -- lParam: " & lParam)
                    End Select
                Case Else
            End Select
        Case Else: Let WndProc = apiCallWindowProc(Portal.WndProc, hWnd, uMsg, wParam, lParam)
    End Select
End Function

Private Function IsInSox(inSocket As Long) As Long ' Returns the Sockets() address of a WinSock Socket
    For IsInSox = 0 To Portal.Sockets
        If Sockets(IsInSox).Socket = inSocket Then Exit For
    Next IsInSox
    If IsInSox = Portal.Sockets + 1 Then Let IsInSox = INVALID_SOCKET
End Function

Public Function SetOption(inSox As Long, inOption As enmSoxOptions, inValue As Long) As Long
    Const Procedure As String = "SetOption"
    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Call RaiseError(INVALID_SOCKET, 0, Procedure, "Invalid Sox: " & inSox)
        Let SetOption = INVALID_SOCKET
    Else
        Select Case inOption
            Case soxSO_TCP_NODELAY
                If apiSetSockOpt(Sockets(inSox).Socket, IPPROTO_TCP, Not inOption, inValue, 4) = SOCKET_ERROR Then
                    Call RaiseError(inSox, Err.LastDllError, Procedure, "Option: " & inOption & " & Value: " & inValue)
                    Let SetOption = SOCKET_ERROR
                End If
            Case Else
                If apiSetSockOpt(Sockets(inSox).Socket, SOL_SOCKET, inOption, inValue, 4) = SOCKET_ERROR Then
                    Call RaiseError(inSox, Err.LastDllError, Procedure, "Option: " & inOption & " & Value: " & inValue)
                    Let SetOption = SOCKET_ERROR
                End If
        End Select
    End If
End Function

Public Function GetOption(inSox As Long, inOption As enmSoxOptions) As Long
    Const Procedure As String = "GetOption"
    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Call RaiseError(INVALID_SOCKET, 0, Procedure, "Invalid Sox: " & inSox)
        Let GetOption = INVALID_SOCKET
    Else
        Select Case inOption
            Case soxSO_TCP_NODELAY
                If apiGetSockOpt(Sockets(inSox).Socket, IPPROTO_TCP, Not inOption, GetOption, 4) = SOCKET_ERROR Then
                    Call RaiseError(inSox, Err.LastDllError, Procedure, "Option: " & inOption)
                    Let GetOption = SOCKET_ERROR
                End If
            Case Else
                If apiGetSockOpt(Sockets(inSox).Socket, SOL_SOCKET, inOption, GetOption, 4) = SOCKET_ERROR Then
                    Call RaiseError(inSox, Err.LastDllError, Procedure, "Option: " & inOption)
                    Let GetOption = SOCKET_ERROR
                End If
        End Select
    End If
End Function

Public Function SocketHandle(inSox As Long) As Long
    Let SocketHandle = Sockets(inSox).Socket
End Function

Public Function State(inSox As Long) As enmSoxState
    Let State = Sockets(inSox).State
End Function


Public Function LocalHostName() As String ' The PC's Name eg. RonaldR
    Const Procedure As String = "LocalHostName"
    Let LocalHostName = Space(256)
    If apiGetHostName(LocalHostName, 256) = SOCKET_ERROR Then
        Call RaiseError(SOCKET_ERROR, Err.LastDllError, Procedure, vbNullString)
        Let LocalHostName = vbNullString
    Else
        Let LocalHostName = Trim$(LocalHostName)
    End If
End Function
Public Function Address(inSox As Long) As String ' Returns the address used by a Socket (Either Local or Remote)
    Let Address = StringFromPointer(apiNLToIP(Sockets(inSox).SocketAddr.sin_addr))
End Function
Public Function Port(inSox As Long) As Long
    Let Port = apiNToHS(Sockets(inSox).SocketAddr.sin_port)
End Function

Private Function Byte2Long(inByte() As Byte) As Long ' Given 4 bytes, will directly copy them to a long! WARNING - To speed it up, I have no UBound checks, therefore you MUST send it 4 bytes
    Call apiCopyMemory(Byte2Long, inByte(0), 4)
End Function
Private Function Long2Byte(inLong As Long) As Byte() ' Given a Long, will directly copy it to an array of 4 bytes (0 to 3)!
    Dim tmpLong2Byte(0 To 3) As Byte
    Call apiCopyMemory(tmpLong2Byte(0), inLong, 4)
    Let Long2Byte = tmpLong2Byte
End Function
Private Sub Long2Byte2(inLong As Long, inByte() As Byte) ' Similar to the above, but places the bytes direcly into the given array
    Call apiCopyMemory(inByte(0), inLong, 4)
End Sub

Private Function DataSize(inSox As Long) As Long  ' Given 4 bytes, will directly copy them to a long! WARNING - To speed it up, I have no UBound checks, therefore you MUST send it 4 bytes
    Call apiCopyMemory(DataSize, Sockets(inSox).Data(0), 4)
End Function

Private Function ExtractData(inSox As Long) As Byte()
    Dim tmpData() As Byte
    ReDim tmpData(DataSize(inSox))
    Call apiCopyMemory(tmpData(0), Sockets(inSox).Data(4), DataSize(inSox) + 1)
    If DataSize(inSox) + 4 > Sockets(inSox).DataLen Then
        Let Sockets(inSox).DataLen = UBound(Sockets(inSox).Data) - (DataSize(inSox) + 4)
        Call apiCopyMemory(Sockets(inSox).Data(0), Sockets(inSox).Data(DataSize(inSox) + 5), UBound(Sockets(inSox).Data) - (DataSize(inSox) + 4))
        ReDim Preserve Sockets(inSox).Data(Sockets(inSox).DataLen)
    Else
        Let Sockets(inSox).DataLen = -1
        Erase Sockets(inSox).Data
    End If
'    MsgBox "Len: " & Sockets(inSox).DataLen
    Let ExtractData = tmpData
End Function

Private Sub ShutDown(inSox As Long)
    Const Procedure As String = "Shutdown"
    If inSox < 0 Or inSox > Portal.Sockets Then ' Detect out of Range of our Array ...
        Call RaiseError(INVALID_SOCKET, 0, Procedure, "Invalid Sox: " & inSox)
    Else
        If apiWSAAsyncSelect(Sockets(inSox).Socket, Portal.hWnd, ByVal FD_CLOSE, 0&) = SOCKET_ERROR Then       'FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
            Call RaiseError(inSox, Err.LastDllError, Procedure, "On AsyncSelect")
        Else
            If apiCloseSocket(Sockets(inSox).Socket) = SOCKET_ERROR Then  'I can't get the API that checks the current status of the socket to work :(((
                Call RaiseError(inSox, Err.LastDllError, Procedure, "On Close")
            Else
                Call RaiseState(inSox, soxDisconnected)
                Call RaiseStatus(inSox, Procedure, "Close Successful")
                Call Sox_Close(inSox)
                ' The following code is optional, the entire If statement could be remove successfully
                ' The use of this is to remove the Array member from the Sockets array if it is the last member, thus saving RAM, and iteration time for the IsInSox function
                ' However, that function executes so fast, that it'll propably save us 0.0000000001 seconds :)))
                ' By taking these statements out, it could improve performance if you expect lots of people connecting and disconnecting,
                ' by possibly having an unused Array member ready for use if there was a previous disconnected person
                ' My Suggestion is to comment this out due to the small memory size and possible increase in overall performance, by never making the Sockets Array smaller than its largest size
                
                If inSox = Portal.Sockets Then ' If this is the Last in the Sockets Array, we Remove it, also need to add code to remove all the ending Disconnected Sockets
                    If inSox = 0 Then
                        Erase Sockets
                        Let Portal.Sockets = -1
                    Else
                        ReDim Preserve Sockets(inSox - 1) As typSocket
                        Let Portal.Sockets = inSox - 1
                    End If
                End If
            End If
        End If
    End If
End Sub

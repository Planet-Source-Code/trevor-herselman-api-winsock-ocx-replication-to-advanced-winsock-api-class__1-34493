Attribute VB_Name = "modSox"
Option Explicit

Public Sox As clsSox

Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Let WindowProc = Sox.WndProc(hWnd, uMsg, wParam, lParam)
End Function

Public Sub Sox_Close(inSox As Long) 'Notification of a close
    Let frmSox.txtConnections = frmSox.txtConnections - 1
End Sub

Public Sub Sox_Connect(inSox As Long) 'Notification of connection
    Let frmSox.txtConnections = frmSox.txtConnections + 1
End Sub

Public Sub Sox_DataArrival(inSox As Long, inData() As Byte)
    Let frmSox.txtServerReceived = frmSox.txtServerReceived & "Received: " & UBound(inData) + 1 & " bytes" & vbCrLf
    Let frmSox.txtServerReceived = frmSox.txtServerReceived & Left$(StrConv(inData, vbUnicode), 20) & vbCrLf
End Sub

Public Sub Sox_Connection(inSox As Long) 'Notification of a new connection (From a Listening Port)
    Let frmSox.txtConnections = frmSox.txtConnections + 1
End Sub

' This is the Old WinSock Error Event ... Too complicated and unnecessary ... who used it anyway ???
' Public Event Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Public Sub Sox_Error(inSox As Long, inerror As Long, inDescription As String, inSource As String, inSnipet As String)
    With frmSox.txtStatus
        Let .Text = .Text & "Error: (Socket) " & inSox & " (Error) " & inerror & " (Description) " & inDescription & " (Source) " & inSource & " (Area) " & inSnipet & vbCrLf
        Let .SelStart = Len(.Text) 'Just makes our new message visible
    End With
End Sub

Public Sub Sox_SendComplete(inSox As Long)
    With frmSox.txtStatus
        Let .Text = .Text & "Sox: " & inSox & " Data Sent to WinSock buffers successfully" & vbCrLf
        Let .SelStart = Len(.Text) 'Just makes our new message visible
    End With
End Sub

' Currently unused
Public Sub Sox_SendProgress(inSox As Long, bytesSent As Long, bytesRemaining As Long)

End Sub

Public Sub Sox_Status(inSox As Long, inSource As String, inStatus As String)
    With frmSox.txtStatus
        Let .Text = .Text & "Status: (Socket) " & inSox & " (Source) " & inSource & " (Status) " & inStatus & vbCrLf
        Let .SelStart = Len(.Text) 'Just makes our new message visible
    End With
End Sub

Public Sub Sox_State(inSox As Long, inState As enmSoxState) 'Returns state changes of various sockets
    With frmSox.txtStatus
        Select Case inState
            Case soxDisconnected
                Let .Text = .Text & "State: (Socket) " & inSox & " (State) Disconnected" & vbCrLf
                Let .SelStart = Len(.Text) 'Just makes our new message visible
            Case soxConnecting
                Let .Text = .Text & "State: (Socket) " & inSox & " (State) Connecting" & vbCrLf
                Let .SelStart = Len(.Text) 'Just makes our new message visible
            Case soxConnected
                Let .Text = .Text & "State: (Socket) " & inSox & " (State) Connected" & vbCrLf
                Let .SelStart = Len(.Text) 'Just makes our new message visible
'            Case soxWaitingAnswer ' Not Used - For future use - probably app dependant
'                Let .Text = .Text & "State: (Socket) " & inSox & " (State) WaitingAnswer" & vbCrLf
'                Let .SelStart = Len(.Text) 'Just makes our new message visible
'            Case soxCommandAnswered ' Not Used - For future use - probably app dependant
'                Let .Text = .Text & "State: (Socket) " & inSox & " (State) CommandAnswered" & vbCrLf
'                Let .SelStart = Len(.Text) 'Just makes our new message visible
'            Case soxCommandNotAnswered ' Not Used - For future use - probably app dependant
'                Let .Text = .Text & "State: (Socket) " & inSox & " (State) CommandNotAnswered" & vbCrLf
'                Let .SelStart = Len(.Text) 'Just makes our new message visible
'            Case soxDataReceived
'                Let .Text = .Text & "State: (Socket) " & inSox & " (State) DataReceived" & vbCrLf
'                Let .SelStart = Len(.Text) 'Just makes our new message visible
            Case soxListening
                Let .Text = .Text & "State: (Socket) " & inSox & " (State) Listening" & vbCrLf
                Let .SelStart = Len(.Text) 'Just makes our new message visible
'            Case soxWaitingBinary ' Not Used
'                Let .Text = .Text & "State: (Socket) " & inSox & " (State) WaitingBinary" & vbCrLf
'                Let .SelStart = Len(.Text) 'Just makes our new message visible
'            Case soxERROR
'                Let .Text = .Text & "State: (Socket) " & inSox & " (State) Error" & vbCrLf
'                Let .SelStart = Len(.Text) 'Just makes our new message visible
            Case Else
                Let .Text = .Text & "State: (Socket) " & inSox & " (State) Unknown State: " & inState & vbCrLf
                Let .SelStart = Len(.Text) 'Just makes our new message visible
        End Select
    End With
End Sub

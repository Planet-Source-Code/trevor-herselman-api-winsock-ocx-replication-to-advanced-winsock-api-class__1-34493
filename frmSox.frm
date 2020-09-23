VERSION 5.00
Begin VB.Form frmSox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sox API"
   ClientHeight    =   7275
   ClientLeft      =   1050
   ClientTop       =   1410
   ClientWidth     =   9225
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   9225
   Begin VB.TextBox txtStatus 
      Height          =   1635
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   42
      Top             =   5580
      Width           =   8535
   End
   Begin VB.CommandButton btnTest 
      Caption         =   "Test"
      Height          =   495
      Left            =   7680
      TabIndex        =   39
      Top             =   360
      Width           =   1215
   End
   Begin VB.Frame frmSubClassing 
      Caption         =   "SubClassing"
      Height          =   915
      Left            =   4680
      TabIndex        =   14
      ToolTipText     =   "This must be done before we can receive messages and MUST be undone before form unload"
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton btnHook 
         Caption         =   "Hook"
         Height          =   495
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Hook into messaging sub-system (Must be done before ANY messages will be received from WinSock)"
         Top             =   300
         Width           =   1215
      End
      Begin VB.CommandButton btnUnhook 
         Caption         =   "Unhook"
         Height          =   495
         Left            =   1440
         TabIndex        =   15
         ToolTipText     =   "UnHook messaging sub-system (MUST be done before form unloads ... done automatically there currently)"
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame frmClient 
      Caption         =   "Client"
      Height          =   4275
      Left            =   4620
      TabIndex        =   1
      Top             =   1260
      Width           =   4515
      Begin VB.TextBox txtClientReceived 
         Height          =   1155
         Left            =   960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   3000
         Width           =   3375
      End
      Begin VB.Frame frmConnection 
         Caption         =   "Connection"
         Height          =   1455
         Left            =   180
         TabIndex        =   29
         Top             =   300
         Width           =   4155
         Begin VB.CommandButton btnCloseConnect 
            Caption         =   "Close"
            Height          =   315
            Left            =   360
            TabIndex        =   40
            Top             =   1020
            Width           =   1215
         End
         Begin VB.TextBox txtServerSocket 
            Height          =   315
            Left            =   2700
            Locked          =   -1  'True
            TabIndex        =   37
            ToolTipText     =   "This is the Socket on the Server we have access to"
            Top             =   1020
            Width           =   1215
         End
         Begin VB.TextBox txtServerPort 
            Height          =   315
            Left            =   2700
            TabIndex        =   32
            Text            =   "1234"
            ToolTipText     =   "The network port you want to connect to"
            Top             =   180
            Width           =   1215
         End
         Begin VB.TextBox txtServerAddress 
            Height          =   315
            Left            =   2700
            TabIndex        =   31
            Text            =   "127.0.0.1"
            ToolTipText     =   "The IP / Internet address you want to connect to"
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton btnConnect 
            Caption         =   "Connect"
            Height          =   495
            Left            =   360
            TabIndex        =   30
            ToolTipText     =   "Make a connection to the server at the Port and Address specified"
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblServerSocket 
            AutoSize        =   -1  'True
            Caption         =   "Socket:"
            Height          =   195
            Left            =   2040
            TabIndex        =   38
            ToolTipText     =   "The network card you want to dedicate to Listening"
            Top             =   1020
            Width           =   555
         End
         Begin VB.Label lblServertPort 
            AutoSize        =   -1  'True
            Caption         =   "Port:"
            Height          =   195
            Left            =   2280
            TabIndex        =   34
            Tag             =   "The network port you want to dedicate to Listening"
            Top             =   180
            Width           =   330
         End
         Begin VB.Label lblServerAddress 
            AutoSize        =   -1  'True
            Caption         =   "Address:"
            Height          =   195
            Left            =   2040
            TabIndex        =   33
            ToolTipText     =   "The network card you want to dedicate to Listening"
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.Frame frmServerSend 
         Caption         =   "Send - To Server"
         Height          =   1155
         Left            =   180
         TabIndex        =   19
         Top             =   1800
         Width           =   4155
         Begin VB.CommandButton btnServerSend 
            Caption         =   "Send"
            Height          =   495
            Left            =   180
            TabIndex        =   21
            ToolTipText     =   "Send the message to the Server"
            Top             =   420
            Width           =   1215
         End
         Begin VB.TextBox txtServerMessage 
            Height          =   315
            Left            =   1620
            TabIndex        =   20
            Text            =   "Hello"
            ToolTipText     =   "Type a message to send to the Server here"
            Top             =   660
            Width           =   2355
         End
         Begin VB.Label lblServerMessage 
            AutoSize        =   -1  'True
            Caption         =   "Message:"
            Height          =   195
            Left            =   2340
            TabIndex        =   22
            Top             =   360
            Width           =   690
         End
      End
      Begin VB.Label lblClientReceived 
         Caption         =   "Received:"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   3000
         Width           =   735
      End
   End
   Begin VB.Frame frmServer 
      Caption         =   "Server"
      Height          =   5475
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4515
      Begin VB.Frame frmListen 
         Caption         =   "Listen"
         Height          =   1155
         Left            =   180
         TabIndex        =   23
         ToolTipText     =   "Binding is usually used on Multiple network card configurations"
         Top             =   1500
         Width           =   4155
         Begin VB.CommandButton btnCloseListen 
            Caption         =   "Close"
            Height          =   315
            Left            =   2820
            TabIndex        =   41
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton btnListen 
            Caption         =   "Listen"
            Height          =   495
            Left            =   2820
            TabIndex        =   26
            ToolTipText     =   $"frmSox.frx":0000
            Top             =   180
            Width           =   1215
         End
         Begin VB.TextBox txtConnections 
            Height          =   315
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   25
            Text            =   "0"
            ToolTipText     =   "How many Sockets are being used"
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtListenSocket 
            Height          =   315
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   24
            ToolTipText     =   "The Socket that is dedicated to Listening"
            Top             =   720
            Width           =   975
         End
         Begin VB.Label lblConnections 
            AutoSize        =   -1  'True
            Caption         =   "Current Connections:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   1485
         End
         Begin VB.Label lblListenSocket 
            AutoSize        =   -1  'True
            Caption         =   "Listening On Socket:"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   1485
         End
      End
      Begin VB.TextBox txtServerReceived 
         Height          =   1155
         Left            =   960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   4200
         Width           =   3375
      End
      Begin VB.Frame frmBind 
         Caption         =   "Bind"
         Height          =   1155
         Left            =   180
         TabIndex        =   5
         ToolTipText     =   "Binding is usually used on Multiple network card configurations"
         Top             =   300
         Width           =   4155
         Begin VB.TextBox txtBindAddress 
            Height          =   315
            Left            =   2580
            TabIndex        =   9
            Text            =   "127.0.0.1"
            ToolTipText     =   "The network card you want to dedicate to Listening (Address 255.255.255.255 is invalid)"
            Top             =   660
            Width           =   1455
         End
         Begin VB.TextBox txtBindPort 
            Height          =   315
            Left            =   2520
            TabIndex        =   7
            Text            =   "1234"
            ToolTipText     =   "The network port you want to dedicate to Listening (Valid ports R in range 0 to 5000)"
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton btnBind 
            Caption         =   "Bind"
            Height          =   495
            Left            =   300
            TabIndex        =   6
            ToolTipText     =   "Bind this Port and Address for Listening"
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblBindAddress 
            AutoSize        =   -1  'True
            Caption         =   "Address:"
            Height          =   195
            Left            =   1800
            TabIndex        =   10
            ToolTipText     =   "The network card you want to dedicate to Listening"
            Top             =   660
            Width           =   615
         End
         Begin VB.Label lblBindPort 
            AutoSize        =   -1  'True
            Caption         =   "Port:"
            Height          =   195
            Left            =   2085
            TabIndex        =   8
            Tag             =   "The network port you want to dedicate to Listening"
            Top             =   300
            Width           =   330
         End
      End
      Begin VB.Frame frmClientSend 
         Caption         =   "Send - To Client(s)"
         Height          =   1455
         Left            =   180
         TabIndex        =   2
         Top             =   2700
         Width           =   4155
         Begin VB.TextBox txtClientMessage 
            Height          =   315
            Left            =   1020
            TabIndex        =   12
            Text            =   "Hello"
            Top             =   960
            Width           =   2955
         End
         Begin VB.TextBox txtClientSocket 
            Height          =   315
            Left            =   2940
            TabIndex        =   4
            Text            =   "0"
            ToolTipText     =   "Socket 0 will broadcast to all sockets"
            Top             =   420
            Width           =   1035
         End
         Begin VB.CommandButton btnClientSend 
            Caption         =   "Send"
            Height          =   495
            Left            =   420
            TabIndex        =   3
            ToolTipText     =   "Send the message to the connected client Socket"
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblClientMessage 
            AutoSize        =   -1  'True
            Caption         =   "Message:"
            Height          =   195
            Left            =   180
            TabIndex        =   13
            Top             =   960
            Width           =   690
         End
         Begin VB.Label lblClientSocket 
            AutoSize        =   -1  'True
            Caption         =   "To Socket:"
            Height          =   195
            Left            =   2040
            TabIndex        =   11
            Top             =   420
            Width           =   795
         End
      End
      Begin VB.Label lblServerReceived 
         Caption         =   "Received:"
         Height          =   255
         Left            =   180
         TabIndex        =   18
         Top             =   4200
         Width           =   735
      End
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Status:"
      Height          =   195
      Left            =   60
      TabIndex        =   43
      Top             =   5580
      Width           =   495
   End
End
Attribute VB_Name = "frmSox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'IMPORTANT - Do NOT hit the Stop button in the IDE while the form is hooked or VB WILL crash (Tell them not to do something and I guarantee they'll do it :)))

'NB - You must Hook the Form before invoking Sox.Listen(Me.hWnd) or you WON'T get any WinSock messages !!! The listen command will execute properly anyway !!! (This is good for debugging in case your app crashes)

'Public Sox As clsSox 'WISH I could use Events here :((( but SubClassing doesn't allow us to generate events, so I've removed them all :((( They'll be handled inside the Class

Private Sub btnTest_Click()
'    MsgBox Sox.GetOption(txtServerSocket, soxSO_RCVBUF) ' Current status, will be 0 = No keep alive
'    MsgBox Sox.SetOption(txtServerSocket, soxSO_RCVBUF, 1) ' Enable Keep Alive (ANY non zero value will enable it)
'    MsgBox Sox.GetOption(txtServerSocket, soxSO_RCVBUF) ' Check the new Keep alive status (Will return 1 = Enabled)
    Call Sox.CloseIt(1)
End Sub

'Form Specific

Private Sub Form_Activate()
    Set Sox = New clsSox
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Sox = Nothing
End Sub

'SubClassing

Private Sub btnHook_Click()
    Call Sox.Hook    'Dangerous stuff :))) Don't forget to UnHook (Done for you in Form_UnLoad)
End Sub

Private Sub btnUnhook_Click()
    Call Sox.Unhook
End Sub

'Server

Private Sub btnBind_Click()
    Call Sox.Bind(Val(txtBindPort), txtBindAddress)
End Sub

Private Sub btnListen_Click()
    Let txtListenSocket.Text = Sox.Listen(txtBindAddress, txtBindPort)
End Sub

Private Sub btnClientSend_Click()
    Call Sox.SendData(txtClientSocket, txtClientMessage)
End Sub

Private Sub btnCloseListen_Click()
    Call Sox.CloseIt(txtListenSocket)
End Sub

'Client

Private Sub btnConnect_Click()
    Let txtServerSocket = Sox.Connect(txtServerAddress, Val(txtServerPort))
    Let txtClientSocket = txtServerSocket
End Sub

Private Sub btnServerSend_Click()
    Dim Start As Single
    Dim TestArray(90) As Byte
    Let Start = Timer
'    If Sox.SendData(Val(txtServerSocket), txtServerMessage) = SOCKET_ERROR Then
'        Call MsgBox("Error On Sending")
'    End If

'    If Sox.SendData(Val(txtServerSocket), Space(5000000)) = SOCKET_ERROR Then
'        Call MsgBox("Error On Sending")
'    End If

'    If Sox.SendData(Val(txtServerSocket), TestArray) = SOCKET_ERROR Then
'        Call MsgBox("Error On Sending")
'    End If

    If Sox.SendData(Val(txtServerSocket), "Hello") = soxERROR Then
        Call MsgBox("Error On Sending")
    End If
    If Sox.SendData(Val(txtServerSocket), String(10000, 97)) = soxERROR Then
        Call MsgBox("Error On Sending")
    End If

'    If Sox.SendData(Val(txtServerSocket), "1") = SOCKET_ERROR Then
'        Call MsgBox("Error On Sending")
'    End If
'    If Sox.SendData(Val(txtServerSocket), "2") = SOCKET_ERROR Then
'        Call MsgBox("Error On Sending")
'    End If
'    If Sox.SendData(Val(txtServerSocket), "3") = SOCKET_ERROR Then
'        Call MsgBox("Error On Sending")
'    End If
'    If Sox.SendData(Val(txtServerSocket), "4") = SOCKET_ERROR Then
'        Call MsgBox("Error On Sending")
'    End If
    
'    MsgBox Timer - Start
End Sub

Private Sub btnCloseConnect_Click()
    Call Sox.CloseIt(txtServerSocket)
End Sub

'Misc

Private Sub txtBindAddress_Click()
    Let txtBindAddress = ""
End Sub

Private Sub txtServerAddress_Click()
    Let txtServerAddress = ""
End Sub

Private Sub txtServerReceived_DblClick()
    Let txtServerReceived.Text = ""
End Sub

Private Sub txtStatus_DblClick()
    Let txtStatus.Text = ""
End Sub

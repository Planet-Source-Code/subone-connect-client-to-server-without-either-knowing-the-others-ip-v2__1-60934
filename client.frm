VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   2805
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrBroadcast 
      Interval        =   1000
      Left            =   0
      Top             =   840
   End
   Begin MSWinsockLib.Winsock sckBroadcast 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckClient 
      Left            =   0
      Top             =   420
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblRemotePort 
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1260
      TabIndex        =   7
      Top             =   960
      Width           =   1515
   End
   Begin VB.Label lblRemoteHostIP 
      Caption         =   "[Connecting... 10]"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1260
      TabIndex        =   6
      Top             =   660
      Width           =   1515
   End
   Begin VB.Label lblLocalPort 
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1260
      TabIndex        =   5
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblLocalHostIP 
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1260
      TabIndex        =   4
      Top             =   60
      Width           =   1335
   End
   Begin VB.Label lblStatic 
      Alignment       =   1  'Right Justify
      Caption         =   "RemotePort:"
      Height          =   255
      Index           =   3
      Left            =   60
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblStatic 
      Alignment       =   1  'Right Justify
      Caption         =   "RemoteHostIP:"
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   2
      Top             =   660
      Width           =   1095
   End
   Begin VB.Label lblStatic 
      Alignment       =   1  'Right Justify
      Caption         =   "LocalPort:"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblStatic 
      Alignment       =   1  'Right Justify
      Caption         =   "LocalHostIP:"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1095
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Structures used for converting numbers to binary strings
Private Type LngType
    L As Long
End Type

Private Type StrType
    S As String * 8
End Type


'Port constants
Private Const CLIENT_BROADCAST_PORT = 6125
Private Const SERVER_BROADCAST_PORT = 6126
Private Const SERVER_PORT = 6127

'Protocol constants
Private Const DELIMITER = "åî"
Private Const CONNECT_TO_HOST = 1
Private Const IDENTIFY_SERVER = "BCs2"
Private Const IDENTIFY_CLIENT = "BCc2"


Private Sub Form_Load()

    If App.PrevInstance Then End
    
    'First we setup the client winsock
    With sckClient
        .Protocol = sckTCPProtocol
        .RemotePort = SERVER_PORT
    End With

    'Then the broadcast winsock
    With sckBroadcast
        .Protocol = sckUDPProtocol
        .LocalPort = CLIENT_BROADCAST_PORT
        .RemotePort = SERVER_BROADCAST_PORT
        .RemoteHost = "255.255.255.255"
        'This part is important, I'm not sure why, but if you
        'don't send a packet over the broadcast address from
        'one sock then it won't start receiving them from other
        'socks. I dunno if it's just VB, but if you know a
        'better way let me know. This is the part that some
        'other tutorials leave out BTW.
        .SendData ""
    End With
    
End Sub


Private Sub sckBroadcast_DataArrival(ByVal bytesTotal As Long)

    If bytesTotal = 0 Then Exit Sub
    
    Dim dat As String
    Dim param() As String
    
    'Extract the data to a local variable...
    sckBroadcast.GetData dat
    '...and seperate it into parameters
    param = Split(dat, DELIMITER)
    
    'Check that the message is from the server
    Select Case param(0)
        'That's him alright!
        Case IDENTIFY_SERVER

            'Check what the server wants
            Select Case CVL(param(1))
            
                'The server wants us to connect
                Case CONNECT_TO_HOST
                    If sckClient.State = sckClosed Then sckClient.Connect param(2)
                '}
                
            End Select
        
        '}
        
    End Select
    
End Sub


Private Sub sckClient_Connect()
    tmrBroadcast.Enabled = False
    
    sckClient.SendData IDENTIFY_CLIENT & DELIMITER & MKL(CONNECT_TO_HOST) & _
        DELIMITER & sckClient.RemoteHostIP
        
    'Show the connection info on our labels
    lblLocalHostIP.Caption = sckClient.LocalIP
    lblLocalPort.Caption = sckClient.LocalPort
    lblRemoteHostIP.Caption = sckClient.RemoteHostIP
    lblRemotePort.Caption = sckClient.RemotePort
End Sub


'Used to convert a 4 byte string into a 4 byte long integer
Public Function CVL(Value As String) As Long
    Dim nStr As StrType
    Dim nLng As LngType
    
    nStr.S = StrConv(Left$(Value, 4), vbFromUnicode)
    LSet nLng = nStr
    CVL = nLng.L
End Function


'Used to convert a 4 byte long integer into a 4 byte string
Public Function MKL(ByVal Value As Long) As String
    Dim nStr As StrType
    Dim nLng As LngType
    
    nLng.L = Value
    LSet nStr = nLng
    MKL = Left(StrConv(nStr.S, vbUnicode), 4)
End Function


Private Sub tmrBroadcast_Timer()
    Static nTries As Long
    
    lblRemoteHostIP.Caption = "[Connecting... " & 9 - nTries & "]"
    
    sckBroadcast.SendData IDENTIFY_CLIENT & DELIMITER & MKL(CONNECT_TO_HOST) & _
        DELIMITER & sckClient.LocalHostName
                          
    nTries = nTries + 1
    If nTries = 10 Then
        nTries = 0
        tmrBroadcast.Enabled = False
        lblRemoteHostIP.Caption = "[Unable to connect]"
    End If
End Sub

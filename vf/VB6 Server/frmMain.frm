VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   Caption         =   "Server"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrConnected 
      Interval        =   500
      Left            =   3840
      Top             =   0
   End
   Begin VB.TextBox txtMessage 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   4455
   End
   Begin VB.TextBox txtReceived 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   4455
   End
   Begin VB.TextBox txtErrors 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   3840
      Width           =   4455
   End
   Begin MSWinsockLib.Winsock wsClients 
      Index           =   0
      Left            =   1560
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
      LocalPort       =   80
   End
   Begin MSWinsockLib.Winsock wsServer 
      Left            =   2640
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
      LocalPort       =   80
   End
   Begin VB.Label Label2 
      Caption         =   "Send Message"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3015
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   4560
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   4560
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label4 
      Caption         =   "Received"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Error Log"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    wsServer.LocalPort = 80
    wsServer.RemotePort = 80
    wsServer.Listen
    ' We Listen to Port 9000, and we don't need a remotehost for it.
End Sub


Private Sub tmrConnected_Timer()
Dim i As Integer, intCon As Integer

    intCon = 0
    ' cycle thru all clients
    For i = 0 To wsClients.UBound
        If wsClients(i).State = 7 Then ' if conencted
            intCon = intCon + 1        ' count the client
        End If
    Next i
        
    Me.Caption = "Server - " & intCon & " clients"
    ' Just show, how many Cons are at this time
End Sub

Private Sub txtMessage_KeyPress(KeyCode As Integer)
Dim i As Integer
    
    If KeyCode = 13 And Trim(txtMessage.Text) <> "" Then ' enter and not Empty
        ServerSendData txtMessage.Text
        ' direct broadcast
        txtMessage.Text = ""
        ' no print on the Server-Chatbox - Server don't need to see
        ' what it sended ^^
        KeyCode = 0
        ' Don't pleep, please
    End If

End Sub

Private Sub wsClients_DataArrival(index As Integer, ByVal bytesTotal As Long)
' called, if Server gets data from the clients.
Dim Data As String

    wsClients(index).GetData Data
    ' get the Clientdata
    
    ServerSendData Data
    ' .. and echo it as Server
     
    txtReceived.SelStart = Len(txtReceived.Text)
    txtReceived.SelText = Data & vbCrLf
    ' Print it.

End Sub


Private Sub wsServer_ConnectionRequest(ByVal requestID As Long)
' if an Client requests a connection
Dim index As Integer
    
    index = GetOpenWinsock
    wsClients(index).Accept requestID
    ' attention:
    ' Request at wsServer, accept at wsClient
End Sub



Private Function GetOpenWinsock() As Integer
'Searches the first open Winsock for us.
Static intUsedPorts As Integer
Dim i As Integer, bOpenWinSockFound As Boolean

bOpenWinSockFound = False
    
    For i = wsClients.UBound To 0 Step -1
        If wsClients(i).State = 0 And Not bOpenWinSockFound Then
            bOpenWinSockFound = True
            GetOpenWinsock = i
        End If
    Next i

    If Not bOpenWinSockFound Then
    ' no open ClientSock found.
        Load wsClients(wsClients.UBound + 1)
        ' load a new Client-winsock into tha array
        intUsedPorts = intUsedPorts + 1 ' new port
        wsClients(wsClients.UBound).LocalPort = wsClients(wsClients.UBound).LocalPort + intUsedPorts

        GetOpenWinsock = wsClients.UBound
        'return the new winsock in the array.
    End If


End Function



Private Sub wsClients_Error(index As Integer, ByVal iNumber As Integer, strDescription As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'called, if an error occured
    txtErrors.SelStart = Len(txtErrors.Text)
    txtErrors.SelText = "wsClients(" & index & ") - " & iNumber & " - " & strDescription & vbCrLf
    wsClients(index).Close
    ' Close the Winsock on which the error happend.
End Sub
Private Sub wsServer_Error(ByVal iNumber As Integer, strDescription As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
' same as above for wsServer-Winsock
    txtErrors.SelStart = Len(txtErrors.Text)
    txtErrors.SelText = "wsServer - " & iNumber & " - " & strDescription & vbCrLf
End Sub



' would be fine in an Module too - but i did it here to keep the source short.
Public Function ServerSendData(pstrSendData As String)
' cycle thru all clients, sending the same data.
Dim i As Integer

    For i = 0 To wsClients.UBound
        If wsClients(i).State = 7 Then
            wsClients(i).SendData "server: " & pstrSendData
            DoEvents ' important: * Not to lock the PC
                     '            * There seems to been a bug with winsock.
        End If
    Next i
End Function


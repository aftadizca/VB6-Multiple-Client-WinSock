VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Client 
   BackColor       =   &H00404040&
   Caption         =   "CLIENT"
   ClientHeight    =   5940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7125
   FillColor       =   &H00C0C0FF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   7125
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox IpText 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   360
      Width           =   3255
   End
   Begin VB.TextBox PortText 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton Start 
      Appearance      =   0  'Flat
      Caption         =   "CONNECT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   3960
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton CloseBtn 
      Appearance      =   0  'Flat
      Caption         =   "CLOSE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   5520
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox ChatDisplay 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1560
      Width           =   6735
   End
   Begin VB.TextBox Message 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   5040
      Width           =   4815
   End
   Begin VB.CommandButton SendMsg 
      Caption         =   "SEND"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   0
      Top             =   5040
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6720
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Port"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "IP Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CloseBtn_Click()
    Winsock1.Close
    ChatDisplay.Text = ChatDisplay.Text & "----Connection Closed :" & Winsock1.RemoteHost & vbCrLf
    Start.Enabled = True
    CloseBtn.Enabled = False
End Sub

Private Sub Form_Load()
    IpText.Text = Winsock1.LocalIP
    PortText.Text = "11111"
End Sub

Private Sub SendMsg_Click()
    If Message.Text <> "" And Winsock1.State = StateConstants.sckConnected Then
        Winsock1.SendData Message.Text
        ChatDisplay.Text = ChatDisplay.Text & "ME :" & Message.Text & vbCrLf
    End If
End Sub

Private Sub Start_Click()
    Winsock1.Connect Winsock1.LocalIP, "11111"
    Start.Enabled = False
    CloseBtn.Enabled = True
End Sub

Private Sub Winsock1_Close()
    ChatDisplay.Text = ChatDisplay.Text & "----Connection Closed :" & Winsock1.RemoteHost & vbCrLf
    Start.Enabled = True
    CloseBtn.Enabled = False
    Winsock1.Close
End Sub

Private Sub Winsock1_Connect()
    ChatDisplay.Text = ChatDisplay.Text & "----Connection Started :" & Winsock1.RemoteHost & vbCrLf
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    
    Winsock1.Accept requestID
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim MsgData As String
    Winsock1.GetData MsgData, vbString
    ChatDisplay.Text = ChatDisplay.Text & Winsock1.RemoteHost & " : " & MsgData & vbCrLf
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ChatDisplay.Text = ChatDisplay.Text & "----Connection Failed :" & Winsock1.RemoteHost & vbCrLf
    Winsock1.Close
    Start.Enabled = True
    CloseBtn.Enabled = False
End Sub

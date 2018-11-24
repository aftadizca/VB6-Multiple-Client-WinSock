VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Server 
   BackColor       =   &H00404040&
   Caption         =   "SERVER"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7035
   DrawStyle       =   5  'Transparent
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
   ScaleHeight     =   5985
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton SendMsg 
      Caption         =   "SEND"
      Height          =   735
      Left            =   5040
      TabIndex        =   8
      Top             =   5040
      Width           =   1815
   End
   Begin VB.TextBox Message 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   5040
      Width           =   4815
   End
   Begin VB.TextBox ChatDisplay 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      Height          =   3375
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1560
      Width           =   6735
   End
   Begin VB.CommandButton CloseBtn 
      Appearance      =   0  'Flat
      Caption         =   "CLOSE"
      Enabled         =   0   'False
      Height          =   840
      Left            =   5400
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Start 
      Appearance      =   0  'Flat
      Caption         =   "LISTEN"
      Height          =   840
      Left            =   3840
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   6600
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   3255
   End
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3255
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3255
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
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3495
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public NumSockets As Integer
Private Sub CloseBtn_Click()
    Dim Index As Integer
        For Index = 0 To NumSockets
            Winsock1(Index).Close
        Next Index
    ChatDisplay.Text = ChatDisplay.Text & "----Server shutdown" & vbCrLf
    Start.Enabled = True
    CloseBtn.Enabled = False
End Sub

Private Sub Form_Load()
    IpText.Text = Winsock1(0).LocalIP
    PortText.Text = "11111"
    Dim Client1 As New Client
    Dim Client2 As New Client
    Dim Client3 As New Client
    
    Client1.Show
    Client2.Show
    Client3.Show
End Sub

Private Sub SendMsg_Click()
    Dim Index As Integer
    If Message.Text <> "" Then
        For Index = 1 To NumSockets
            Winsock1(Index).SendData Message.Text
        Next Index
        ChatDisplay.Text = ChatDisplay.Text & "ME :" & Message.Text & vbCrLf
    End If
End Sub

Private Sub Start_Click()
    Winsock1(0).LocalPort = PortText.Text
    ChatDisplay.Text = ChatDisplay.Text & "----Starting server in port " & Winsock1(0).LocalPort & vbCrLf
    Winsock1(0).Listen
    Start.Enabled = False
    CloseBtn.Enabled = True
End Sub

Private Sub Winsock1_Close(Index As Integer)
    ChatDisplay.Text = ChatDisplay.Text & "----Connection Closed :" & Winsock1(Index).RemoteHostIP & vbCrLf
    Winsock1(Index).Close
    'Unload Winsock1(Index)
End Sub

Private Sub Winsock1_Connect(Index As Integer)
    ChatDisplay.Text = ChatDisplay.Text & "----Connected : " & Winsock1(Index).RemoteHostIP & vbCrLf
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    NumSockets = NumSockets + 1
    Load Winsock1(NumSockets)
    Winsock1(NumSockets).Accept requestID
    ChatDisplay.Text = ChatDisplay.Text & "----Connected : " & Winsock1(NumSockets).RemoteHostIP & vbCrLf
    Print NumSockets
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim MsgData As String
    Winsock1(Index).GetData MsgData, vbString
    ChatDisplay.Text = ChatDisplay.Text & Winsock1(Index).RemoteHostIP & " : " & MsgData & vbCrLf
End Sub

VERSION 5.00
Begin VB.Form frmStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DNS-Bot 1.0"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraRawSend 
      Caption         =   "Send Raw:"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   4920
      Width           =   9375
      Begin VB.CommandButton cmdRawSend 
         Caption         =   "Send"
         Height          =   375
         Left            =   8400
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtRawSend 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   8175
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   9600
      TabIndex        =   4
      Top             =   5880
      Width           =   975
   End
   Begin VB.Timer tmrStatus 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   480
      Top             =   5760
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9600
      TabIndex        =   3
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   9600
      TabIndex        =   2
      Top             =   4920
      Width           =   975
   End
   Begin VB.Frame fraLog 
      Caption         =   "Log:"
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      Begin VB.TextBox txtLog 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   240
         Width           =   10215
      End
   End
   Begin DNS_Bot.IRClient IRC 
      Left            =   0
      Top             =   5760
      _ExtentX        =   767
      _ExtentY        =   767
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private tServer() As Server

Private Const strChan As String = "#opers"

Private Type Server
    IP As String
    Name As String
    Load As Long
End Type

Private Sub cmdConnect_Click()
    cmdConnect.Enabled = False
    IRC.Connect "baller-srv1", 6667, "OpenIRCNet", "OINBot"
    cmdDisconnect.Enabled = True
    tmrStatus.Enabled = True
End Sub

Private Sub cmdDisconnect_Click()
    cmdDisconnect.Enabled = False
    Msg "Close request issued from console."
    IRC.Quit ("Close request issued from console.")
    cmdConnect.Enabled = True
End Sub

Private Sub cmdExit_Click()
    cmdDisconnect.Enabled = False
    Msg "Exit request issued from console."
    IRC.Quit ("Exit request issued from console.")
    cmdConnect.Enabled = True
    Unload Me
End Sub

Private Sub cmdRawSend_Click()
    IRC.RawSend txtRawSend.Text
    txtRawSend.Text = ""
End Sub

Private Sub IRC_ConnectComplete()
    Msg "***Connected."
    Msg "Attempting to /oper"
    IRC.RawSend "OPER OIN botmaster1"
    IRC.RawSend "MAP"
    IRC.Message strChan, "OpenIRCNet Bot (" & App.Major & "." & App.Minor & "." & App.Revision & ") Online."
End Sub

Private Sub Msg(strMsg As String)
    txtLog.Text = txtLog.Text & strMsg & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
End Sub

Private Sub IRC_DataArrival(Data As String)
    Static blnMapStarted As Boolean

    Msg Data
    
    'Check for !exit
    If InStr(Data, "!EXIT") <> 0 Then
        IRC.Message strChan, "Closing..."
        cmdDisconnect.Enabled = False
        Msg "Close request issued from channel."
        IRC.Quit ("Close request issued from channel.")
        cmdConnect.Enabled = True
    End If
    
    'Check for !map
    If InStr(Data, "!MAP") <> 0 Then
        Dim strMsg As String
        Dim i As Long
        For i = 1 To UBound(tServer)
            If Len(strMsg) <> 0 Then
                strMsg = strMsg & " | "
            End If
            strMsg = strMsg & "Server: " & tServer(i).Name & "[" & tServer(i).IP & "] Load: " & tServer(i).Load
        Next i
        IRC.Message strChan, strMsg
    End If
    
    'No operation
    If InStr(Data, "!DNS") <> 0 Then
        If SocketsInitialize() Then
            Dim strName As String
            Dim strAddress As String
            strName = Trim(Mid(Data, InStr(Data, "!DNS") + 4))
            strAddress = GetIPFromHostName(strName)
            IRC.Message strChan, strName & " is " & strAddress
            SocketsCleanup
        End If
    End If
    
    'No operation
    If InStr(Data, "!NE") <> 0 Then
        IRC.Message strChan, "Sorry, this function not yet implemented."
    End If
    
    'Process map data
    If InStr(Data, "<MAP>") <> 0 Then
        If blnMapStarted = False Then
            'Clear array
            blnMapStarted = True
            ReDim tServer(0)
            ReDim tServer(1 To 1)
        Else
            ReDim Preserve tServer(1 To UBound(tServer) + 1)
        End If
        Data = Mid(Data, InStr(Data, "<MAP>") + 6)
        Data = Replace(Data, "(", "")
        Data = Replace(Data, ")", "")
        Data = Replace(Data, "|-", "")
        Data = Replace(Data, "`-", "")
        Data = Replace(Data, "`", "")
        Data = Replace(Data, "|", "")
        Do Until InStr(Data, "  ") = 0
            Data = Replace(Data, "  ", " ")
        Loop
        Dim sData() As String
        Data = Trim(Data)
        sData = VBA.Split(Data, " ")
        tServer(UBound(tServer)).Name = sData(0)
        tServer(UBound(tServer)).Load = sData(1)
        
        If SocketsInitialize() Then
            tServer(UBound(tServer)).IP = GetIPFromHostName(sData(0))
            SocketsCleanup
        End If
    End If
    
    If InStr(Data, "</MAP>") <> 0 Then
        blnMapStarted = False
    End If
    
    DoEvents
End Sub

Private Sub tmrStatus_Timer()
    If IRC.Connected = False Then
        Msg "***Disconnected."
        tmrStatus.Enabled = False
    End If
End Sub

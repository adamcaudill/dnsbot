VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl IRClient 
   CanGetFocus     =   0   'False
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   435
   InvisibleAtRuntime=   -1  'True
   Picture         =   "IRClient.ctx":0000
   ScaleHeight     =   435
   ScaleWidth      =   435
   Begin MSWinsockLib.Winsock SockIdent 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock SockIRC 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "IRClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***************************************************************************
'**                           IRClient (SquIRC)                           **
'***************************************************************************
'
'Yes, its finally arrived. After months and months of rewrites, bug bashing
'and testing it to the limit the squIRC control source is being released to
'you :-D
'
'If you want to learn something, read the comments, and also visit
'http://bishop.mc.duke.edu/irchelp/rfc/rfc.html
'to learn about the IRC protocol
'
'The code displayed here is in the public domain. Do what you want with it,
'with or without credit I care not. Anyways enough talk :-)

Dim mIdent As Boolean           'Do we use Ident server?
Dim mAutoPing As Boolean        'Do we automatically reply to pings to keep client alive?
Dim mOutParse As Boolean        'Do we parse outgoing text as if it was incoming?

Dim mNick As String             'The nickname used by the client
Dim mUserName As String         'The username used by the client that shows up in WHOIS
Dim mServer As String           'The server to connect to

Dim mFinger As String           'A FINGER reply message
Dim mVersion As String          'A VERSION reply message. The control adds "(squIRC.OCX)" to the end automatically
Dim mTStamp As String           'Timestamp to return the incoming parsed text with

Public Event ConnectComplete()              'Event fires when a successful attempt has been made with the server
Public Event DataArrival(Data As String)    'Incoming data is parsed and this event fires

'******************
'*** PROPERTIES ***
'******************

'Let = allow the application to SET a property
'Get = allow the application to READ a property
'Use just like any property for any normal control, e.g Form.Caption

'See mFinger above
Public Property Let FingerMsg(sFinger As String)
    mFinger = sFinger
End Property
Public Property Get FingerMsg() As String
    FingerMsg = mFinger
End Property
'See mVersion above
Public Property Let VersionMsg(sVersion As String)
    mVersion = sVersion
End Property
Public Property Get VersionMsg() As String
    VersionMsg = mVersion
End Property
'See mIdent above
Public Property Let Ident(bIdent As Boolean)
    mIdent = bIdent
End Property
Public Property Get Ident() As Boolean
    Ident = mIdent
End Property
'See mAutoPing above
Public Property Let AutoPing(bAuto As Boolean)
    mAutoPing = bAuto
End Property
Public Property Get AutoPing() As Boolean
    AutoPing = mAutoPing
End Property
'See mOutParse above
Public Property Let OutParse(bOut As Boolean)
    mOutParse = bOut
End Property
Public Property Get OutParse() As Boolean
    OutParse = mOutParse
End Property
'See mTStamp above
Public Property Let TimeStamp(sStamp As String)
    mTStamp = sStamp
End Property
Public Property Get TimeStamp() As String
    TimeStamp = mTStamp
End Property
'Read-only (no Property Let) of whether the client is connected or not
Public Property Get Connected() As Boolean
    If SockIRC.State <> sckConnected Then
        Connected = False
    Else
        Connected = True
    End If
End Property

'*************************
'*** USERCONTROL/PROPS ***
'*************************

'When the usercontrol's properties are intialialized, this event fires
'We can set default values for any variables we need
Private Sub UserControl_InitProperties()
    mAutoPing = True
    mTStamp = "hh:nn"
End Sub

'Introducing the PropBag
'This is what we see when we are creating forms in the IDE
'The little list of properties that shows up in the properties window

'ReadProperties allows us to grab information FROM the properties window
'and store in our local variables

'SYNTAX : Variable = PropBag.ReadProperty(PropertyName, DefaultValue)

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mAutoPing = PropBag.ReadProperty("AutoPing", True)
    mVersion = PropBag.ReadProperty("VersionMsg", UserControl.Name)
    mFinger = PropBag.ReadProperty("FingerMsg", "Hello world!")
    mIdent = PropBag.ReadProperty("Ident", True)
    mOutParse = PropBag.ReadProperty("OutParse", False)
    mTStamp = PropBag.ReadProperty("TimeStamp", "hh:nn")
End Sub

'I want the usercontrol to have a fixed size
Private Sub UserControl_Resize()
    UserControl.Width = 435
    UserControl.Height = 435
End Sub

'When the application closes, or the control is unloaded, this event fires
'We use it to close any open winsocks before the control is destroyed
Private Sub UserControl_Terminate()
    SockIdent.Close
    Quit
End Sub

'See PropBag info above

'WriteProperties allows us to write values TO the properties window
'from stored local variables

'SYNTAX : PropBag.WriteProperty PropertyName, Value, DefaultValue

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "AutoPing", mAutoPing, True
    PropBag.WriteProperty "VersionMsg", mVersion, UserControl.Name
    PropBag.WriteProperty "FingerMsg", mFinger, "Hello world!"
    PropBag.WriteProperty "Ident", mIdent, True
    PropBag.WriteProperty "Outparse", mOutParse, False
    PropBag.WriteProperty "TimeStamp", mTStamp, "hh:nn"
End Sub

'*************************
'*** PUBLIC SUBS/FUNCS ***
'*************************

'These subs and functions will be visible and useable by the application
'on which the control lies. They will be seen as METHODS and show up in
'VB Intellisense

'Connect to server
Public Sub Connect(sServer As String, lPort As Long, sName As String, sNick As String)
    If SockIRC.State <> sckClosed Then
        'If the sock is not closed, raise an error to indicate this
        Err.Raise vbObjectError + 10101, , "Client not ready for connection"
    Else
        'Start the ball rolling, telling the IRC winsock to connect
        'to the specified server/port
        SockIRC.Connect sServer, lPort
        'Set all the local variables here
        mServer = LCase$(sServer)
        mNick = sNick
        mUserName = sName
        'Let Windows do its stuff
        DoEvents
        'Make sure the Ident winsock is closed
        If SockIdent.State <> sckClosed Then SockIdent.Close
    End If
End Sub

'Send server a QUIT message
Public Sub Quit(Optional sReason As String)
    'Close the Ident winsock if necessary
    If SockIdent.State <> sckClosed Then SockIdent.Close
    If SockIRC.State = sckConnected Then
        If Len(sReason) > 0 Then
            'Quit with specified reason
            Send "QUIT :" & sReason
        Else
            'Quit with default "Leaving" reason
            Send "QUIT :Leaving"
        End If
        'Let Windows do its stuff
        DoEvents
    End If
    'Close the winsock
    SockIRC.Close
End Sub

'Send a regular chat message
Public Sub Message(sTarget As String, sText As String)
    If SockIRC.State <> sckConnected Then
        'No connection, raise an error
        Err.Raise vbObjectError + 10102, , "Client not connected to IRC"
    Else
        'Send the PRIVMSG command to specified target
        Send "PRIVMSG " & sTarget & " :" & sText
    End If
End Sub

'Send an action-style message
Public Sub Action(sTarget As String, sText As String)
    If SockIRC.State <> sckConnected Then
        'No connection, raise an error
        Err.Raise vbObjectError + 10102, , "Client not connected to IRC"
    Else
        'Action commands are merely regular PRIVMSG commands
        'with chr$(1) on either end, and the text "ACTION"
        Send "PRIVMSG " & sTarget & " :" & Chr$(1) & "ACTION " & sText & Chr$(1)
    End If
End Sub

'Send a NOTICE message
Public Sub Notice(sTarget As String, sText As String)
    If SockIRC.State <> sckConnected Then
        Err.Raise vbObjectError + 10102, , "Client not connected to IRC"
    Else
        Send "NOTICE " & sTarget & " :" & sText
    End If
End Sub

'Request to join a channel, with parameters if required (e.g. a password)
Public Sub Join(sChannel As String, Optional sParams As String)
    If SockIRC.State <> sckConnected Then
        Err.Raise vbObjectError + 10102, , "Client not connected to IRC"
    Else
        SockIRC.SendData "JOIN " & sChannel & " " & sParams & vbCrLf
    End If
End Sub

'Leave a channel
Public Sub Part(sChannel As String)
    If SockIRC.State <> sckConnected Then
        Err.Raise vbObjectError + 10102, , "Client not connected to IRC"
    Else
        Send "PART " & sChannel
    End If
End Sub

'Invite another user to join a channel you are on
'Some channels are +i (invite-only) so may need to use this
Public Sub Invite(sChannel As String, sNick As String)
    If SockIRC.State <> sckConnected Then
        Err.Raise vbObjectError + 10102, , "Client not connected to IRC"
    Else
        SockIRC.SendData "INVITE " & sNick & " " & sChannel & vbCrLf
    End If
End Sub

'Change the topic of a channel you are on, if at all possible
Public Sub Topic(sChannel As String, sTopic As String)
    If SockIRC.State <> sckConnected Then
        Err.Raise vbObjectError + 10102, , "Client not connected to IRC"
    Else
        SockIRC.SendData "TOPIC " & sChannel & " :" & sTopic & vbCrLf
    End If
End Sub

'Send completely raw text to the server, for example a command not
'supported by this control, e.g. "STATS"
Public Sub RawSend(sData As String)
    If SockIRC.State <> sckConnected Then
        Err.Raise vbObjectError + 10102, , "Client not connected to IRC"
    Else
        SockIRC.SendData sData & vbCrLf
    End If
End Sub

'Common WHOIS command used to retrieve information about a specific user
Public Sub Whois(sNick As String)
    If SockIRC.State <> sckConnected Then
        Err.Raise vbObjectError + 10102, , "Client not connected to IRC"
    Else
        SockIRC.SendData "WHOIS " & sNick & vbCrLf
    End If
End Sub

'Similar to WHOIS, but for nicks that recently left IRC
Public Sub Whowas(sNick As String)
    If SockIRC.State <> sckConnected Then
        Err.Raise vbObjectError + 10102, , "Client not connected to IRC"
    Else
        SockIRC.SendData "WHOWAS " & sNick & vbCrLf
    End If
End Sub

'Kick a user from a channel which you are an operator on, with a reason if required
Public Sub Kick(sChannel As String, sNick As String, Optional sReason As String)
    If SockIRC.State <> sckConnected Then
        Err.Raise vbObjectError + 10102, , "Client not connected to IRC"
    Else
        SockIRC.SendData "KICK " & sChannel & " " & sNick & " :" & sReason & vbCrLf
    End If
End Sub

'Alter a channel or usermode
Public Sub Mode(sChannel As String, sParams1 As String, Optional sParams2 As String)
    If SockIRC.State <> sckConnected Then
        Err.Raise vbObjectError + 10102, , "Client not connected to IRC"
    Else
        SockIRC.SendData "MODE " & sChannel & " " & sParams1 & " " & sParams2 & vbCrLf
    End If
End Sub

'**************************
'*** PRIVATE SUBS/FUNCS ***
'**************************

'These functions and subs are not visible to the application, they are
'encapsulated by the control and are for internal use only

'Send some data
Private Sub Send(sData As String)
    'Send the data
    SockIRC.SendData sData & vbCrLf
    If mOutParse Then
        'If OutParse is set True, we need to create a fake hostmask
        'and then parse it as normal
        Dim strOut As String
        strOut = ":" & mNick & "!~" & mNick & "@host" & SockIRC.LocalIP & ".someisp.com "
        Parse strOut & sData & vbCrLf
    End If
End Sub

'Parse some data, i.e. turn it from raw IRC protocol format into
'something more useable. The method I chose was <COMMAND PARAM:whatever PARAM:whatever>
'I have been much criticised for use of this output style
'Feel free to alter as you wish

Private Sub Parse(DataString As String)

    Dim sLine() As String       'Array to hold each individual line
    Dim sWord() As String       'Array to hold each individual word of a line
    Dim i As Long               'Counter
    Dim j As Long               'Counter
    
    Dim uNick As String         'The nickname from which the command originates
    Dim uText As String         'The text of the commands
    Dim strOut As String        'The outgoing text
    
    On Error Resume Next        'Errors fly around like crazy, just ignore them
    
    Debug.Print DataString
    
    'Break the block of text into individual lines
    Split DataString, sLine, vbCrLf
    
    'Loop through the lines
    For i = LBound(sLine) To UBound(sLine)
        
        'Split the current line into individual words
        Split sLine(i), sWord, " "

        'If Mid$(LCase$(sWord(0)), 2) = mServer Then
        If Mid$(LCase$(sWord(0)), 2) = "irc.shadowofthebat.com" Then
            'Dealing with server messages
            
            'For this block, the variable j is used to specify how many words
            'are joined to the end of the outgoing string
            
            j = 3
            If sWord(1) = "NOTICE" Then
                'Server notice message
                strOut = "<NOTICE From:" & mServer & " To:" & sWord(2) & "> "
                sWord(3) = Mid$(sWord(3), 2)
            Else
                'Some other server message (much more common)
                Select Case CInt(sWord(1))
                Case 1
                    'Welcome message
                    strOut = "<INFO From:" & mServer & "> "
                    RaiseEvent ConnectComplete
                    
                Case 6
                    'MAP
                    strOut = "<MAP>"
                Case 7
                    'End MAP
                    strOut = "</MAP>"
                Case 2, 3, 250 To 266
                    'Various messages here
                    'Early versions did parse them all differently
                    'but eventually I decided to lump them all together
                    'sorry :-)
                    strOut = "<INFO From:" & mServer & "> "
                
                Case 311
                    'The WHOIS replies I deliberately separated so that
                    'the application has some way to distinguish between each
                    'line
                    'These replies can be generated by either a WHOIS or WHOWAS request
                    strOut = "<WHOIS1 Nick:" & sWord(3) & "> "
                Case 319
                    strOut = "<WHOIS2 Nick:" & sWord(3) & "> "
                Case 312
                    strOut = "<WHOIS3 Nick:" & sWord(3) & "> "
                Case 307
                    strOut = "<WHOIS4 Nick:" & sWord(3) & "> "
                Case 317
                    strOut = "<WHOIS5 Nick:" & sWord(3) & "> "
                
                Case 308 To 310, 313 To 316, 318, 366
                    'Other WHOIS info, which tends to be network-specific
                    'Or the "END OF WHOIS" reply
                    strOut = "<WHOISWAS Nick:" & sWord(3) & "> "
                
                Case 372, 375
                    'Start MOTD (Message of the Day)
                    'This is sent to the client upon initial connection
                    strOut = "<MOTD From:" & mServer & "> "
                Case 376
                    'End of MOTD
                    strOut = "<END_MOTD>"
                    j = 1
                    
                Case 353
                    'Names list for a specific channel
                    strOut = "<NAMES Chan:" & sWord(4) & "> "
                    sWord(5) = Mid$(sWord(5), 2)
                    j = 5
                Case 366
                    'End of names list
                    strOut = "<END_NAMES Chan:" & sWord(3) & ">"
                    j = 0
                    
                Case 433
                    'Nickname in use
                    strOut = "<ERR_NICKINUSE Nick:" & sWord(2) & "> "
                    j = 0
                Case 432
                    'Erroneous nickname
                    strOut = "<ERR_INVALIDNICK Nick:" & sWord(2) & "> "
                    j = 0
                    
                Case 443
                    'An error, tried to invite a nick who was already in the channel
                    strOut = "<ERR_ALREADYONCHANNEL Nick:" & sWord(3) & " Chan:" & sWord(4) & ">"
                    j = 0
                Case 401
                    'Can be triggered by WHOIS, WHOWAS, INVITE, MODE, etc
                    'Basically a "nick does not exist" error
                    strOut = "<ERR_NOSUCHNICK Nick:" & sWord(3) & ">"
                    j = 0
                End Select
            End If
                
            'Reconstruct string if needed
            If j Then strOut = strOut & VB6.Join(sWord, j)
            
        ElseIf Right$(sWord(0), 5) = "ERROR" Then
            'Error eg
            ': ERROR :Closing Link: AutoVB[unknown@255.255.255.255] (You are not authorized to use this server)
            strOut = VB6.Join(sWord)
            
        ElseIf Right$(sWord(0), 4) = "PING" Then
            'Server is PINGING us, better ping back
            If mAutoPing Then
                strOut = "PONG " & sWord(1)
                Send strOut
                DoEvents
                Exit Sub
            End If
        
        ElseIf Right$(sWord(0), 6) = "NOTICE" Then
            'Authentication notices
            'can probably ignore them
            strOut = "<ID & AUTH>"
            
        ElseIf Left$(sWord(0), 1) = ":" Then
            'This is the meat of the parser, stuff from other users
            
            'First, grab the nick of the sender
            j = InStr(1, sWord(0), "!") - 2
            uNick = Mid$(sWord(0), 2, j)
            
            'So now we know who did it,
            'but what did they do?
            
            Select Case sWord(1)
            Case "PRIVMSG"
                'PRIVMSG is the most common, it is any chat message, action,
                'CTCP (client-to-client-protocol) text
                sWord(3) = UCase$(sWord(3))
            
                If sWord(3) = ":" & Chr$(1) & "ACTION" Then
                    'They are doing an action
                    strOut = "<ACTION From:" & uNick & " To:" & sWord(2) & "> "
                    strOut = strOut & VB6.Join(sWord, 4)
                
                ElseIf sWord(3) = ":" & Chr$(1) & "PING" Then
                    'CTCP ping
                    strOut = "<PING From:" & uNick & " To:" & sWord(2) & ">"
                    Send "NOTICE " & uNick & " :" & Chr$(1) & "PING " & sWord(4)
                
                ElseIf sWord(3) = ":" & Chr$(1) & "TIME" & Chr$(1) Then
                    'CTCP time
                    strOut = "<TIME From:" & uNick & " To:" & sWord(2) & ">"
                    Send "NOTICE " & uNick & " :" & Chr$(1) & "TIME " & Now & Chr$(1)
                
                ElseIf sWord(3) = ":" & Chr$(1) & "VERSION" & Chr$(1) Then
                    'CTCP version
                    strOut = "<VERSION From:" & uNick & " To:" & sWord(2) & ">"
                    Send "NOTICE " & uNick & " :" & Chr$(1) & "VERSION " & mVersion & Chr$(1)
                
                ElseIf sWord(3) = ":" & Chr$(1) & "FINGER" & Chr$(1) Then
                    'CTCP finger
                    strOut = "<FINGER From:" & uNick & " To:" & sWord(2) & ">"
                    Send "NOTICE " & uNick & " :" & Chr$(1) & "FINGER " & mFinger & Chr$(1)
                
                ElseIf Left$(sWord(3), 2) = ":" & Chr$(1) Then
                    'Unknown CTCP (they happen)
                    strOut = "<CTCP From:" & uNick & " To:" & sWord(2) & ">"
                    sWord(3) = Mid$(sWord(3), 2)
                    strOut = strOut & VB6.Join(sWord, 3)
                
                Else
                    'Regular chat
                    strOut = "<MESSAGE From:" & uNick & " To:" & sWord(2) & "> "
                    sWord(3) = Mid$(sWord(3), 2)
                    strOut = strOut & VB6.Join(sWord, 3)
                End If
        
            Case "NOTICE"
                'Noticed
                strOut = "<NOTICE From:" & uNick & " To:" & sWord(2) & "> "
                sWord(3) = Mid$(sWord(3), 2)
                strOut = strOut & VB6.Join(sWord, 3)
                
            Case "JOIN"
                'They joined the channel
                sWord(2) = Mid$(sWord(2), 2)
                strOut = "<JOIN Nick:" & uNick & " Chan:" & sWord(2) & ">"
                
            Case "PART"
                'They left the channel
                sWord(2) = Mid$(sWord(2), 2)
                strOut = "<PART Nick:" & uNick & " Chan:" & sWord(2) & ">"
                
            Case "TOPIC"
                'They changed the topic
                strOut = "<TOPIC Nick:" & uNick & " Chan:" & sWord(2) & "> "
                sWord(3) = Mid$(sWord(3), 2)
                strOut = strOut & VB6.Join(sWord, 3)
                
            Case "MODE"
                'They changed a mode
                strOut = "<MODE Nick:" & uNick & " Chan:" & sWord(2) & "> "
                strOut = strOut & VB6.Join(sWord, 3)
                
            Case "QUIT"
                'The quit IRC
                strOut = "<QUIT Nick:" & uNick & ">"
                
            Case "NICK"
                'Nick change
                sWord(2) = Mid$(sWord(2), 2)
                strOut = "<NICK Old:" & uNick & " New:" & sWord(2) & ">"
                
            Case "KICK"
                'User was kicked
                strOut = "<KICK Kicked:" & uNick & " By:" & sWord(3) & " Chan:" & sWord(2) & "> "
                strOut = strOut & VB6.Join(sWord, 4)
                
            Case Else
                'Some other message
                strOut = "Unknown command : <" & sWord(1) & ">"
                strOut = strOut & VB6.Join(sWord, 3)
            End Select
            
        End If
        
        DoEvents
        
        'If we have stuff to return, return it
        If Len(strOut) > 2 Then
            'Add timestamp
            If Len(mTStamp) > 0 Then strOut = "[" & Format$(Time, mTStamp) & "] " & strOut
            'Raise the event
            RaiseEvent DataArrival(strOut)
        End If
        strOut = ""
    Next i
End Sub

'**********************
'*** WINSOCK EVENTS ***
'**********************

Private Sub SockIRC_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim DataString As String
    'Grab the text
    SockIRC.GetData DataString, vbString
    DoEvents
    'Parse the text
    Parse DataString
End Sub

Private Sub SockIRC_Connect()
    'Send the NICK message
    SockIRC.SendData "NICK " & mNick & vbCrLf
    'Send the USER message, parameters left with . indicate those used
    'by servers to connect to each other
    SockIRC.SendData "USER " & mNick & " . . :" & mUserName & vbCrLf
    
    If mIdent Then
        'If we want Ident, listen on 113, the Ident port
        SockIdent.LocalPort = 113
        SockIdent.Listen
    End If
End Sub

Private Sub SockIRC_Close()
    'Close the winsock
    SockIRC.Close
End Sub

Private Sub SockIdent_ConnectionRequest(ByVal requestID As Long)
    'When the server tried to connect to the Ident port, close
    'the ident winsock and accept
    SockIdent.Close
    SockIdent.Accept requestID
End Sub

Private Sub SockIdent_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim DataString As String

    'Grab the text
    SockIdent.GetData DataString, vbString
    DoEvents
    
    Dim sBack As String
    'Send back the text with our own stuff added on the end.
    'We add UNIX, yes this is a lie but everybody does it, even mIRC
    sBack = Left$(DataString, Len(DataString) - 2) & " : USERID : UNIX : " & mNick & vbCrLf
    'Send back the reply
    SockIdent.SendData sBack
End Sub

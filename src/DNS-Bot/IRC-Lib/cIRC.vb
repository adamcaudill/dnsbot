Imports System.Net.Sockets
Imports System.Threading

Public Class cIRC

    Public Event NickChange(ByVal UserName As String, ByVal strNewUserName As String, ByVal strUserMask As String)
    Public Event DataArrival(ByVal Data As String)
    Public Event DataArrival_StrArray(ByVal Data() As String)
    Public Event PrivateMessage(ByVal Data As String, ByVal strRecievedFrom As String, ByVal strUserMask As String)
    Public Event ChannelMessage(ByVal Data As String, ByVal strChannel As String, ByVal strUserMask As String)
    Public Event ChannelJoin(ByVal UserName As String, ByVal strChannel As String, ByVal strUserMask As String)
    Public Event ChannelPart(ByVal UserName As String, ByVal strChannel As String, ByVal strUserMask As String)
    Public Event ConnectComplete()

    Private m_strVersion As String

    Private m_strIDENTString As String

    Private m_strRealName As String

    Private m_strNickname As String

    Private m_intPort As Integer

    Private m_strServer As String

    Private m_sckIRC As Socket

    Private m_strChannel As String

    Public Property Server() As String
        Get
            Return m_strServer
        End Get
        Set(ByVal Value As String)
            m_strServer = Value
        End Set
    End Property

    Public Property Port() As Integer
        Get
            Return m_intPort
        End Get
        Set(ByVal Value As Integer)
            m_intPort = Value
        End Set
    End Property

    Public ReadOnly Property Channel() As String
        Get
            Return m_strChannel
        End Get
    End Property

    Public Property Nickname() As String
        Get
            Return m_strNickname
        End Get
        Set(ByVal Value As String)
            m_strNickname = Value
        End Set
    End Property

    Public Property RealName() As String
        Get
            Return m_strRealName
        End Get
        Set(ByVal Value As String)
            m_strRealName = Value
        End Set
    End Property

    Public Property IDENTString() As String
        Get
            Return m_strIDENTString
        End Get
        Set(ByVal Value As String)
            m_strIDENTString = Value
        End Set
    End Property

    Public Property Version() As String
        Get
            Return m_strVersion
        End Get
        Set(ByVal Value As String)
            m_strVersion = Value
        End Set
    End Property

    Private Sub Listen()
        Do While (m_sckIRC.Connected)
            Dim bytBuffer(4095) As Byte 'BUGBUG -- This limits the max data received at once to 4096 bytes
            m_sckIRC.Receive(bytBuffer)
            Dim strLines() As String = Text.Encoding.ASCII.GetString(bytBuffer).Replace(Chr(10), "").Split(ControlChars.Cr)

            Dim i As Integer
            For i = 0 To strLines.GetUpperBound(0)
                If strLines(i).Length > 1 Then
                    Debug.WriteLine(strLines(i))
                    'RaiseEvent DataArrival(strLines(i))
                    'Process each line of data

                    'Split the line into word for we can see just what we are dealing with
                    Dim strWord() As String = strLines(i).Split(" ")

                    'Check to see if the first word is the server name
                    'If strWord(0).Substring(2).ToLower = m_strServer Then                ' <----- NOTE: Make sure this line is uncommented. (Adam)
                    If strWord(0).ToLower = ":polyfractal.ath.cx" Then    ' <----- NOTE: The is to fix a local DNS issue. Do not use with this line. (Adam)
                        'This is a server message
                        If strWord(1) = "NOTICE" Then
                            'Server NOTICE
                            Dim strMsg As String
                            strMsg = strLines(i).Substring(InStr(strLines(i), "NOTICE") + "NOTICE".Length)
                            RaiseEvent DataArrival("Sever: " & strMsg)
                        Else
                            'Some other server message (much more common)
                            Select Case CInt(strWord(1))
                                Case 1
                                    'Server welcome message
                                    Dim strMsg As String
                                    strMsg = strLines(i).Substring(InStr(strLines(i), m_strNickname) + m_strNickname.Length + 1)
                                    RaiseEvent DataArrival("Sever: Welcome. (" & strMsg & ")")
                                Case 6
                                    Dim strMsg As String
                                    strMsg = strLines(i).Substring(InStr(strLines(i), m_strNickname) + m_strNickname.Length + 1)
                                    RaiseEvent DataArrival("MAP: " & strMsg)
                                Case 7
                                    RaiseEvent DataArrival("MAP: End of /MAP")
                                Case 2, 3, 251, 255, 265, 266
                                    'Host info
                                    Dim strMsg As String
                                    strMsg = strLines(i).Substring(InStr(strLines(i), m_strNickname) + m_strNickname.Length + 1)
                                    RaiseEvent DataArrival("Host Info: " & strMsg)
                                Case 4, 5, 252, 254
                                    'Host info
                                    Dim strMsg As String
                                    strMsg = strLines(i).Substring(InStr(strLines(i), m_strNickname) + m_strNickname.Length)
                                    RaiseEvent DataArrival("Host Info: " & strMsg)
                                Case 332
                                    'Channel topic
                                    Dim strMsg As String
                                    strMsg = strLines(i).Substring(InStr(strLines(i), strWord(3)) + strWord(3).Length + 1)
                                    RaiseEvent DataArrival(strWord(3) & ": (Topic) " & strMsg)
                                Case 372
                                    'Start MOTD (Message of the Day)
                                    'This is sent to the client upon initial connection
                                    Dim strMsg As String
                                    strMsg = strLines(i).Substring(InStr(strLines(i), m_strNickname) + m_strNickname.Length + 1)
                                    RaiseEvent DataArrival("MOTD: " & strMsg)
                                Case 375
                                    RaiseEvent DataArrival("MOTD: MOTD Follows:")
                                Case 376
                                    'End of MOTD
                                    RaiseEvent DataArrival("MOTD: End of MOTD")
                                    RaiseEvent ConnectComplete()
                                Case 353
                                    'Raise NAMES
                                    RaiseEvent DataArrival_StrArray(strWord)
                                Case Else
                                    'Debug.WriteLine("------" & strWord(3) & " " & strWord(4) & " " & strWord(1) & "-------")
                            End Select

                        End If
                    ElseIf strWord(0).EndsWith("PING") = True Then
                        'Server is PINGING us, better ping back
                        Send("PONG " & strWord(1).Replace(":", ""))
                    ElseIf strWord(0).StartsWith(":") Then
                        'These should all be messages
                        Dim strUserName As String
                        If InStr(strWord(0), "!") <> 0 Then
                            strUserName = strWord(0).Replace(":", "").Substring(0, InStr(strWord(0).Replace(":", ""), "!") - 1)
                        Else
                            strUserName = strWord(0).Replace(":", "")
                        End If
                        Select Case strWord(1)
                            Case "PRIVMSG"
                                'Received a message
                                If strWord(3) = ":" & Chr(1) & "ACTION" Then
                                    'Action
                                    Dim strMsg As String
                                    strMsg = strLines(i).Substring(InStr(strLines(i), strWord(3)) + strWord(3).Length).Replace(Chr(1), "")
                                    RaiseEvent DataArrival(strWord(2) & ": (" & strWord(0).Replace(":", "") & ") [Action] " & strMsg)
                                ElseIf strWord(3) = ":" & Chr(1) & "PING" Then
                                    'CTCP ping request
                                    Send("NOTICE " & strUserName & " :" & Chr(1) & "PING " & strWord(4))
                                    RaiseEvent DataArrival(strWord(0).Replace(":", "") & ": CTCP PING")
                                ElseIf strWord(3) = ":" & Chr(1) & "TIME" & Chr(1) Then
                                    'CTCP time request
                                    Send("NOTICE " & strUserName & " :" & Chr(1) & "TIME " & Now & Chr(1))
                                    RaiseEvent DataArrival(strWord(0).Replace(":", "") & ": CTCP TIME")
                                ElseIf strWord(3) = ":" & Chr(1) & "VERSION" & Chr(1) Then
                                    'CTCP Version request
                                    Send("NOTICE " & strUserName & " :" & Chr(1) & "VERSION " & m_strVersion & Chr(1))
                                    RaiseEvent DataArrival(strWord(0).Replace(":", "") & ": CTCP VERSION")
                                ElseIf strWord(2) = m_strNickname Then
                                    Dim strMsg As String
                                    strMsg = strLines(i).Substring(InStr(strLines(i), strWord(2)) + strWord(2).Length + 1)
                                    RaiseEvent DataArrival("PM: (" & strWord(0).Replace(":", "") & ") " & strMsg)
                                    RaiseEvent PrivateMessage(strMsg, Left$(strWord(0).Replace(":", ""), InStr(strWord(0).Replace(":", ""), "!") - 1), strWord(0).Replace(":", ""))
                                Else
                                    'Normal channel message
                                    Dim strMsg As String
                                    strMsg = strLines(i).Substring(InStr(strLines(i), strWord(2)) + strWord(2).Length + 1)
                                    RaiseEvent DataArrival(strWord(2) & ": (" & strWord(0).Replace(":", "") & ") " & strMsg)
                                    RaiseEvent ChannelMessage(strMsg, strWord(2), strWord(0).Replace(":", ""))
                                End If
                            Case "JOIN"
                                'a user has joinded a channel
                                RaiseEvent DataArrival(strWord(2).Replace(":", "") & ": JOIN: " & strWord(0).Replace(":", ""))
                                If strUserName <> m_strNickname Then
                                    RaiseEvent ChannelJoin(strUserName, strWord(2).Replace(":", ""), strWord(0).Replace(":", ""))
                                End If

                            Case "NICK"
                                'a user has changed his nick
                                RaiseEvent DataArrival(strUserName & ": NICK: " & strWord(2).Replace(":", "") & "  " & strWord(0).Replace(":", ""))
                                If strUserName <> m_strNickname Then
                                    RaiseEvent NickChange(strUserName, strWord(2).Replace(":", ""), strWord(0).Replace(":", ""))
                                End If

                            Case "PART"
                                'a user has joinded a channel
                                RaiseEvent DataArrival(strWord(2).Replace(":", "") & ": PART: " & strWord(0).Replace(":", ""))
                                If strUserName <> m_strNickname Then
                                    RaiseEvent ChannelPart(strUserName, strWord(2).Replace(":", ""), strWord(0).Replace(":", ""))
                                End If
                        End Select
                    End If
                End If
            Next
        Loop
    End Sub

    Public Sub Send(ByVal strCommand As String)
        m_sckIRC.Send(Text.Encoding.ASCII.GetBytes(strCommand & ControlChars.CrLf))
    End Sub

    Public Sub SendMessage(ByVal strCommand As String, ByVal strTarget As String)
        m_sckIRC.Send(Text.Encoding.ASCII.GetBytes("PRIVMSG " & strTarget & " :" & strCommand & ControlChars.CrLf))
    End Sub

    Public Sub Raw(ByVal data As String)
        'What is this for?
        'Seems Raw & Send do the same thing,
        'Perhaps we should make Send private? - Adam
        m_sckIRC.Send(Text.Encoding.ASCII.GetBytes(data & ControlChars.CrLf))
    End Sub

    Public Sub Connect()
        RaiseEvent DataArrival("***Connecting")
        m_sckIRC = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
        m_sckIRC.Connect(New System.Net.IPEndPoint(System.Net.Dns.Resolve(m_strServer).AddressList(0), m_intPort))

        Dim thdListen As Thread = New Thread(AddressOf Listen)
        thdListen.IsBackground = True
        thdListen.Start()

        'Send the NICK message
        Send("NICK " & m_strNickname)
        'Send the USER message, parameters left with . indicate those used
        'by servers to connect to each other
        Send("USER " & m_strNickname & " . . :" & m_strRealName)
    End Sub

    Public Sub Join(ByVal strChannel As String, Optional ByVal strParams As String = "")
        Send("JOIN " & strChannel & " " & strParams)
        m_strChannel = strChannel
        RaiseEvent DataArrival("Join: Attempting to join " & strChannel)
    End Sub

    Public Sub PartJoin(ByVal strChannel As String, Optional ByVal strParams As String = "")
        Send("PART " & m_strChannel)
        m_strChannel = ""
        RaiseEvent DataArrival("PART: Attempting to part " & m_strChannel)

        Send("JOIN " & strChannel & " " & strParams)
        m_strChannel = strChannel
        RaiseEvent DataArrival("Join: Attempting to join " & strChannel)
    End Sub

    Public Sub ReJoin(Optional ByVal strParams As String = "")
        Send("PART " & m_strChannel)
        Send("JOIN " & m_strChannel & " " & strParams)
        RaiseEvent DataArrival("Join: Attempting to rejoin " & m_strChannel)
    End Sub

    Public Sub Quit(Optional ByVal strReason As String = "Leaving()")
        Send("QUIT :" & strReason)
        m_sckIRC.Close()
        RaiseEvent DataArrival("***Connection Closed")
    End Sub

    Public Sub ChangeNick(ByVal strNick As String)
        Send("NICK :" & strNick)
        m_strNickname = strNick
    End Sub
End Class

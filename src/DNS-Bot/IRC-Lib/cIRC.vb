Imports System.Net.Sockets
Imports System.Threading

Public Class cIRC

    Public Event DataArrival(ByVal Data As String)

    Private m_strVersion As String

    Private m_strIDENTString As String

    Private m_strRealName As String

    Private m_strNickname As String

    Private m_intPort As Integer

    Private m_strServer As String

    Private m_sckIRC As Socket

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
        Do While (True)
            Dim bytBuffer(4095) As Byte
            m_sckIRC.Receive(bytBuffer)
            Dim strLines() As String = Text.Encoding.ASCII.GetString(bytBuffer).Split(ControlChars.CrLf)

            Dim i As Integer
            For i = 0 To strLines.GetUpperBound(0)
                If strLines(i).Length > 1 Then
                    RaiseEvent DataArrival(strLines(i))
                    'Process each line of data

                    'Split the line into word for we can see just what we are dealing with
                    Dim strWord() As String = strLines(i).Split(" ")

                    'Check to see if the first word is the server name
                    'If strWord(0).Substring(2).ToLower = m_strServer Then                ' <----- NOTE: Make sure this line is uncommented. (Adam)
                    If strWord(0).Substring(2).ToLower = "irc.shadowofthebat.com" Then    ' <----- NOTE: The is to fix a local DNS issue. Do not use with this line. (Adam)
                        'This is a server message
                        If strWord(1) = "NOTICE" Then
                            'Server NOTICE
                        Else
                            'Some other server message (much more common)
                            Select Case CInt(strWord(1))
                                Case 1
                                    'Server welcome message
                            End Select
                        End If
                    ElseIf strWord(0).EndsWith("PING") = True Then
                        'Server is PINGING us, better ping back
                        Send("PONG " & strWord(1).Replace(":", ""))
                    End If
                End If
            Next
        Loop
    End Sub

    Public Sub Send(ByVal strCommand As String)
        m_sckIRC.Send(Text.Encoding.ASCII.GetBytes(strCommand & ControlChars.CrLf))
    End Sub

    Public Sub Connect()
        m_sckIRC = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
        m_sckIRC.Connect(New System.Net.IPEndPoint(System.Net.Dns.Resolve(m_strServer).AddressList(0), m_intPort))


        'Send the NICK message
        Send("NICK " & m_strNickname)
        'Send the USER message, parameters left with . indicate those used
        'by servers to connect to each other
        Send("USER " & m_strNickname & " . . :" & m_strRealName)

        Dim thdListen As Thread = New Thread(AddressOf Listen)
        thdListen.Start()
    End Sub
End Class

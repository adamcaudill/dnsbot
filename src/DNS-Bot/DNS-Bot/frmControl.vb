Public Class frmControl
    Inherits System.Windows.Forms.Form

    Dim WithEvents IRC As New IRC_Lib.cIRC

    Dim Settings As New clsXMLCfgFile(AppPath() & "Settings.xml")

    Dim blnTestMode As Boolean

    Dim tServer() As Server



    Private m_intPriority As Int32
    Private m_strPriorityPass As String
    Private m_blnIsPriority As Boolean

    Private Structure Server
        Dim IP As String
        Dim Name As String
        Dim Load As Long
        Dim Ignore As Boolean
    End Structure

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents txtReceived As System.Windows.Forms.TextBox
    Friend WithEvents tmrRefresh As System.Windows.Forms.Timer
    Friend WithEvents tmrChallenge As System.Windows.Forms.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.txtReceived = New System.Windows.Forms.TextBox
        Me.tmrRefresh = New System.Windows.Forms.Timer(Me.components)
        Me.tmrChallenge = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'txtReceived
        '
        Me.txtReceived.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtReceived.Location = New System.Drawing.Point(0, 0)
        Me.txtReceived.Multiline = True
        Me.txtReceived.Name = "txtReceived"
        Me.txtReceived.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtReceived.Size = New System.Drawing.Size(584, 336)
        Me.txtReceived.TabIndex = 1
        Me.txtReceived.Text = ""
        Me.txtReceived.WordWrap = False
        '
        'tmrRefresh
        '
        '
        'tmrChallenge
        '
        Me.tmrChallenge.Interval = 1000
        '
        'frmControl
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(584, 331)
        Me.Controls.Add(Me.txtReceived)
        Me.Name = "frmControl"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub IRC_DataArrival(ByVal Data As String) Handles IRC.DataArrival
        txtReceived.Text += Data & ControlChars.CrLf
        If Data.Substring(0, 4) = "MAP:" Then
            ProcMap(Data)
        End If
    End Sub

    Private Sub IRC_ConnectComplete() Handles IRC.ConnectComplete
        IRC.Send("MAP")
        IRC.Join("#DNS-Bot")
        Application.DoEvents()
        IRC.SendMessage("DNS-Bot (v" & Application.ProductVersion & ") Online.", "#DNS-Bot")
        tmrRefresh.Interval = Settings.GetConfigInfo("General", "RefreshDelay", 30)(1) * 1000
        tmrRefresh.Enabled = True
    End Sub

    Private Sub IRC_PrivateMessage(ByVal Data As String, ByVal strChannel As String, ByVal strUserMask As String) Handles IRC.PrivateMessage
        Dim strWord() As String = Data.Split(" ")
        Select Case strWord(0)
            Case "AUTH"
                'we have just been requested to AUTH ourselves to another bot
                IRC.SendMessage("AUTHPASS " & m_strPriorityPass, strChannel)
            Case "AUTHPASS"
                'we have just recieved a auth password, verify
                If strWord(1) = m_strPriorityPass Then
                    tmrChallenge.Enabled = False
                    m_blnIsPriority = False
                    IRC.SendMessage("Transfering Priority to " & strChannel, IRC.Channel)
                    IRC.SendMessage("AUTHGRANTED", strChannel)
                Else
                    IRC.SendMessage("AUTHDENIED", strChannel)
                End If


        End Select
    End Sub

    Private Sub IRC_ChannelMessage(ByVal Data As String, ByVal strChannel As String, ByVal strUserMask As String) Handles IRC.ChannelMessage
        Dim strWord() As String = Data.Split(" ")

        ' <Added by: Adam at: 7/11/2004-05:44:24 on machine: BALLER-STA1>
        'This will support thr new MultiBot mode & accept commands prefixed with the bot name
        If strWord(0) = IRC.Nickname Then
            Data = Data.Substring(InStr(Data, IRC.Nickname) + IRC.Nickname.Length)
            strWord = Data.Split(" ")
        ElseIf strWord(0) = IRC.Nickname & ":" Then
            Data = Data.Substring(InStr(Data, IRC.Nickname) + IRC.Nickname.Length + 1)
            strWord = Data.Split(" ")
        Else
            If CBool(Settings.GetConfigInfo("General", "MultiBot", False)(1)) = True Then
                Exit Sub
            End If
        End If
        ' </Added by: Adam at: 7/11/2004-05:44:25 on machine: BALLER-STA1>

        Select Case strWord(0).ToLower
            Case "!exit"
                If CBool(Settings.GetConfigInfo("Auth", strUserMask, False)(1)) = True Then
                    SaveSettings()
                    IRC.Quit("Leaving(Exit(" & strUserMask & "))[http://sourceforge.net/projects/dnsbot/]")
                Else
                    IRC.SendMessage("You are not authorized to use this command.", strChannel)
                End If
            Case "!resolve", "!dns"
                If CBool(Settings.GetConfigInfo("Auth", strUserMask, False)(1)) = True Then
                    IRC.SendMessage(strWord(1) & " is " & System.Net.Dns.Resolve(strWord(1)).AddressList(0).ToString, strChannel)
                Else
                    IRC.SendMessage("You are not authorized to use this command.", strChannel)
                End If
            Case "!hm"
                IRC.SendMessage("Your hostmask is " & strUserMask, strChannel)
            Case "!die"
                If CBool(Settings.GetConfigInfo("Auth", strUserMask, False)(1)) = True Then
                    SaveSettings()
                    IRC.Quit("Leaving(Die(" & strUserMask & "))[http://sourceforge.net/projects/dnsbot/]")
                    Application.Exit()
                Else
                    IRC.SendMessage("You are not authorized to use this command.", strChannel)
                End If
            Case "!about"
                IRC.SendMessage("I'm DNS-Bot(v" & Application.ProductVersion & "), a load-balancing system for IRC servers that is free & licensed under the GPL. For more information go to http://sourceforge.net/projects/dnsbot/", strChannel)
                IRC.SendMessage("- Developers: Adam Caudill, Zachary Tong - Special Thanks: Andrew Radamis, Paul Crocket, Roland De Meester, OpenIRCNet & everybody else that has helped!", strChannel)
            Case "!nick"
                If CBool(Settings.GetConfigInfo("Auth", strUserMask, False)(1)) = True Then
                    IRC.SendMessage("Changing name to: " & strWord(1), strChannel)
                    IRC.ChangeNick(strWord(1))
                Else
                    IRC.SendMessage("You are not authorized to use this command.", strChannel)
                End If
            Case "!map"
                Dim strMsg As String
                Dim i As Long
                For i = 0 To UBound(tServer)
                    If Len(strMsg) <> 0 Then
                        strMsg = strMsg & " | "
                    End If
                    strMsg += "Server: " & tServer(i).Name & "[" & tServer(i).IP & "] Load: " & tServer(i).Load
                    If tServer(i).Ignore = True Then
                        strMsg += " (IGNORED)"
                    End If
                Next i
                IRC.SendMessage(strMsg, strChannel)
            Case "!highload"
                IRC.SendMessage("Server with the highest load is " & tServer(GetServerHighLoadAsInt()).Name & " at " & tServer(GetServerHighLoadAsInt()).Load & " users.", strChannel)
            Case "!lowload"
                IRC.SendMessage("Server with the lowest load is " & tServer(GetServerLowLoadAsInt()).Name & " at " & tServer(GetServerLowLoadAsInt()).Load & " users.", strChannel)
            Case "!current"
                IRC.SendMessage("Current server is " & GetCurrentServerAsName(), strChannel)
            Case "!refresh"
                If CBool(Settings.GetConfigInfo("Auth", strUserMask, False)(1)) = True Then
                    IRC.SendMessage("Reloading /MAP Data.", strChannel)
                    IRC.Send("MAP")
                Else
                    IRC.SendMessage("You are not authorized to use this command.", strChannel)
                End If
            Case "!ignore"
                If CBool(Settings.GetConfigInfo("Auth", strUserMask, False)(1)) = True Then
                    Dim i As Long
                    Dim intServer As Integer
                    For i = 0 To UBound(tServer)
                        If tServer(i).name = strWord(1) Then
                            intServer = i + 1
                        End If
                    Next
                    If intServer <> 0 Then
                        tServer(intServer - 1).Ignore = True
                        Settings.WriteConfigInfo("Ingore", tServer(intServer - 1).Name, True)
                        IRC.SendMessage("Server added to ignore: " & tServer(intServer - 1).Name, strChannel)
                    Else
                        IRC.SendMessage("Unknown server: " & strWord(1), strChannel)
                    End If
                Else
                    IRC.SendMessage("You are not authorized to use this command.", strChannel)
                End If
            Case "!unignore"
                If CBool(Settings.GetConfigInfo("Auth", strUserMask, False)(1)) = True Then
                    Dim i As Long
                    Dim intServer As Integer
                    For i = 0 To UBound(tServer)
                        If tServer(i).name = strWord(1) Then
                            intServer = i + 1
                        End If
                    Next
                    If intServer <> 0 Then
                        tServer(intServer - 1).Ignore = False
                        Settings.WriteConfigInfo("Ingore", tServer(intServer - 1).Name, False)
                        IRC.SendMessage("Server removed from ignore: " & tServer(intServer - 1).Name, strChannel)
                    Else
                        IRC.SendMessage("Unknown server: " & strWord(1), strChannel)
                    End If
                Else
                    IRC.SendMessage("You are not authorized to use this command.", strChannel)
                End If
            Case "!auth"
                If CBool(Settings.GetConfigInfo("Auth", strUserMask, False)(1)) = True Then
                    Settings.WriteConfigInfo("Auth", strWord(1), True)
                    IRC.SendMessage("User added.", strChannel)
                Else
                    IRC.SendMessage("You are not authorized to use this command.", strChannel)
                End If
            Case "!unauth"
                If CBool(Settings.GetConfigInfo("Auth", strUserMask, False)(1)) = True Then
                    Settings.WriteConfigInfo("Auth", strWord(1), False)
                    IRC.SendMessage("User removed.", strChannel)
                Else
                    IRC.SendMessage("You are not authorized to use this command.", strChannel)
                End If
            Case "!mode"
                If CBool(Settings.GetConfigInfo("Auth", strUserMask, False)(1)) = True Then
                    If blnTestMode = True Then
                        IRC.SendMessage("Running in test mode, changes WILL NOT be applied.", strChannel)
                    Else
                        IRC.SendMessage("Running in live mode, changes WILL be applied.", strChannel)
                    End If
                Else
                    IRC.SendMessage("You are not authorized to use this command.", strChannel)
                End If
            Case "!setmode"
                If CBool(Settings.GetConfigInfo("Auth", strUserMask, False)(1)) = True Then
                    If strWord(1).ToLower = "test" Then
                        blnTestMode = True
                    Else
                        blnTestMode = False
                    End If
                    Settings.WriteConfigInfo("General", "TestMode", blnTestMode)
                    IRC.SendMessage("Run mode set.", strChannel)
                Else
                    IRC.SendMessage("You are not authorized to use this command.", strChannel)
                End If
            Case "!help"
                If strWord.GetUpperBound(0) = 0 Then
                    IRC.SendMessage("DNS-Bot Help:", strChannel)
                    IRC.SendMessage("!exit, !resolve, !dns, !hm, !about, !die, !nick, !map, !highload, !lowload, !current, !refresh, !ignore, !unignore, !auth, !unauth, !mode, !setmode, !users", strChannel)
                Else
                    Select Case strWord(1)
                        Case "!exit"
                            IRC.SendMessage("!exit - Discconects from the server but leaves the application running", strChannel)
                        Case "!resolve", "!dns"
                            IRC.SendMessage("!resolve - Resolves a domain to a IP address. Syntax: !resolve www.google.com", strChannel)
                        Case "!hm"
                            IRC.SendMessage("!hm - Resolves a domain to a IP address", strChannel)
                        Case "!about"
                            IRC.SendMessage("!about - Displays information about DNS-Bot", strChannel)
                        Case "!die"
                            IRC.SendMessage("!die - Discconects from the server and terminates the application", strChannel)
                        Case "!nick"
                            IRC.SendMessage("!nick - Changes the bot's nickname. Syntax: !nick NewName", strChannel)
                        Case "!map"
                            IRC.SendMessage("!map - Displays the parsed map data", strChannel)
                        Case "!highload"
                            IRC.SendMessage("!highload - Displays the server with the highest load", strChannel)
                        Case "!lowload"
                            IRC.SendMessage("!lowload - Displays the server with the lowest load", strChannel)
                        Case "!current"
                            IRC.SendMessage("!current - Displays the current server", strChannel)
                        Case "!refresh"
                            IRC.SendMessage("!refresh - Reloads the MAP data", strChannel)
                        Case "!ignore"
                            IRC.SendMessage("!ignore - Adds a server to the ignore list. Syntax: !ignore irc.server.tld", strChannel)
                        Case "!unignore"
                            IRC.SendMessage("!unignore - Removes a server from the ignore list. Syntax: !unignore irc.server.tld", strChannel)
                        Case "!auth"
                            IRC.SendMessage("!auth - Adds a user to the auth list. Syntax: !auth Nick!name@domain.tld", strChannel)
                        Case "!unauth"
                            IRC.SendMessage("!unauth - Removes a user from the auth list. Syntax: !unauth Nick!name@domain.tld", strChannel)
                        Case "!mode"
                            IRC.SendMessage("!mode - Displays the current running mode", strChannel)
                        Case "!setmode"
                            IRC.SendMessage("!setmode - Sets the current running mode. Syntax: !setmode test", strChannel)
                        Case "!users"
                            IRC.SendMessage("!users - Shows the number of users on the network", strChannel)
                    End Select
                End If
            Case "!users"
                '***********************
                'Added by Excaliber
                '7/10/04
                'Simply determines total number of users and displays
                '***********************
                Dim i As Int32, sum As Int32
                For i = 0 To tServer.GetUpperBound(0)
                    sum += tServer(i).Load
                Next
                IRC.SendMessage("There are currently " & sum & " users on " & tServer.GetUpperBound(0) + 1 & " servers", strChannel)
            Case "!multibot"
                ' <Added by: Adam at: 7/11/2004-05:46:00 on machine: BALLER-STA1>
                'MultiBot suport
                If CBool(Settings.GetConfigInfo("Auth", strUserMask, False)(1)) = True Then
                    If strWord.GetUpperBound(0) = 1 Then
                        If strWord(1).ToLower = "on" Then
                            Settings.WriteConfigInfo("General", "MultiBot", True)
                            IRC.SendMessage("I'm now running in MultiBot mode", strChannel)
                        ElseIf strWord(1).ToLower = "off" Then
                            Settings.WriteConfigInfo("General", "MultiBot", False)
                            IRC.SendMessage("I'm now running in single bot mode", strChannel)
                        End If
                    Else
                        If CBool(Settings.GetConfigInfo("General", "MultiBot", False)(1)) = True Then
                            IRC.SendMessage("I'm currently running in MultiBot mode", strChannel)
                        Else
                            IRC.SendMessage("I'm currently running in single bot mode", strChannel)
                        End If
                    End If
                End If

                    Case "!map"
                    Dim strMsg As String
                    Dim i As Long
                    For i = 0 To UBound(tServer)
                        If Len(strMsg) <> 0 Then
                            strMsg = strMsg & " | "
                        End If
                        strMsg += "Server: " & tServer(i).Name & "[" & tServer(i).IP & "] Load: " & tServer(i).Load
                        If tServer(i).Ignore = True Then
                            strMsg += " (IGNORED)"
                        End If
                    Next i
                    IRC.SendMessage(strMsg, strChannel)
                    Case "!highload"
                    IRC.SendMessage("Server with the highest load is " & tServer(GetServerHighLoadAsInt()).Name & " at " & tServer(GetServerHighLoadAsInt()).Load & " users.", strChannel)
                    Case "!lowload"
                    IRC.SendMessage("Server with the lowest load is " & tServer(GetServerLowLoadAsInt()).Name & " at " & tServer(GetServerLowLoadAsInt()).Load & " users.", strChannel)
                    Case "!current"
                    IRC.SendMessage("Current server is " & GetCurrentServerAsName(), strChannel)
                    Case "!refresh"
                    If CBool(Settings.GetConfigInfo("Auth", strUserMask, False)(1)) = True Then
                        IRC.SendMessage("Reloading /MAP Data.", strChannel)
                        IRC.Send("MAP")
                    Else
                        IRC.SendMessage("You are not authorized to use this command.", strChannel)
                    End If
                    Case "!ignore"
                    If CBool(Settings.GetConfigInfo("Auth", strUserMask, False)(1)) = True Then
                        Dim i As Long
                        Dim intServer As Integer
                        For i = 0 To UBound(tServer)
                            If tServer(i).name = strWord(1) Then
                                intServer = i + 1
                            End If
                        Next
                        If intServer <> 0 Then
                            tServer(intServer - 1).Ignore = True
                            Settings.WriteConfigInfo("Ingore", tServer(intServer - 1).Name, True)
                            IRC.SendMessage("Server added to ignore: " & tServer(intServer - 1).Name, strChannel)
                        Else
                            IRC.SendMessage("Unknown server: " & strWord(1), strChannel)
                        End If
                    Else
                        IRC.SendMessage("You are not authorized to use this command.", strChannel)
                    End If
                    Case "!unignore"
                    If CBool(Settings.GetConfigInfo("Auth", strUserMask, False)(1)) = True Then
                        Dim i As Long
                        Dim intServer As Integer
                        For i = 0 To UBound(tServer)
                            If tServer(i).name = strWord(1) Then
                                intServer = i + 1
                            End If
                        Next
                        If intServer <> 0 Then
                            tServer(intServer - 1).Ignore = False
                            Settings.WriteConfigInfo("Ingore", tServer(intServer - 1).Name, False)
                            IRC.SendMessage("Server removed from ignore: " & tServer(intServer - 1).Name, strChannel)
                        Else
                            IRC.SendMessage("Unknown server: " & strWord(1), strChannel)
                        End If
                    Else
                        IRC.SendMessage("You are not authorized to use this command.", strChannel)
                    End If
                    Case "!auth"
                    If UBound(strWord) > 0 Then
                        If CBool(Settings.GetConfigInfo("Auth", strUserMask, False)(1)) = True Then
                            Settings.WriteConfigInfo("Auth", strWord(1), True)
                            IRC.SendMessage("User added.", strChannel)
                        Else
                            IRC.SendMessage("You are not authorized to use this command.", strChannel)
                        End If
                    End If

                    Case "!unauth"
                    If CBool(Settings.GetConfigInfo("Auth", strUserMask, False)(1)) = True Then
                        Settings.WriteConfigInfo("Auth", strWord(1), False)
                        IRC.SendMessage("User removed.", strChannel)
                    Else
                        IRC.SendMessage("You are not authorized to use this command.", strChannel)
                    End If
                    Case "!mode"
                    If CBool(Settings.GetConfigInfo("Auth", strUserMask, False)(1)) = True Then
                        If blnTestMode = True Then
                            IRC.SendMessage("Running in test mode, changes WILL NOT be applied.", strChannel)
                        Else
                            IRC.SendMessage("Running in live mode, changes WILL be applied.", strChannel)
                        End If
                    Else
                        IRC.SendMessage("You are not authorized to use this command.", strChannel)
                    End If
                    Case "!setmode"
                    If CBool(Settings.GetConfigInfo("Auth", strUserMask, False)(1)) = True Then
                        If strWord(1).ToLower = "test" Then
                            blnTestMode = True
                        Else
                            blnTestMode = False
                        End If
                        Settings.WriteConfigInfo("General", "TestMode", blnTestMode)
                        IRC.SendMessage("Run mode set.", strChannel)
                    Else
                        IRC.SendMessage("You are not authorized to use this command.", strChannel)
                    End If
                    Case "!help"
                    If strWord.GetUpperBound(0) = 0 Then
                        IRC.SendMessage("DNS-Bot Help:", strChannel)
                        IRC.SendMessage("!exit, !resolve, !dns, !hm, !about, !die, !nick, !map, !highload, !lowload, !current, !refresh, !ignore, !unignore, !auth, !unauth, !mode, !setmode", strChannel)
                    Else
                        Select Case strWord(1)
                            Case "!exit"
                                IRC.SendMessage("!exit - Discconects from the server but leaves the application running", strChannel)
                            Case "!resolve", "!dns"
                                IRC.SendMessage("!resolve - Resolves a domain to a IP address. Syntax: !resolve www.google.com", strChannel)
                            Case "!hm"
                                IRC.SendMessage("!hm - Resolves a domain to a IP address", strChannel)
                            Case "!about"
                                IRC.SendMessage("!about - Displays information about DNS-Bot", strChannel)
                            Case "!die"
                                IRC.SendMessage("!die - Discconects from the server and terminates the application", strChannel)
                            Case "!nick"
                                IRC.SendMessage("!nick - Changes the bot's nickname. Syntax: !nick NewName", strChannel)
                            Case "!map"
                                IRC.SendMessage("!map - Displays the parsed map data", strChannel)
                            Case "!highload"
                                IRC.SendMessage("!highload - Displays the server with the highest load", strChannel)
                            Case "!lowload"
                                IRC.SendMessage("!lowload - Displays the server with the lowest load", strChannel)
                            Case "!current"
                                IRC.SendMessage("!current - Displays the current server", strChannel)
                            Case "!refresh"
                                IRC.SendMessage("!refresh - Reloads the MAP data", strChannel)
                            Case "!ignore"
                                IRC.SendMessage("!ignore - Adds a server to the ignore list. Syntax: !ignore irc.server.tld", strChannel)
                            Case "!unignore"
                                IRC.SendMessage("!unignore - Removes a server from the ignore list. Syntax: !unignore irc.server.tld", strChannel)
                            Case "!auth"
                                IRC.SendMessage("!auth - Adds a user to the auth list. Syntax: !auth Nick!name@domain.tld", strChannel)
                            Case "!unauth"
                                IRC.SendMessage("!unauth - Removes a user from the auth list. Syntax: !unauth Nick!name@domain.tld", strChannel)
                            Case "!mode"
                                IRC.SendMessage("!mode - Displays the current running mode", strChannel)
                            Case "!setmode"
                                IRC.SendMessage("!setmode - Sets the current running mode. Syntax: !setmode test", strChannel)
                        End Select
                    End If


                    '***********************
                    'Added by Zach - 7/10/04
                    'Simply determines total number of users and displays
                    '***********************
                    Case "!users"
                    Dim i As Int32, sum As Int32
                    For i = 0 To UBound(tServer)
                        sum += tServer(i).Load
                    Next
                    IRC.SendMessage("There are currently " & sum & " global users on OpenIRCNet", strChannel)


                    '***********************
                    'Added by Zach - 7/11/04
                    'Checks the channel for a lower priority
                    '***********************
                    Case "!priority"
                    IRC.Send("NAMES #" & strChannel)
                    Case "!rejoin"
                    If CBool(Settings.GetConfigInfo("Auth", strUserMask, False)(1)) = True Then
                        IRC.ReJoin()
                    Else
                        IRC.SendMessage("You are not authorized to use this command.", strChannel)
                    End If
                    Case "!partjoin"
                    If CBool(Settings.GetConfigInfo("Auth", strUserMask, False)(1)) = True Then
                        If UBound(strWord) = 0 Then
                            IRC.SendMessage("Not enough parameters.", strChannel)
                        ElseIf UBound(strWord) = 1 Then
                            IRC.PartJoin(strWord(1))
                        ElseIf UBound(strWord) = 2 Then
                            IRC.PartJoin(strWord(1), strWord(2))
                        End If

                    Else
                        IRC.SendMessage("You are not authorized to use this command.", strChannel)
                    End If
                End Select

    End Sub

    Private Sub IRC_ChannelJoin(ByVal UserName As String, ByVal strChannel As String, ByVal strUserMask As String) Handles IRC.ChannelJoin
        IRC.SendMessage("Hello " & UserName, strChannel)
        IRC.Send("NAMES #dns-bot")
    End Sub




    Private Sub ProcMap(ByVal Data As String)
        Static blnMapStarted As Boolean

        If Data <> "MAP: End of /MAP" Then
            If blnMapStarted = False Then
                'Clear array
                blnMapStarted = True
                ReDim tServer(0)
            Else
                ReDim Preserve tServer(tServer.GetUpperBound(0) + 1)
            End If
            Data = Data.Substring(InStr(Data, "MAP:") + 4)
            Data = Data.Replace("(", "")
            Data = Data.Replace(")", "")
            Data = Data.Replace("|-", "")
            Data = Data.Replace("`-", "")
            Data = Data.Replace("`", "")
            Data = Data.Replace("|", "")
            Do Until InStr(Data, "  ") = 0
                Data = Data.Replace("  ", " ")
            Loop
            Dim sData() As String
            Data = Trim(Data)
            sData = Split(Data, " ")
            tServer(UBound(tServer)).Name = sData(0)
            tServer(UBound(tServer)).Load = sData(1)
            tServer(UBound(tServer)).IP = System.Net.Dns.Resolve(sData(0)).AddressList(0).ToString
            tServer(UBound(tServer)).Ignore = Settings.GetConfigInfo("Ingore", sData(0), False)(1)
        Else
            blnMapStarted = False
        End If
    End Sub

    Private Sub frmControl_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        blnTestMode = Settings.GetConfigInfo("General", "TestMode", False)(1)
        IRC.Nickname = Settings.GetConfigInfo("General", "NickName", "DNS-Dev-ex-2")(1)
        IRC.Server = Settings.GetConfigInfo("Network", "Server", "polyfractal.ath.cx")(1)
        IRC.Port = Settings.GetConfigInfo("Network", "Port", "6667")(1)
        IRC.RealName = Settings.GetConfigInfo("General", "Name", "DNS Bot")(1)
        IRC.Version = "DNS-Bot v" & Application.ProductVersion

        m_intPriority = Settings.GetConfigInfo("General", "Priority", "2")(1)
        m_strPriorityPass = Settings.GetConfigInfo("general", "PriorityPass", "pass")(1)

        IRC.Connect()

    End Sub

    Private Function AppPath() As String
        Return System.AppDomain.CurrentDomain.BaseDirectory()
    End Function

    Private Sub SaveSettings()
        Settings.WriteConfigInfo("General", "NickName", IRC.Nickname)
        Settings.WriteConfigInfo("Network", "Server", IRC.Server)
        Settings.WriteConfigInfo("Network", "Port", IRC.Port)
        Settings.WriteConfigInfo("General", "Name", IRC.RealName)
    End Sub

    Private Sub tmrRefresh_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrRefresh.Tick
        IRC.Send("MAP")
        Application.DoEvents()
        If GetCurrentServerAsInt() <> -1 Or tServer(GetCurrentServerAsInt()).Ignore = False Then
            If GetCurrentServerAsInt() <> GetServerLowLoadAsInt() Then
                If tServer(GetCurrentServerAsInt).Load > tServer(GetServerLowLoadAsInt).Load Then
                    SetCurrentServerByInt(GetServerLowLoadAsInt())
                End If
            End If
        Else
            SetCurrentServerByInt(GetServerLowLoadAsInt())
        End If
    End Sub

    Private Function GetCurrentServerAsName() As String
        Dim strCurrentIP As String = System.Net.Dns.Resolve("openircnet.ath.cx").AddressList(0).ToString
        Dim i As Long
        Dim intCurrent As Integer
        For i = 0 To tServer.GetUpperBound(0)
            If tServer(i).IP = strCurrentIP Then
                intCurrent = i + 1
            End If
        Next
        If intCurrent <> 0 Then
            Return tServer(intCurrent - 1).Name
        Else
            Return strCurrentIP
        End If
    End Function

    Private Function GetCurrentServerAsInt() As Integer
        Dim strCurrentIP As String = System.Net.Dns.Resolve("openircnet.ath.cx").AddressList(0).ToString
        Dim i As Long
        Dim intCurrent As Integer
        For i = 0 To tServer.GetUpperBound(0)
            If tServer(i).IP = strCurrentIP Then
                intCurrent = i + 1
            End If
        Next
        If intCurrent <> 0 Then
            Return intCurrent - 1
        Else
            Return -1
        End If
    End Function

    Private Function GetServerHighLoadAsInt() As Integer
        Dim i As Long
        Dim intHigh As Integer
        For i = 1 To tServer.GetUpperBound(0)
            If tServer(i).Load > tServer(intHigh).Load And tServer(i).Ignore = False Then
                intHigh = i
            End If
        Next
        Return intHigh
    End Function

    Private Function GetServerLowLoadAsInt() As Integer
        Dim i As Long
        Dim intLow As Integer
        For i = 1 To tServer.GetUpperBound(0)
            If tServer(i).Load < tServer(intLow).Load And tServer(i).Ignore = False Then
                intLow = i
            End If
        Next
        Return intLow
    End Function

    Private Sub SetCurrentServerByInt(ByVal intServer As Integer)
        If intServer <> -1 Then
            If blnTestMode = False Then
                'Change sever
            Else
                'Test mode, just announce what we should do.
                IRC.SendMessage("Changing server to: " & tServer(intServer).Name, "#DNS-Bot")
            End If
        End If
    End Sub




    Private Sub IRC_DataArrival_StrArray(ByVal Data() As String) Handles IRC.DataArrival_StrArray
        Dim x As Int32
        Dim l_blnLowestPri As Boolean = True 'priority value defaults to us being the lowest
        Dim l_strLowestName As String = IRC.Nickname

        For x = 5 To UBound(Data) - 1
            'remove any nasty ":"s that might mess things up
            Data(x) = Replace(Data(x), ":", "")

            'make sure the name matches
            If Mid(Data(x), 1, 11) = Mid(IRC.Nickname, 1, 11) Then
                'check to see if there is a priority lower than me
                If Mid$(Data(x), 12, 1) < m_intPriority Then 'mid is always one more than above mid value

                    'there is someone lower than us, so we dont have to worry about it
                    l_blnLowestPri = False

                    'set the name for output
                    l_strLowestName = Data(x)
                    Exit For
                End If
            End If
        Next

        If l_blnLowestPri = True Then
            'we are now priority, set variable as such
            m_blnIsPriority = True
            IRC.SendMessage("I am currently holding Priority", Data(4))
        Else
            m_blnIsPriority = False
            IRC.SendMessage("AUTH", l_strLowestName)
            tmrChallenge.Interval = 5000
            tmrChallenge.Enabled = True

            'IRC.SendMessage(l_strLowestName & " is currently the lowest priority.", Data(4))
        End If
    End Sub

    Private Sub IRC_NickChange(ByVal UserName As String, ByVal strChannel As String, ByVal strUserMask As String) Handles IRC.NickChange
        IRC.Send("NAMES #" & IRC.Channel)
    End Sub

    Private Sub txtReceived_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtReceived.TextChanged
        txtReceived.SelectionStart = Len(txtReceived.Text)
    End Sub

    Private Sub tmrChallenge_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrChallenge.Tick
        IRC.SendMessage("I am retaining Priority", IRC.Channel)
        tmrChallenge.Stop()
        tmrChallenge.Enabled = False
    End Sub
End Class


Public Class frmControl
    Inherits System.Windows.Forms.Form

    Dim WithEvents IRC As New IRC_Lib.cIRC

    Dim tServer() As Server

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
    Friend WithEvents btnConnect As System.Windows.Forms.Button
    Friend WithEvents txtReceived As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnConnect = New System.Windows.Forms.Button
        Me.txtReceived = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'btnConnect
        '
        Me.btnConnect.Location = New System.Drawing.Point(440, 288)
        Me.btnConnect.Name = "btnConnect"
        Me.btnConnect.Size = New System.Drawing.Size(136, 48)
        Me.btnConnect.TabIndex = 0
        Me.btnConnect.Text = "Connect"
        '
        'txtReceived
        '
        Me.txtReceived.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtReceived.Location = New System.Drawing.Point(0, 0)
        Me.txtReceived.Multiline = True
        Me.txtReceived.Name = "txtReceived"
        Me.txtReceived.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtReceived.Size = New System.Drawing.Size(584, 272)
        Me.txtReceived.TabIndex = 1
        Me.txtReceived.Text = ""
        Me.txtReceived.WordWrap = False
        '
        'frmControl
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(584, 339)
        Me.Controls.Add(Me.txtReceived)
        Me.Controls.Add(Me.btnConnect)
        Me.Name = "frmControl"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnConnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConnect.Click
        'Dim IRC As New IRC_Lib.cIRC
        IRC.Nickname = "DNS-Dev-Bot"
        IRC.Server = "baller-srv1"
        IRC.Port = 6667
        IRC.RealName = "DNS-DevBot"
        IRC.Version = "DNS-Bot v" & Application.ProductVersion
        IRC.Connect()
    End Sub

    Private Sub IRC_DataArrival(ByVal Data As String) Handles IRC.DataArrival
        txtReceived.Text += Data & ControlChars.CrLf
        If Data.Substring(0, 4) = "MAP:" Then
            ProcMap(Data)
        End If
    End Sub

    Private Sub IRC_ConnectComplete() Handles IRC.ConnectComplete
        IRC.Send("MAP")
        IRC.Join("#DNS-Bot", "test")
        Application.DoEvents()
        IRC.SendMessage("DNS-Bot (v" & Application.ProductVersion & ") Online.", "#DNS-Bot")
    End Sub

    Private Sub IRC_ChannelMessage(ByVal Data As String, ByVal strChannel As String, ByVal strUserMask As String) Handles IRC.ChannelMessage
        Dim strWord() As String = Data.Split(" ")
        Select Case strChannel.ToLower
            Case "#dns-bot"
                Select Case strWord(0).ToLower
                    Case "!exit"
                        IRC.Quit("Leaving(Channel Exit(" & strUserMask & "))")
                    Case "!resolve"
                        IRC.SendMessage(strWord(1) & " is " & System.Net.Dns.Resolve(strWord(1)).AddressList(0).ToString, strChannel)
                    Case "!hm"
                        IRC.SendMessage("Your hostmask is " & strUserMask, strChannel)
                    Case "!die"
                        IRC.Quit("Leaving(Channel Die(" & strUserMask & "))")
                        Application.Exit()
                    Case "!about"
                        IRC.SendMessage("I'm DNS-Bot, created by Adam Caudill with the help of several great people. For more information go to http://sourceforge.net/projects/dnsbot/ - DNS-Bot (v" & Application.ProductVersion & ")", strChannel)
                    Case "!nick"
                        IRC.SendMessage("Changing name to: " & strWord(1), strChannel)
                        IRC.ChangeNick(strWord(1))
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
                        Dim i As Long
                        Dim intHigh As Integer
                        For i = 1 To UBound(tServer)
                            If tServer(i).Load > tServer(intHigh).Load Then
                                intHigh = i
                            End If
                        Next
                        IRC.SendMessage("Server with the highest load is " & tServer(intHigh).Name & " at " & tServer(intHigh).Load & " users.", strChannel)
                    Case "!lowload"
                        Dim i As Long
                        Dim intLow As Integer
                        For i = 1 To UBound(tServer)
                            If tServer(i).Load < tServer(intLow).Load Then
                                intLow = i
                            End If
                        Next
                        IRC.SendMessage("Server with the lowest load is " & tServer(intLow).Name & " at " & tServer(intLow).Load & " users.", strChannel)
                    Case "!current"
                        Dim strCurrentIP As String = System.Net.Dns.Resolve("openircnet.ath.cx").AddressList(0).ToString
                        Dim i As Long
                        Dim intCurrent As Integer
                        For i = 0 To UBound(tServer)
                            If tServer(i).IP = strCurrentIP Then
                                intCurrent = i + 1
                            End If
                        Next
                        If intCurrent <> 0 Then
                            IRC.SendMessage("Current server is " & tServer(intCurrent - 1).Name, strChannel)
                        Else
                            IRC.SendMessage("Current server is " & strCurrentIP, strChannel)
                        End If
                    Case "!refresh"
                        IRC.SendMessage("Reloading /MAP Data.", strChannel)
                        IRC.Send("MAP")
                End Select
        End Select
    End Sub

    Private Sub IRC_ChannelJoin(ByVal UserName As String, ByVal strChannel As String, ByVal strUserMask As String) Handles IRC.ChannelJoin
        IRC.SendMessage("Hello " & UserName, strChannel)
    End Sub

    Private Sub ProcMap(ByVal Data As String)
        Static blnMapStarted As Boolean

        If Data <> "MAP: End of /MAP" Then
            If blnMapStarted = False Then
                'Clear array
                blnMapStarted = True
                ReDim tServer(0)
            Else
                ReDim Preserve tServer(UBound(tServer) + 1)
            End If
            Data = Mid(Data, InStr(Data, "MAP:") + 5)
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
            tServer(UBound(tServer)).Ignore = False
        Else
            blnMapStarted = False
        End If
    End Sub
End Class

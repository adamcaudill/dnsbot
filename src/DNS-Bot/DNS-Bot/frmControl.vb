Public Class frmControl
    Inherits System.Windows.Forms.Form

    Dim WithEvents IRC As New IRC_Lib.cIRC

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
        IRC.RealName = "DNS Dev Bot"
        IRC.Version = "DNS-Bot v" & Application.ProductVersion
        IRC.Connect()
    End Sub

    Private Sub IRC_DataArrival(ByVal Data As String) Handles IRC.DataArrival
        txtReceived.Text += Data & ControlChars.CrLf
    End Sub

    Private Sub IRC_ConnectComplete() Handles IRC.ConnectComplete
        IRC.Join("#DNS-Bot")
        Application.DoEvents()
        IRC.SendMessage("DNS-Bot (" & Application.ProductVersion & ") Online.", "#DNS-Bot")
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
                End Select
        End Select
    End Sub
End Class

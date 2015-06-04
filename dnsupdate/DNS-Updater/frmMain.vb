Public Class frmMain
    Inherits System.Windows.Forms.Form

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
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(288, 69)
        Me.Name = "frmMain"
        Me.Text = "Updater"

    End Sub

#End Region
    Dim Settings As New clsXMLCfgFile(System.AppDomain.CurrentDomain.BaseDirectory & "\Settings.xml")


    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim wc As New System.Net.WebClient
        Dim strVersionData() As String

        Dim strBuffer As String = System.Text.Encoding.ASCII.GetString(wc.DownloadData(Settings.GetConfigInfo("Network", "BinaryLocation", "")(1) & "\version.txt"))
        Dim strLocalBuffer As String

        strVersionData = Split(strBuffer, "#")

        Dim oStream As New System.IO.FileStream(System.AppDomain.CurrentDomain.BaseDirectory & "\version.txt", IO.FileMode.Open)
        Dim oReader As New System.IO.StreamReader(oStream)
        Dim oWriter As System.IO.StreamWriter

        strLocalBuffer = oReader.ReadToEnd

        oReader.Close()
        oStream.Close()

        If strVersionData(0).ToString <> strLocalBuffer Then

            ' Declare a Process class.
            Dim proc As Process
            ' Loop through the Processes and write the name of the Process to the output window.
            For Each proc In Process.GetProcesses
                If proc.ProcessName = "DNS-Bot" Then
                    proc.Kill()
                    proc.Dispose()

                End If
            Next

            System.Threading.Thread.CurrentThread.Sleep(3000)

            Dim i As Int32
            For i = 1 To UBound(strVersionData)
                Dim strRemoteLocation As String
                Dim strLocalLocation As String
                Dim tempArr() As String


                tempArr = Split(strVersionData(i), "*")
                
                strRemoteLocation = tempArr(0)
                strLocalLocation = tempArr(1)

                'Kill kill kill
                Debug.WriteLine(System.AppDomain.CurrentDomain.BaseDirectory & strLocalLocation)
                Try
                    Kill(System.AppDomain.CurrentDomain.BaseDirectory & strLocalLocation)
                Catch ex As Exception
                End Try

                'download the file
                Debug.WriteLine(Settings.GetConfigInfo("Network", "BinaryLocation", "")(1) & strRemoteLocation)
                Debug.WriteLine(System.AppDomain.CurrentDomain.BaseDirectory & strLocalLocation)

                wc.DownloadFile(Settings.GetConfigInfo("Network", "BinaryLocation", "")(1) & strRemoteLocation, System.AppDomain.CurrentDomain.BaseDirectory & strLocalLocation)


            Next

            Process.Start(System.AppDomain.CurrentDomain.BaseDirectory & "\DNS-Bot.exe")

        End If
        Kill(System.AppDomain.CurrentDomain.BaseDirectory & "\version.txt")
        oStream = New System.IO.FileStream(System.AppDomain.CurrentDomain.BaseDirectory & "\version.txt", IO.FileMode.CreateNew)
        oWriter = New System.IO.StreamWriter(oStream)

        owriter.Write(strVersionData(0).ToString)
        owriter.Flush()

        owriter.Close()
        oStream.Close()


        Application.Exit()


    End Sub
End Class

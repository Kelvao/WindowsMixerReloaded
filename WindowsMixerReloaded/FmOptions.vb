Imports System.Globalization
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading
Imports IWshRuntimeLibrary
Imports File = System.IO.File

Public Class FmOptions
    Private Const startupKeyPath As String = "Software\Microsoft\Windows\CurrentVersion\Run"
    Private Const mixerExecutableName As String = "SndVol.exe"
    Private Const shortcutExtension As String = ".lnk"

    Private trayIcon As NotifyIcon
    Private contextMenu As New ContextMenuStrip
    Private cbStartup As New ToolStripMenuItem(My.Resources.cb_startup)

    Public Sub New()
        InitializeComponent()
        SetupCulture()
        SetupIconTrayButtons()
        SetupTrayIcon()
    End Sub

    Private Sub SetupCulture()
        Dim systemCulture As CultureInfo = CultureInfo.CurrentCulture
        Thread.CurrentThread.CurrentCulture = systemCulture
        Thread.CurrentThread.CurrentUICulture = systemCulture
    End Sub

    Private Sub SetupIconTrayButtons()
        cbStartup.CheckOnClick = True
        contextMenu.Items.Add(My.Resources.cms_open)
        contextMenu.Items.Add(cbStartup)
        contextMenu.Items.Add(My.Resources.cms_exit)
    End Sub

    Private Sub SetupTrayIcon()
        trayIcon = New NotifyIcon() With {
          .Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath),
          .Text = Application.ProductName,
          .ContextMenuStrip = contextMenu,
          .Visible = True
        }

        AddHandler contextMenu.ItemClicked, AddressOf TrayIconRightClick
        AddHandler trayIcon.Click, AddressOf TrayIconLeftClick
    End Sub

    Private Sub TrayIconRightClick(sender As Object, e As ToolStripItemClickedEventArgs)
        cbStartup.Checked = ShortcutExist()

        Select Case e.ClickedItem.Text
            Case My.Resources.cms_open
                StartWindowsMixerAtBottomRight()
            Case My.Resources.cb_startup
                If cbStartup.Checked Then
                    DeleteShortcut()
                Else
                    CreateShortcut()
                End If
            Case My.Resources.cms_exit
                trayIcon.Visible = False
                Application.Exit()
        End Select
    End Sub

    Private Sub TrayIconLeftClick(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then
            StartWindowsMixerAtBottomRight()
        End If
    End Sub

    Private Sub StartWindowsMixerAtBottomRight()
        Try
            Dim windowsMixer = New Process()
            windowsMixer.StartInfo.FileName = mixerExecutableName
            windowsMixer.Start()

            While True
                Dim hWnd As IntPtr = windowsMixer.MainWindowHandle

                Dim screenWidth As Integer = Screen.PrimaryScreen.WorkingArea.Width
                Dim screenHeight As Integer = Screen.PrimaryScreen.WorkingArea.Height
                Dim windowRect As New RECT()
                GetWindowRect(hWnd, windowRect)

                Dim windowWidth As Integer = windowRect.Right - windowRect.Left
                Dim windowHeight As Integer = windowRect.Bottom - windowRect.Top

                Dim newX As Integer = screenWidth - windowWidth
                Dim newY As Integer = screenHeight - windowHeight

                If SetWindowPos(hWnd, IntPtr.Zero, newX, newY, windowWidth, windowHeight, &H1) Then
                    Exit While
                End If
            End While

        Catch ex As Exception
            MessageBox.Show(My.Resources.msgbox_error_message & " " & ex.Message, My.Resources.msgbox_error_title, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub FmOptions_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Hide()
        Me.ShowInTaskbar = False
        cbStartup.Checked = ShortcutExist()
    End Sub
    Private Sub FmOptions_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        e.Cancel = True
        Me.Hide()
    End Sub

    Private Function ShortcutExist() As Boolean
        Dim startupFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.Startup)
        Dim shortcutPath As String = Path.Combine(startupFolder, Application.ProductName & shortcutExtension)

        Return File.Exists(shortcutPath)
    End Function

    Private Sub CreateShortcut()
        Dim startupFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.Startup)
        Dim shortcutPath As String = Path.Combine(startupFolder, Application.ProductName & shortcutExtension)

        Dim shell As New WshShell()

        Dim shortcut As IWshShortcut = CType(shell.CreateShortcut(shortcutPath), IWshShortcut)
        shortcut.TargetPath = Application.ExecutablePath
        shortcut.WorkingDirectory = Path.GetDirectoryName(Application.ExecutablePath)
        shortcut.Save()
    End Sub

    Private Sub DeleteShortcut()
        Dim startupFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.Startup)
        Dim shortcutPath As String = Path.Combine(startupFolder, Application.ProductName & shortcutExtension)

        If File.Exists(shortcutPath) Then
            File.Delete(shortcutPath)
        End If
    End Sub

    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As UInteger) As Boolean

    Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As IntPtr, ByRef lpRect As RECT) As Boolean

    <StructLayout(LayoutKind.Sequential)>
    Public Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure
End Class

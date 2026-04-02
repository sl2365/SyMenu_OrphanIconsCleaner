Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Windows.Forms
Imports System.Xml
Imports Microsoft.VisualBasic

Public Class MainForm
    Inherits Form

    ' === Controls ===
    Private WithEvents btnBrowse As Button
    Private lblPath As Label
    Private WithEvents chkDelete As CheckBox
    Private lblExcLabel As Label
    Private txtExclude As TextBox
    Private WithEvents btnGo As Button
    Private WithEvents btnEdit As Button
    Private WithEvents btnSave As Button
    Private splitter As SplitContainer
    Private lblLogLabel As Label
    Private txtLog As TextBox
    Private WithEvents chkNewLog As CheckBox
    Private WithEvents chkVerbose As CheckBox
    Private toolTip As ToolTip

    ' === Settings ===
    Private s_SyMenuPath As String = ""
    Private b_DeleteAction As Boolean = False
    Private b_NewLogFile As Boolean = True
    Private b_VerboseLog As Boolean = False
    Private s_TextLog As String = ""

    ' === Paths ===
    Private ReadOnly s_AppDir As String = AppDomain.CurrentDomain.BaseDirectory
    Private ReadOnly s_LogFile As String
    Private ReadOnly s_ExcludeFile As String
    Private ReadOnly s_SettingsFile As String

    ' === Exclusion list ===
    Private a_ExcludeList As New Dictionary(Of String, Boolean)(StringComparer.OrdinalIgnoreCase)

    Public Sub New()
        s_LogFile = Path.Combine(s_AppDir, "Log.ini")
        s_ExcludeFile = Path.Combine(s_AppDir, "Exclusions.ini")
        s_SettingsFile = Path.Combine(s_AppDir, "Settings.ini")

        Me.Size = New Size(800, 500)
        Me.MinimumSize = New Size(650, 350)
        LoadSettings()
        LoadExcludeList()
        InitializeUI()
    End Sub

    Private Sub InitializeUI()
        ' === Form setup ===
        Me.Text = "SyMenu Orphan Icons Cleaner"
        Me.Padding = New Padding(10, 0, 10, 0)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath)
        Me.Font = New Font("Segoe UI", 9)

        toolTip = New ToolTip()
        toolTip.AutoPopDelay = 5000
        toolTip.InitialDelay = 500
        toolTip.ReshowDelay = 200

        ' === Top panel ===
        Dim topPanel As New Panel()
        topPanel.Dock = DockStyle.Top
        topPanel.Height = 60

        btnBrowse = New Button()
        btnBrowse.Text = "Browse..."
        btnBrowse.Location = New Point(5, 8)
        btnBrowse.Size = New Size(75, 25)
        toolTip.SetToolTip(btnBrowse, "Select SyMenu installation folder")
        topPanel.Controls.Add(btnBrowse)

        lblPath = New Label()
        lblPath.Text = If(s_SyMenuPath <> "", "  " & s_SyMenuPath, " (Click 'Browse' and locate the SyMenu installation folder)")
        lblPath.Location = New Point(85, 12)
        lblPath.AutoSize = False
        lblPath.ForeColor = If(s_SyMenuPath <> "", Color.Green, Color.Red)
        lblPath.Font = New Font(lblPath.Font, FontStyle.Bold)
        lblPath.BorderStyle = BorderStyle.FixedSingle
        topPanel.Controls.Add(lblPath)
        lblPath.Size = New Size(topPanel.ClientSize.Width - lblPath.Left, 18)
        lblPath.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

        chkDelete = New CheckBox()
        chkDelete.Text = "Delete Icons"
        chkDelete.Location = New Point(12, 37)
        chkDelete.AutoSize = True
        chkDelete.Checked = b_DeleteAction
        toolTip.SetToolTip(chkDelete, "Checked - Permanently deletes icons." & vbCrLf & "Unchecked - move to _Trash\_OrphanIcons")
        topPanel.Controls.Add(chkDelete)

        Me.Controls.Add(topPanel)

        ' === Main split container ===
        splitter = New SplitContainer()
        splitter.Dock = DockStyle.Fill
        splitter.SplitterWidth = 6
        splitter.BorderStyle = BorderStyle.None

        ' === Left panel (exclusions) ===
        lblExcLabel = New Label()
        lblExcLabel.Text = "Excluded icons (one filename.ico per line):"
        lblExcLabel.ForeColor = Color.Purple
        lblExcLabel.Dock = DockStyle.Top
        lblExcLabel.Height = 22
        lblExcLabel.Padding = New Padding(0, 0, 0, 2)
        splitter.Panel1.Controls.Add(lblExcLabel)

        txtExclude = New TextBox()
        txtExclude.Multiline = True
        txtExclude.ScrollBars = ScrollBars.Vertical
        txtExclude.WordWrap = True
        txtExclude.Dock = DockStyle.Fill
        txtExclude.ReadOnly = True
        txtExclude.Text = GetExcludeDisplay()
        toolTip.SetToolTip(txtExclude, "One filename.ico per line." & vbCrLf & "Comments are ignored")
        splitter.Panel1.Controls.Add(txtExclude)

        Dim leftBtnPanel As New Panel()
        leftBtnPanel.Dock = DockStyle.Bottom
        leftBtnPanel.Height = 48

        btnGo = New Button()
        btnGo.Text = "Go!"
        btnGo.Location = New Point(5, 10)
        btnGo.Size = New Size(75, 28)
        toolTip.SetToolTip(btnGo, "Scan for orphan icons")
        leftBtnPanel.Controls.Add(btnGo)

        btnEdit = New Button()
        btnEdit.Text = "Edit"
        btnEdit.Location = New Point(85, 10)
        btnEdit.Size = New Size(75, 28)
        toolTip.SetToolTip(btnEdit, "Edit the exclusion list")
        leftBtnPanel.Controls.Add(btnEdit)

        btnSave = New Button()
        btnSave.Text = "Save"
        btnSave.Location = New Point(165, 10)
        btnSave.Size = New Size(75, 28)
        btnSave.Enabled = False
        toolTip.SetToolTip(btnSave, "Save changes to the exclusion list")
        leftBtnPanel.Controls.Add(btnSave)

        splitter.Panel1.Controls.Add(leftBtnPanel)

        ' === Right panel (log) ===
        lblLogLabel = New Label()
        lblLogLabel.Text = "Log Window:"
        lblLogLabel.ForeColor = Color.Blue
        lblLogLabel.Dock = DockStyle.Top
        lblLogLabel.Height = 22
        lblLogLabel.Padding = New Padding(0, 0, 0, 2)
        splitter.Panel2.Controls.Add(lblLogLabel)

        txtLog = New TextBox()
        txtLog.Multiline = True
        txtLog.ScrollBars = ScrollBars.Vertical
        txtLog.WordWrap = True
        txtLog.Dock = DockStyle.Fill
        txtLog.ReadOnly = True
        splitter.Panel2.Controls.Add(txtLog)

        Dim rightBtnPanel As New Panel()
        rightBtnPanel.Dock = DockStyle.Bottom
        rightBtnPanel.Height = 48

        chkNewLog = New CheckBox()
        chkNewLog.Text = "Delete Log File?"
        chkNewLog.Location = New Point(5, 5)
        chkNewLog.AutoSize = True
        chkNewLog.Checked = b_NewLogFile
        toolTip.SetToolTip(chkNewLog, "New log file each run")
        rightBtnPanel.Controls.Add(chkNewLog)

        chkVerbose = New CheckBox()
        chkVerbose.Text = "Log All Events?"
        chkVerbose.Location = New Point(5, 25)
        chkVerbose.AutoSize = True
        chkVerbose.Checked = b_VerboseLog
        toolTip.SetToolTip(chkVerbose, "Log all icons, not just orphans")
        rightBtnPanel.Controls.Add(chkVerbose)

        splitter.Panel2.Controls.Add(rightBtnPanel)

        ' Add splitter to form (add AFTER topPanel so docking works correctly)
'         Left panel - add Bottom, then Fill, then Top last
'         splitter.Panel1.Controls.Add(leftBtnPanel)
'         splitter.Panel1.Controls.Add(txtExclude)
'         splitter.Panel1.Controls.Add(lblExcLabel)

'         Right panel - add Bottom, then Fill, then Top last
'         splitter.Panel2.Controls.Add(rightBtnPanel)
'         splitter.Panel2.Controls.Add(txtLog)
'         splitter.Panel2.Controls.Add(lblLogLabel)

        ' Add splitter to form (add AFTER topPanel so docking works correctly)
        Me.Controls.Add(splitter)

        ' Ensure correct Z-order: splitter fills space below topPanel
        splitter.BringToFront()
    End Sub

    ' ==========================================================================
    ' SETTINGS PERSISTENCE (simple INI-style text file)
    ' ==========================================================================
    Private Sub LoadSettings()
        If Not File.Exists(s_SettingsFile) Then Return

        Try
            For Each line As String In File.ReadAllLines(s_SettingsFile, Encoding.UTF8)
                Dim parts() As String = line.Split(New Char() {"="c}, 2)
                If parts.Length <> 2 Then Continue For
                Dim key As String = parts(0).Trim()
                Dim value As String = parts(1).Trim()
                Select Case key
                    Case "SyMenuPath" : s_SyMenuPath = value
                    Case "DeleteAction" : b_DeleteAction = (value = "1")
                    Case "NewLogFile" : b_NewLogFile = (value = "1")
                    Case "VerboseLog" : b_VerboseLog = (value = "1")
                    Case "WindowWidth"
                        Dim w As Integer
                        If Integer.TryParse(value, w) AndAlso w > 0 Then Me.Width = w
                    Case "WindowHeight"
                        Dim h As Integer
                        If Integer.TryParse(value, h) AndAlso h > 0 Then Me.Height = h
                    Case "SplitterDistance"
                        ' Applied after UI init
                End Select
            Next
        Catch
            ' Ignore corrupt settings
        End Try
    End Sub

    Private Sub SaveSettings()
        Try
            Dim sb As New StringBuilder()
            sb.AppendLine("SyMenuPath=" & s_SyMenuPath)
            sb.AppendLine("DeleteAction=" & If(b_DeleteAction, "1", "0"))
            sb.AppendLine("NewLogFile=" & If(b_NewLogFile, "1", "0"))
            sb.AppendLine("VerboseLog=" & If(b_VerboseLog, "1", "0"))
            sb.AppendLine("WindowWidth=" & Me.Width.ToString())
            sb.AppendLine("WindowHeight=" & Me.Height.ToString())
            If splitter IsNot Nothing Then
                sb.AppendLine("SplitterDistance=" & splitter.SplitterDistance.ToString())
            End If
            File.WriteAllText(s_SettingsFile, sb.ToString(), Encoding.UTF8)
        Catch
            ' Ignore save errors
        End Try
    End Sub

    ' ==========================================================================
    ' EXCLUSION LIST
    ' ==========================================================================
    Private Sub LoadExcludeList()
        a_ExcludeList.Clear()

        If Not File.Exists(s_ExcludeFile) Then
            Dim defaultContent As String =
                "; SyMenu Orphan Icons - Exclusion List" & vbCrLf &
                "; ----------------------------------------" & vbCrLf &
                "; Add one icon filename per line to prevent it from being moved or deleted." & vbCrLf &
                "; Commented lines starting with ; are ignored." & vbCrLf &
                "; Blank lines are also ignored." & vbCrLf &
                "; Comments can be amended and saved." & vbCrLf &
                "; Examples:" & vbCrLf &
                ";  MyCustomIcon.ico" & vbCrLf &
                ";  AnotherIcon.ico" & vbCrLf
            File.WriteAllText(s_ExcludeFile, defaultContent)
            Return
        End If

        For Each line As String In File.ReadAllLines(s_ExcludeFile)
            Dim trimmed As String = line.Trim()
            If trimmed = "" OrElse trimmed.StartsWith(";") Then Continue For
            a_ExcludeList(trimmed) = True
        Next
    End Sub

    Private Function GetExcludeDisplay() As String
        If Not File.Exists(s_ExcludeFile) Then
            Return "(No Exclusions.txt found. Click 'Edit' to add icon filenames.)"
        End If
        Return File.ReadAllText(s_ExcludeFile)
    End Function

    Private Function IsExcluded(filename As String) As Boolean
        Return a_ExcludeList.ContainsKey(filename)
    End Function

    ' ==========================================================================
    ' GUI EVENT HANDLERS
    ' ==========================================================================
    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        Using fbd As New FolderBrowserDialog()
            fbd.Description = "Select the folder of the working SyMenu"
            fbd.ShowNewFolderButton = False
            If s_SyMenuPath <> "" AndAlso Directory.Exists(s_SyMenuPath) Then
                fbd.SelectedPath = s_SyMenuPath
            End If
            If fbd.ShowDialog() = DialogResult.OK Then
                s_SyMenuPath = fbd.SelectedPath.TrimEnd("\"c)
                lblPath.Text = "  " & s_SyMenuPath
                lblPath.ForeColor = Color.Green
'                 SaveSettings()
            End If
        End Using
    End Sub

    Protected Overrides Sub OnLoad(e As EventArgs)
        MyBase.OnLoad(e)
        Try
            splitter.Panel1MinSize = 260
            splitter.Panel2MinSize = 200
            splitter.SplitterDistance = splitter.Width \ 2
        Catch
        End Try
    End Sub

    Private Sub chkDelete_CheckedChanged(sender As Object, e As EventArgs) Handles chkDelete.CheckedChanged
        b_DeleteAction = chkDelete.Checked
'         SaveSettings()
    End Sub

    Private Sub chkNewLog_CheckedChanged(sender As Object, e As EventArgs) Handles chkNewLog.CheckedChanged
        b_NewLogFile = chkNewLog.Checked
'         SaveSettings()
    End Sub

    Private Sub chkVerbose_CheckedChanged(sender As Object, e As EventArgs) Handles chkVerbose.CheckedChanged
        b_VerboseLog = chkVerbose.Checked
'         SaveSettings()
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        txtExclude.ReadOnly = False
        btnEdit.Enabled = False
        btnSave.Enabled = True
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim newContent As String = txtExclude.Text
        File.WriteAllText(s_ExcludeFile, newContent)

        a_ExcludeList.Clear()
        For Each line As String In newContent.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.None)
            Dim trimmed As String = line.Trim()
            If trimmed = "" OrElse trimmed.StartsWith(";") Then Continue For
            a_ExcludeList(trimmed) = True
        Next

        txtExclude.ReadOnly = True
        btnEdit.Enabled = True
        btnSave.Enabled = False

        MessageBox.Show("Exclusion list saved: " & a_ExcludeList.Count & " icon(s) excluded.",
                        "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        SaveSettings()
        MyBase.OnFormClosing(e)
    End Sub

    ' ==========================================================================
    ' LOGGING
    ' ==========================================================================
    Private Sub ScriptLog(msg As String, Optional tabDeep As Integer = 0)
        Dim prefix As String = New String(vbTab, tabDeep)
        If tabDeep > 0 Then prefix &= "|"
        Dim fullMsg As String = prefix & msg & vbCrLf

        s_TextLog &= fullMsg
        txtLog.Text = s_TextLog
        txtLog.SelectionStart = txtLog.Text.Length
        txtLog.ScrollToCaret()

        Try
            File.AppendAllText(s_LogFile, fullMsg)
        Catch
            ' Ignore file write errors
        End Try
    End Sub

    Private Function TimeStamp() As String
        Return DateTime.Now.ToString("yyyy.MM.dd-HH:mm:ss")
    End Function

    ' ==========================================================================
    ' MAIN PROCESS
    ' ==========================================================================
    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        b_DeleteAction = chkDelete.Checked
        b_NewLogFile = chkNewLog.Checked
        b_VerboseLog = chkVerbose.Checked
        SaveSettings()

        If s_SyMenuPath = "" Then
            MessageBox.Show("Please select a SyMenu folder first using the 'Browse...' button.",
                            Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim s_SyMenuPathIcons As String = Path.Combine(s_SyMenuPath, "Icons")
        Dim s_SyMenuPathConfig As String = Path.Combine(s_SyMenuPath, "Config")
        Dim s_SyMenuPath_Trash As String = Path.Combine(s_SyMenuPath, "ProgramFiles", "SPSSuite", "SyMenuSuite", "_Trash")
        Dim s_OrphanIconsPath As String = Path.Combine(s_SyMenuPath_Trash, "_OrphanIcons")
        Dim s_ZipFile As String = Path.Combine(s_SyMenuPathConfig, "SyMenuItem.zip")

        ' Validate paths
        If Not Directory.Exists(s_SyMenuPathIcons) Then
            MessageBox.Show("Icons folder not found at:" & vbCrLf & s_SyMenuPathIcons,
                            "Folder Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If Not File.Exists(s_ZipFile) Then
            MessageBox.Show("SyMenuItem.zip not found at:" & vbCrLf & s_SyMenuPathConfig,
                            "Config Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' Clean log if requested
        If b_NewLogFile AndAlso File.Exists(s_LogFile) Then
            Try
                File.Delete(s_LogFile)
            Catch
            End Try
        End If

        s_TextLog = ""
        txtLog.Text = ""
        ScriptLog(TimeStamp() & " - SyMenu Orphan Icons Cleaner STARTED")

        ' ============================================================
        ' PHASE 1: Extract SyMenuItem.xml from Config\SyMenuItem.zip
        ' ============================================================
        Dim s_SyMenuItemConfig As String = ""

        Try
            If Not Directory.Exists(s_OrphanIconsPath) Then
                Directory.CreateDirectory(s_OrphanIconsPath)
            End If

            ' Extract SyMenuItem.xml directly from zip into memory
            Using archive As ZipArchive = ZipFile.OpenRead(s_ZipFile)
                Dim entry As ZipArchiveEntry = archive.GetEntry("SyMenuItem.xml")
                If entry Is Nothing Then
                    MessageBox.Show("SyMenuItem.xml not found inside SyMenuItem.zip",
                                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
                Using reader As New StreamReader(entry.Open())
                    s_SyMenuItemConfig = reader.ReadToEnd()
                End Using
            End Using

            ScriptLog(TimeStamp() & " - Config loaded from: Config\SyMenuItem.zip", 1)

        Catch ex As Exception
            MessageBox.Show("Error extracting SyMenuItem.zip:" & vbCrLf & ex.Message,
                            Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try

        ' ============================================================
        ' PHASE 2: Scan icons against config
        ' ============================================================
        Dim excludedCount As Integer = 0
        Dim orphanCount As Integer = 0
        Dim activeCount As Integer = 0

        For Each icoFile As String In Directory.GetFiles(s_SyMenuPathIcons, "*.ico")
            Dim fileName As String = Path.GetFileName(icoFile)

            If IsExcluded(fileName) Then
                excludedCount += 1
                If b_VerboseLog Then
                    ScriptLog(TimeStamp() & " - EXCLUDED (skipped): " & fileName, 2)
                End If
                Continue For
            End If

            If s_SyMenuItemConfig.IndexOf(fileName, StringComparison.OrdinalIgnoreCase) >= 0 Then
                activeCount += 1
                If b_VerboseLog Then
                    ScriptLog(TimeStamp() & " - Active Icon: " & fileName, 2)
                End If
            Else
                orphanCount += 1
                ScriptLog(TimeStamp() & " - Orphan Icon moved: " & fileName, 2)
                Try
                    Dim destFile As String = Path.Combine(s_OrphanIconsPath, fileName)
                    If File.Exists(destFile) Then File.Delete(destFile)
                    File.Move(icoFile, destFile)
                Catch ex As Exception
                    ScriptLog(TimeStamp() & " - ERROR moving " & fileName & ": " & ex.Message, 2)
                End Try
            End If
        Next

        ' ============================================================
        ' PHASE 3: Cleanup
        ' ============================================================
        If b_DeleteAction Then
            If b_VerboseLog Then
                ScriptLog(TimeStamp() & " - Deleting orphan icons folder", 1)
            End If
            Try
                If Directory.Exists(s_OrphanIconsPath) Then
                    Directory.Delete(s_OrphanIconsPath, True)
                End If
            Catch ex As Exception
                ScriptLog(TimeStamp() & " - ERROR deleting folder: " & ex.Message, 1)
            End Try
        End If

        ' Summary
        Dim summary As String = String.Format("Finished. Active: {0} | Orphans: {1} | Excluded: {2}",
                                               activeCount, orphanCount, excludedCount)
        ScriptLog(TimeStamp() & " - " & summary)
        ScriptLog(TimeStamp() & " - " & Environment.UserName & "! " & New String("-"c, 60))

        MessageBox.Show(summary & vbCrLf & Environment.UserName & "!",
                        Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ' ==========================================================================
    ' ENTRY POINT
    ' ==========================================================================
    <STAThread>
    Public Shared Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Try
            Application.Run(New MainForm())
        Catch ex As Exception
            MessageBox.Show("Startup error:" & Environment.NewLine & ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
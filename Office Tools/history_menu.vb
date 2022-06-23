Imports System.IO

Public Class history_menu
    ReadOnly logPath As String = "log/log"
    ReadOnly advLogPath As String = "log/advlog"
    ReadOnly resLogPath As String = "log/reslog"
    ReadOnly roboLogPath As String = "log/robolog"
    ReadOnly errPath As String = "log/err"
    ReadOnly advErrPath As String = "log/adverr"
    ReadOnly resErrPath As String = "log/reserr"
    Dim logCountPath As String = "log/expLog"
    Private Sub history_load(sender As Object, e As EventArgs) Handles MyBase.Load
        AllowTransparency = False
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RichTextBox1.Text = ""
        Label1.Text = "Backup"
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        RichTextBox1.Text = ""
        Label1.Text = "Archive"
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        RichTextBox1.Text = ""
        Label1.Text = "Restore"
    End Sub
    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        If Label1.Text = "" Then
            MsgBox("Please choose log menu on the left !", MsgBoxStyle.Information, "Office Tools")
        ElseIf Label1.Text = "Backup" Then
            ShowLog("Backup history", logPath)
        ElseIf Label1.Text = "Archive" Then
            ShowLog("Archive history", advLogPath)
        ElseIf Label1.Text = "Restore" Then
            ShowLog("Restore history", resLogPath)
        End If
    End Sub
    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        If Label1.Text = "" Then
            MsgBox("Please choose log menu on the left !", MsgBoxStyle.Information, "Office Tools")
        ElseIf Label1.Text = "Backup" Then
            ShowLog("Backup error history", errPath)
        ElseIf Label1.Text = "Archive" Then
            ShowLog("Archive error history", advErrPath)
        ElseIf Label1.Text = "Restore" Then
            ShowLog("Restore error history", resErrPath)
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim validation As Integer
        validation = MsgBox("This will clear all history !", vbExclamation + vbYesNo + vbDefaultButton2, "Office Tools")
        If validation = vbYes Then
            ClearLog(logPath, "Backup history")
            ClearLog(advLogPath, "Archive history")
            ClearLog(errPath, "Backup error history")
            ClearLog(advErrPath, "Archive error history")
            ClearLog(resLogPath, "Restore history")
            ClearLog(resErrPath, "Restore error history")
            ClearLog(roboLogPath, "Robocopy hustory")
            If New FileInfo(logPath).Length.Equals(0) Or New FileInfo(advLogPath).Length.Equals(0) Or New FileInfo(resLogPath).Length.Equals(0) Or New FileInfo(errPath).Length.Equals(0) Or New FileInfo(advErrPath).Length.Equals(0) Or New FileInfo(resErrPath).Length.Equals(0) Then
                MsgBox("All history cleared !", MsgBoxStyle.Information, "Office Tools")
                RichTextBox1.Text = ""
            End If
        Else
            MsgBox("Cancel clear history !", MsgBoxStyle.Information, "Office Tools")
        End If
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ExportLog(logPath, "log", "Backup history", logCountPath)
        ExportLog(advLogPath, "advLog", "Archive history", logCountPath)
        ExportLog(errPath, "err", "Backup error history", logCountPath)
        ExportLog(advErrPath, "advErr", "Archive error history", logCountPath)
        ExportLog(resLogPath, "resLog", "Restore history", logCountPath)
        ExportLog(resErrPath, "resErr", "Restore error history", logCountPath)
        ExportLog(roboLogPath, "robolog", "Robolog history", logCountPath)
        If CInt(PathVal(logCountPath, 0)) = 7 Then
            MsgBox("All log exported !", MsgBoxStyle.Information, "Office Tools")
        Else
            MsgBox("Export log error !", MsgBoxStyle.Critical, "Office Tools")
        End If
        PrepareNotif(logCountPath)
        RichTextBox1.Text = ""
    End Sub
    Private Sub ShowLog(log As String, path As String)
        RichTextBox1.Text = ""
        If File.Exists(path) Then
            If New FileInfo(path).Length.Equals(0) Then
                MsgBox(log & " is empty !", MsgBoxStyle.Critical, "Office Tools")
            Else
                RichTextBox1.Text = File.ReadAllText(path)
            End If
        Else
            MsgBox(log & " does not exist !", MsgBoxStyle.Critical, "Office Tools")
        End If
    End Sub
    Private Shared Sub ClearLog(log As String, log2 As String)
        If File.Exists(log) Then
            File.Delete(log)
            File.Create(log).Dispose()
        End If
    End Sub
    Private Shared Sub ExportLog(logpath As String, filename As String, logname As String, logCountPath As String)
        Dim logCount As Integer
        Dim curCount As Integer
        Dim totalCount As Integer
        If File.Exists(logpath) Then
            If File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) & "/" & filename & ".txt") Then
                File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) & "/" & filename & ".txt")
                File.Copy(logpath, Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) & "/" & filename & ".txt")
            Else
                File.Copy(logpath, Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) & "/" & filename & ".txt")
            End If
            logCount += 1
        End If
        curCount = CheckNull(logCountPath)
        totalCount = curCount + logCount
        PrepareNotif(logCountPath)
        Dim destwriter As New StreamWriter(logCountPath, True)
        destwriter.WriteLine(totalCount)
        destwriter.Close()
    End Sub
    Private Shared Function CheckNull(curCount As String) As Integer
        Dim result As Integer
        If PathVal(curCount, 0).Equals("null") Then
            result = 0
            Return result
        Else
            result = CInt(PathVal(curCount, 0))
            Return result
        End If
    End Function
End Class
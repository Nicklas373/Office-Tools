Imports System.IO
Imports Syncfusion.WinForms.Controls

Public Class HistoryMenu
    Inherits SfForm
    ReadOnly logPath As String = "log/log"
    ReadOnly advLogPath As String = "log/advlog"
    ReadOnly resLogPath As String = "log/reslog"
    ReadOnly roboLogPath As String = "log/robolog"
    ReadOnly errPath As String = "log/err"
    ReadOnly advErrPath As String = "log/adverr"
    ReadOnly resErrPath As String = "log/reserr"
    ReadOnly logCountPath As String = "log/expLog"
    Private Sub history_load(sender As Object, e As EventArgs) Handles MyBase.Load
        AllowTransparency = False
        Style.TitleBar.IconBackColor = Color.FromArgb(15, 161, 212)
        BackColor = Color.AliceBlue
        Style.TitleBar.TextHorizontalAlignment = HorizontalAlignment.Center
        Style.TitleBar.TextVerticalAlignment = VisualStyles.VerticalAlignment.Center
    End Sub
    Private Sub Close_Button(sender As Object, e As EventArgs) Handles Button4.Click
        Close()
    End Sub
    Private Sub Backup_Button(sender As Object, e As EventArgs) Handles Button5.Click
        RichTextBox1.Text = ""
        Label1.Text = "Backup"
    End Sub
    Private Sub Archive_Button(sender As Object, e As EventArgs) Handles Button1.Click
        RichTextBox1.Text = ""
        Label1.Text = "Archive"
    End Sub
    Private Sub Restore_Button(sender As Object, e As EventArgs) Handles Button2.Click
        RichTextBox1.Text = ""
        Label1.Text = "Restore"
    End Sub
    Private Sub Success_Specified_Button(sender As Object, e As EventArgs) Handles Button19.Click
        If Label1.Text = "" Then
            MsgBox("Please choose log menu on the left !", MsgBoxStyle.Information, "Office Tools")
        ElseIf Label1.Text = "Backup" Then
            RichTextBox1.Text = ShowLog("Backup history", logPath)
        ElseIf Label1.Text = "Archive" Then
            RichTextBox1.Text = ShowLog("Archive history", advLogPath)
        ElseIf Label1.Text = "Restore" Then
            RichTextBox1.Text = ShowLog("Restore history", resLogPath)
        End If
    End Sub
    Private Sub Error_Specified_Button(sender As Object, e As EventArgs) Handles Button20.Click
        If Label1.Text = "" Then
            MsgBox("Please choose log menu on the left !", MsgBoxStyle.Information, "Office Tools")
        ElseIf Label1.Text = "Backup" Then
            RichTextBox1.Text = ShowLog("Backup error history", errPath)
        ElseIf Label1.Text = "Archive" Then
            RichTextBox1.Text = ShowLog("Archive error history", advErrPath)
        ElseIf Label1.Text = "Restore" Then
            RichTextBox1.Text = ShowLog("Restore error history", resErrPath)
        End If
    End Sub
    Private Sub Clear_Log_Button(sender As Object, e As EventArgs) Handles Button3.Click
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
    Private Sub Export_Log_Button(sender As Object, e As EventArgs) Handles Button6.Click
        Dim validation As Integer
        validation = MsgBox("This will export all log !", vbExclamation + vbYesNo + vbDefaultButton2, "Office Tools")
        If validation = vbYes Then
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
        Else
            MsgBox("Cancel export log!", MsgBoxStyle.Information, "Office Tools")
        End If
    End Sub
End Class
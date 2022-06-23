Imports System.IO
Imports Microsoft.Win32.TaskScheduler

Public Class settings_menu
    Dim filedialog As New FolderBrowserDialog
    Dim confPath As String = "conf/config"
    Dim timePath As String = "conf/cli_backup/cliTimeInit"
    Dim cliSrcPath As String = "conf/cli_backup/cliSrcPath"
    Dim cliDestPath As String = "conf/cli_backup/cliDestPath"
    Dim cliDatePath As String = "conf/cli_backup/cliDatePath"
    Dim cliProcessor As String = "conf/cli_backup/cliProcessor"
    Private Sub Settings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AllowTransparency = False
        dir_bck_set.Visible = False
        daily_sched_pnl.Visible = False
        weekly_sched_pnl.Visible = False
        tsk_info_panel.Visible = False
        bck_info_pnl.Visible = False
        DateTimePicker4.Format = DateTimePickerFormat.Time
        DateTimePicker4.ShowUpDown = True
        DateTimePicker6.Format = DateTimePickerFormat.Time
        DateTimePicker6.ShowUpDown = True
        GetBackPref()
        WriteLogicalCount(cliProcessor)
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.ReadOnly = False
        TextBox2.ReadOnly = False
        Button1.Visible = True
        Button2.Visible = False
        Button3.Visible = True
        Label7.Visible = True
        ComboBox1.Visible = True
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        dir_bck_set.Visible = True
        daily_sched_pnl.Visible = False
        weekly_sched_pnl.Visible = False
        tsk_info_panel.Visible = False
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        dir_bck_set.Visible = True
        daily_sched_pnl.Visible = True
        weekly_sched_pnl.Visible = False
        tsk_info_panel.Visible = False
    End Sub
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        weekly_sched_pnl.Visible = True
        bck_info_pnl.Visible = False
        tsk_info_panel.Visible = False
    End Sub
    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        tsk_info_panel.Visible = True
        bck_info_pnl.Visible = True
    End Sub
    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        bck_info_pnl.Visible = False
        tsk_info_panel.Visible = True
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If TextBox1.ReadOnly = True Then
            MsgBox("Configuration menu is locked, Please click edit !", MsgBoxStyle.Information, "Office Tools")
        Else
            filedialog.InitialDirectory = Environment.SpecialFolder.UserProfile
            filedialog.ShowDialog()
        End If
    End Sub
    Private Sub FolderBrowserDialog1_Disposed(sender As Object, e As EventArgs) Handles Button7.Click
        If TextBox1.ReadOnly = False Then
            TextBox1.Text = filedialog.SelectedPath.ToString
        End If
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If TextBox2.ReadOnly = True Then
            MsgBox("Configuration menu is locked, Please click edit !", MsgBoxStyle.Information, "Office Tools")
        Else
            filedialog.InitialDirectory = Environment.SpecialFolder.UserProfile
            filedialog.ShowDialog()
        End If
    End Sub
    Private Sub FolderBrowserDialog2_Disposed(sender As Object, e As EventArgs) Handles Button6.Click
        If TextBox2.ReadOnly = False Then
            TextBox2.Text = filedialog.SelectedPath.ToString
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MsgBox("Please fill source folder !", MsgBoxStyle.Critical, "Office Tools")
        ElseIf TextBox2.Text = "" Then
            MsgBox("Please fill destination folder !", MsgBoxStyle.Critical, "Office Tools")
        Else
            If ComboBox1.Text = "" Then
                MsgBox("Please select backup time !", MsgBoxStyle.Critical, "Office Tools")
            Else
                If File.Exists(confPath) Then
                    GC.Collect()
                    GC.WaitForPendingFinalizers()
                    File.Delete(confPath)
                    File.Create(confPath).Dispose()
                    Dim writer As New StreamWriter(confPath, True)
                    writer.WriteLine("Office Tools Config v1.2")
                    writer.WriteLine("Source Directory: " & TextBox1.Text)
                    writer.WriteLine("Destination Directory: " & TextBox2.Text)
                    writer.WriteLine("Backup Preferences: " & ComboBox1.Text)
                    writer.Close()
                    WriteForAutoBackup(confPath, cliSrcPath, cliDestPath, cliDatePath, timePath)
                    MsgBox("Config created !", MsgBoxStyle.Information, "Office Tools")
                Else
                    File.Create(confPath).Dispose()
                    Dim writer As New StreamWriter(confPath, True)
                    writer.WriteLine("Office Tools Config v1.2")
                    writer.WriteLine("Source Directory: " & TextBox1.Text)
                    writer.WriteLine("Destination Directory: " & TextBox2.Text)
                    writer.WriteLine("Backup Preferences: " & ComboBox1.Text)
                    writer.Close()
                    WriteForAutoBackup(confPath, cliSrcPath, cliDestPath, cliDatePath, timePath)
                    MsgBox("Config created !", MsgBoxStyle.Information, "Office Tools")
                End If
                TextBox1.ReadOnly = True
                TextBox2.ReadOnly = True
                Button1.Visible = False
                Button2.Visible = True
                Button3.Visible = False
                Label7.Visible = False
                ComboBox1.Visible = False
            End If
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox1.ReadOnly = True
        TextBox1.Text = ""
        TextBox2.Text = ""
        ComboBox1.ResetText()
        TextBox2.ReadOnly = True
        Button1.Visible = False
        Button2.Visible = True
        Button3.Visible = False
        Label7.Visible = False
        ComboBox1.Visible = False
    End Sub
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Button11.Visible = True
        Button9.Visible = False
        Button10.Visible = True
        TextBox3.Text = ""
        ComboBox4.ResetText()
        ComboBox5.ResetText()
    End Sub
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Button11.Visible = False
        Button9.Visible = True
        Button10.Visible = False
        TextBox3.Text = ""
        ComboBox4.ResetText()
        ComboBox5.ResetText()
    End Sub
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If TextBox3.Text = "" Then
            MsgBox("Recurs day can not empty !", MsgBoxStyle.Critical, "Office Tools")
        Else
            Dim custdate As String = DateTimePicker3.Value.ToLongDateString & " " & DateTimePicker4.Value.ToLongTimeString
            If ComboBox4.Text = "Disabled" Then
                ComboBox5.ResetText()
                MsgBox("If repeat task is disabled, then repeat duration will be disable", vbExclamation, "Office Tools")
            End If
            If ComboBox5.Text = "Disabled" Then
                ComboBox4.ResetText()
                ComboBox5.ResetText()
                MsgBox("If repeat duration is disabled, then repeat task will be disable", vbExclamation, "Office Tools")
            End If
            DailyTrigger(custdate, 1, CInt(TextBox3.Text), CustRepDurValDaily(ComboBox5.Text), CustRepDurIntDaily(ComboBox4.Text), ComboBox5.Text, ComboBox4.Text)
            Button10.Visible = False
            Button9.Visible = True
            Button11.Visible = False
        End If
    End Sub
    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text = "0" Then
            MsgBox("Can not set 0 days for recurs day !", MsgBoxStyle.Critical, "Office Tools")
            MsgBox("Please set another value", MsgBoxStyle.Information, "Office Tools")
            TextBox3.Text = ""
        End If
    End Sub
    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        If ComboBox4.Text = "Disabled" Then
            ComboBox5.ResetText()
            MsgBox("If repeat task is disabled, then repeat duration will be disable", vbExclamation, "Office Tools")
        End If
    End Sub
    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        If ComboBox5.Text = "Disabled" Then
            ComboBox4.ResetText()
            ComboBox5.ResetText()
            MsgBox("If repeat duration is disabled, then repeat task will be disable", vbExclamation, "Office Tools")
        End If
    End Sub
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Button14.Visible = True
        Button15.Visible = True
        Button13.Visible = False
        TextBox4.Text = ""
        ComboBox6.ResetText()
        ComboBox7.ResetText()
        CheckBox7.Checked = False
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        CheckBox3.Checked = False
        CheckBox4.Checked = False
        CheckBox5.Checked = False
        CheckBox6.Checked = False
    End Sub
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Button14.Visible = False
        Button15.Visible = False
        Button13.Visible = True
        TextBox4.Text = ""
        ComboBox6.ResetText()
        ComboBox7.ResetText()
        CheckBox7.Checked = False
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        CheckBox3.Checked = False
        CheckBox4.Checked = False
        CheckBox5.Checked = False
        CheckBox6.Checked = False
    End Sub
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        If TextBox4.Text = "" Then
            MsgBox("Recurs week can not empty !", MsgBoxStyle.Critical, "Office Tools")
        Else
            If CheckBox7.Checked = False And CheckBox1.Checked = False And CheckBox2.Checked = False And CheckBox3.Checked = False And CheckBox4.Checked = False And CheckBox5.Checked = False And CheckBox6.Checked = False Then
                MsgBox("Recurs in day can not empty !", vbCritical, "Office Tools")
            Else
                Dim custdate As String = DateTimePicker5.Value.ToLongDateString & " " & DateTimePicker6.Value.ToLongTimeString
                If ComboBox7.Text = "Disabled" Then
                    ComboBox6.ResetText()
                    MsgBox("If repeat task is disabled, then repeat duration will be disable", vbExclamation, "Office Tools")
                End If
                If ComboBox6.Text = "Disabled" Then
                    ComboBox7.ResetText()
                    ComboBox6.ResetText()
                    MsgBox("If repeat duration is disabled, then repeat task will be disable", vbExclamation, "Office Tools")
                End If
                WeeklyTrigger(custdate, 1, CInt(TextBox4.Text), CustRepDurValWeek(ComboBox6.Text), CustRepDurIntWeek(ComboBox7.Text), Cb1, Cb2, Cb3, Cb4, Cb5, Cb6, Cb7, ComboBox6.Text, ComboBox7.Text)
                Button6.Visible = False
                Button7.Visible = True
                Button8.Visible = False
            End If
        End If
    End Sub
    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        ShowTask("Office Tools")
    End Sub
    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Using tService As New TaskService()
            Dim tTask As Task = tService.GetTask("Office Tools")
            If tTask Is Nothing Then
                MsgBox("Office Tools scheduler not exist !", MsgBoxStyle.Information, "Office Tools")
            Else
                tService.RootFolder.DeleteTask("Office Tools")
                RichTextBox1.Text = ""
                MsgBox("Office Tools scheduler removed !", MsgBoxStyle.Information, "Office Tools")
            End If
        End Using
    End Sub
    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        If File.Exists(confPath) Then
            Using tService As New TaskService()
                Dim tTask As Task = tService.GetTask("Office Tools")
                If tTask Is Nothing Then
                    MsgBox("Office Tools scheduler not exist !", MsgBoxStyle.Critical, "Office Tools")
                    MsgBox("Please create new scheduler first !", MsgBoxStyle.Critical, "Office Tools")
                Else
                    tTask.Run()
                    MsgBox("Office Tools is running !", MsgBoxStyle.Information, "Office Tools")
                End If
            End Using
        Else
            MsgBox("Config file is not exist !, Please configure directory first !", MsgBoxStyle.Critical, "Office Tools")
        End If
    End Sub
    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        ShowLog("Config", confPath)
    End Sub
    Private Sub ShowLog(log As String, path As String)
        RichTextBox2.Text = ""
        If File.Exists(path) Then
            If New FileInfo(path).Length.Equals(0) Then
                MsgBox(log & " file is empty !", MsgBoxStyle.Critical, "MigrateToGDrive")
            Else
                RichTextBox2.Text = File.ReadAllText(path)
            End If
        Else
            MsgBox(log & " file does not exist !", MsgBoxStyle.Critical, "MigrateToGDrive")
        End If
    End Sub
    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        ClearLog(confPath, "Config")
        If File.Exists(timePath) Then
            GC.Collect()
            GC.WaitForPendingFinalizers()
            File.Delete(timePath)
            File.Create(timePath).Dispose()
            Dim timeWriter As New StreamWriter(timePath, True)
            timeWriter.WriteLine("null")
            timeWriter.Close()
        Else
            GC.Collect()
            GC.WaitForPendingFinalizers()
            File.Delete(timePath)
            File.Create(timePath).Dispose()
            Dim timeWriter As New StreamWriter(timePath, True)
            timeWriter.WriteLine("null")
            timeWriter.Close()
        End If
        TextBox1.Text = ""
        TextBox2.Text = ""
        ComboBox1.ResetText()
        RichTextBox2.Text = ""
    End Sub
    Private Function Cb1() As DaysOfTheWeek
        Dim monday As DaysOfTheWeek
        If CheckBox1.Checked Then
            monday = DaysOfTheWeek.Monday
            Return monday
        Else
            monday = DaysOfTheWeek.Sunday
            Return monday
        End If
    End Function
    Private Function Cb2() As DaysOfTheWeek
        Dim tuesday As DaysOfTheWeek
        If CheckBox2.Checked Then
            tuesday = DaysOfTheWeek.Tuesday
            Return tuesday
        Else
            tuesday = DaysOfTheWeek.Sunday
            Return tuesday
        End If
    End Function
    Private Function Cb3() As DaysOfTheWeek
        Dim wednesday As DaysOfTheWeek
        If CheckBox3.Checked Then
            wednesday = DaysOfTheWeek.Wednesday
            Return wednesday
        Else
            wednesday = DaysOfTheWeek.Sunday
            Return wednesday
        End If
    End Function
    Private Function Cb4() As DaysOfTheWeek
        Dim thursday As DaysOfTheWeek
        If CheckBox4.Checked Then
            thursday = DaysOfTheWeek.Thursday
            Return thursday
        Else
            thursday = DaysOfTheWeek.Sunday
            Return thursday
        End If
    End Function
    Private Function Cb5() As DaysOfTheWeek
        Dim friday As DaysOfTheWeek
        If CheckBox5.Checked Then
            friday = DaysOfTheWeek.Friday
            Return friday
        Else
            friday = DaysOfTheWeek.Sunday
            Return friday
        End If
    End Function
    Private Function Cb6() As DaysOfTheWeek
        Dim saturday As DaysOfTheWeek
        If CheckBox6.Checked Then
            saturday = DaysOfTheWeek.Saturday
            Return saturday
        Else
            saturday = DaysOfTheWeek.Sunday
            Return saturday
        End If
    End Function
    Private Function Cb7() As DaysOfTheWeek
        Dim sunday As DaysOfTheWeek
        If CheckBox7.Checked Then
            sunday = DaysOfTheWeek.Sunday
            Return sunday
        Else
            sunday = DaysOfTheWeek.Sunday
            Return sunday
        End If
    End Function
    Private Sub GetBackPref()
        If File.Exists(confPath) Then
            If PathVal(confPath, 1).Equals("null") Then
                TextBox1.Text = ""
            Else
                TextBox1.Text = Replace(PathVal(confPath, 1), "Source Directory: ", "")
            End If
            If PathVal(confPath, 2).Equals("null") Then
                TextBox2.Text = ""
            Else
                TextBox2.Text = Replace(PathVal(confPath, 2), "Destination Directory: ", "")
            End If
            If PathVal(confPath, 3).Replace("Backup Preferences: ", "").Equals("Anytime") Then
                ComboBox1.SelectedIndex = 0
            ElseIf PathVal(confPath, 3).Replace("Backup Preferences: ", "").Equals("Today") Then
                ComboBox1.SelectedIndex = 1
            End If
        End If
    End Sub
    Private Sub ShowTask(taskName As String)
        Using tService As New TaskService()
            Dim tTask As Task = tService.GetTask(taskName)
            If tTask Is Nothing Then
                MsgBox("Office Task scheduler not exist !", MsgBoxStyle.Critical, "Office Tools")
                MsgBox("Please create new scheduler first !", MsgBoxStyle.Critical, "Office Tools")
                RichTextBox1.Text = ""
            Else
                RichTextBox1.Text = "Task Name: " & tTask.Name & vbCrLf & "Task State: " & tTask.State.ToString & vbCrLf &
                                    "Task Path: " & tTask.Path.ToString & vbCrLf &
                                    "Next Runtime: " & tTask.NextRunTime.ToLongDateString & " " & tTask.NextRunTime.ToLongTimeString & vbCrLf &
                                    "Last Runtime: " & tTask.LastRunTime.ToLongDateString & " " & tTask.LastRunTime.ToLongTimeString & vbCrLf &
                                    "Last Task Result: " & tTask.LastTaskResult.ToString & vbCrLf &
                                    "Total Failed Task: " & tTask.NumberOfMissedRuns.ToString & vbCrLf
            End If
        End Using
    End Sub
End Class
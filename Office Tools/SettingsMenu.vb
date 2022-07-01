Imports System.IO
Imports Microsoft.Win32.TaskScheduler
Imports Syncfusion.Windows.Forms
Imports Syncfusion.WinForms.Controls
Public Class SettingsMenu
    Inherits SfForm
    Dim openfiledialog As New OpenFileDialog
    Dim openfolderdialog As New FolderBrowserDialog
    Dim confPath As String = "conf/config"
    Dim timePath As String = "conf/cli_backup/cliTimeInit"
    Dim cliSrcPath As String = "conf/cli_backup/cliSrcPath"
    Dim cliDestPath As String = "conf/cli_backup/cliDestPath"
    Dim cliDatePath As String = "conf/cli_backup/cliDatePath"
    Dim cliProcessor As String = "conf/cli_backup/cliProcessor"
    Private Sub Settings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AllowTransparency = False
        Style.TitleBar.IconBackColor = Color.FromArgb(15, 161, 212)
        BackColor = Color.AliceBlue
        Style.TitleBar.TextHorizontalAlignment = HorizontalAlignment.Center
        Style.TitleBar.TextVerticalAlignment = VisualStyles.VerticalAlignment.Center
        dir_bck_set.Visible = False
        daily_sched_pnl.Visible = False
        weekly_sched_pnl.Visible = False
        tsk_info_panel.Visible = False
        bck_info_pnl.Visible = False
        SfDateTimeEdit1.ShowUpDown = True
        DateTimePicker4.Format = DateTimePickerFormat.Time
        DateTimePicker4.ShowUpDown = True
        SfDateTimeEdit2.ShowUpDown = True
        DateTimePicker6.Format = DateTimePickerFormat.Time
        DateTimePicker6.ShowUpDown = True
        GetBackPref()
        WriteLogicalCount(cliProcessor)
        MessageBoxAdv.MessageBoxStyle = MessageBoxAdv.Style.Metro
    End Sub
    Private Sub Edit_Folder_Backup_Dir_Settings_Handler(sender As Object, e As EventArgs) Handles Button2.Click
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
        bck_info_pnl.Visible = False
        tsk_info_panel.Visible = False
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        dir_bck_set.Visible = True
        daily_sched_pnl.Visible = True
        weekly_sched_pnl.Visible = False
        bck_info_pnl.Visible = False
        tsk_info_panel.Visible = False
    End Sub
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        dir_bck_set.Visible = True
        daily_sched_pnl.Visible = True
        weekly_sched_pnl.Visible = True
        bck_info_pnl.Visible = False
        tsk_info_panel.Visible = False
    End Sub
    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        dir_bck_set.Visible = True
        daily_sched_pnl.Visible = True
        weekly_sched_pnl.Visible = True
        bck_info_pnl.Visible = True
        tsk_info_panel.Visible = False
    End Sub
    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        dir_bck_set.Visible = True
        daily_sched_pnl.Visible = True
        weekly_sched_pnl.Visible = True
        bck_info_pnl.Visible = True
        tsk_info_panel.Visible = True
    End Sub
    Private Sub Button24_Click(sender As Object, e As EventArgs)
        dir_bck_set.Visible = True
        daily_sched_pnl.Visible = True
        weekly_sched_pnl.Visible = True
        bck_info_pnl.Visible = True
        tsk_info_panel.Visible = True
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Close()
    End Sub
    Private Sub Source_Folder_Backup_Dir_Settings_Button(sender As Object, e As EventArgs) Handles Button7.Click
        If TextBox1.ReadOnly = True Then
            MessageBoxAdv.Show("Configuration menu is locked, Please click edit !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            openfolderdialog.InitialDirectory = Environment.SpecialFolder.UserProfile
            openfolderdialog.ShowDialog()
        End If
    End Sub
    Private Sub Source_Folder_Backup_Dir_Settings_Handler(sender As Object, e As EventArgs) Handles Button7.Click
        If TextBox1.ReadOnly = False Then
            TextBox1.Text = openfolderdialog.SelectedPath.ToString
        End If
    End Sub
    Private Sub Destination_Folder_Backup_Dir_Settings_Button(sender As Object, e As EventArgs) Handles Button6.Click
        If TextBox2.ReadOnly = True Then
            MessageBoxAdv.Show("Configuration menu is locked, Please click edit !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            openfolderdialog.InitialDirectory = Environment.SpecialFolder.UserProfile
            openfolderdialog.ShowDialog()
        End If
    End Sub
    Private Sub Destination_Folder_Backup_Dir_Settings_Handler(sender As Object, e As EventArgs) Handles Button6.Click
        If TextBox2.ReadOnly = False Then
            TextBox2.Text = openfolderdialog.SelectedPath.ToString
        End If
    End Sub
    Private Sub Save_Folder_Backup_Dir_Settings_Handler(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MessageBoxAdv.Show("Please fill source folder !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf TextBox2.Text = "" Then
            MessageBoxAdv.Show("Please fill destination folder !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If ComboBox1.Text = "" Then
                MessageBoxAdv.Show("Please select backup time !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                Dim SourceBackup As String = FindConfig(confPath, "Source Directory: ")
                Dim DestBackup As String = FindConfig(confPath, "Destination Directory: ")
                Dim BackupPreferences As String = FindConfig(confPath, "Backup Preferences: ")
                If SourceBackup = "null" Then
                    Dim writer As New StreamWriter(confPath, True)
                    writer.WriteLine("Source Directory: " & TextBox1.Text)
                    writer.Close()
                Else
                    Dim SourceReaderOldConf As String = File.ReadAllText(confPath)
                    SourceReaderOldConf = SourceReaderOldConf.Replace(SourceBackup, "Source Directory: " & TextBox1.Text)
                    File.WriteAllText(confPath, SourceReaderOldConf)
                End If
                If DestBackup = "null" Then
                    Dim writer As New StreamWriter(confPath, True)
                    writer.WriteLine("Destination Directory: " & TextBox2.Text)
                    writer.Close()
                Else
                    Dim DestinationReaderOldConf As String = File.ReadAllText(confPath)
                    DestinationReaderOldConf = DestinationReaderOldConf.Replace(DestBackup, "Destination Directory: " & TextBox2.Text)
                    File.WriteAllText(confPath, DestinationReaderOldConf)
                End If
                If BackupPreferences = "null" Then
                    Dim writer As New StreamWriter(confPath, True)
                    writer.WriteLine("Backup Preferences: " & ComboBox1.Text)
                    writer.Close()
                Else
                    Dim BackupPreferencesOldConf As String = File.ReadAllText(confPath)
                    BackupPreferencesOldConf = BackupPreferencesOldConf.Replace(BackupPreferences, "Backup Preferences: " & ComboBox1.Text)
                    File.WriteAllText(confPath, BackupPreferencesOldConf)
                End If
                WriteForAutoBackup(confPath, cliSrcPath, cliDestPath, cliDatePath, timePath)
                MessageBoxAdv.Show("Config updated !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
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
    Private Sub Cancel_Folder_Backup_Dir_Settings_Handler(sender As Object, e As EventArgs) Handles Button1.Click
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
    Private Sub Edit_Daily_Sched_Settings_Handler(sender As Object, e As EventArgs) Handles Button9.Click
        Button11.Visible = True
        Button9.Visible = False
        Button10.Visible = True
        SfDateTimeEdit1.Enabled = True
        DateTimePicker4.Enabled = True
        TextBox3.Enabled = True
        TextBox3.Text = ""
        ComboBox4.Enabled = True
        ComboBox5.Enabled = True
        ComboBox4.ResetText()
        ComboBox5.ResetText()
    End Sub
    Private Sub Cancel_Daily_Sched_Settings_Handler(sender As Object, e As EventArgs) Handles Button10.Click
        Button11.Visible = False
        Button9.Visible = True
        Button10.Visible = False
        SfDateTimeEdit1.Enabled = False
        DateTimePicker4.Enabled = False
        TextBox3.Enabled = False
        TextBox3.Text = ""
        ComboBox4.Enabled = True
        ComboBox5.Enabled = True
        ComboBox4.ResetText()
        ComboBox5.ResetText()
    End Sub
    Private Sub Save_Daily_Sched_Settings_Handler(sender As Object, e As EventArgs) Handles Button11.Click
        If TextBox3.Text = "" Then
            MessageBoxAdv.Show("Recurs day can not empty !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Dim custdate As String = SfDateTimeEdit1.Value & " " & DateTimePicker4.Value.ToLongTimeString
            If ComboBox4.Text = "Disabled" Then
                ComboBox5.ResetText()
                MessageBoxAdv.Show("If repeat task is disabled, then repeat duration will be disable", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
            If ComboBox5.Text = "Disabled" Then
                ComboBox4.ResetText()
                ComboBox5.ResetText()
                MessageBoxAdv.Show("If repeat duration is disabled, then repeat task will be disable", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
            DailyTrigger(custdate, 1, CInt(TextBox3.Text), CustRepDurValDaily(ComboBox5.Text), CustRepDurIntDaily(ComboBox4.Text), ComboBox5.Text, ComboBox4.Text)
            Button10.Visible = False
            Button9.Visible = True
            Button11.Visible = False
            SfDateTimeEdit1.Enabled = False
            DateTimePicker4.Enabled = False
            TextBox3.Enabled = False
            ComboBox4.Enabled = False
            ComboBox5.Enabled = False
        End If
    End Sub
    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If IsNumeric(TextBox3.Text) Then
            If TextBox3.Text = "0" Then
                MessageBoxAdv.Show("Can not set 0 days for recurs day !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                MessageBoxAdv.Show("Please set another value", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
                TextBox3.Text = ""
            End If
        Else
            MessageBoxAdv.Show("Recurs day input only for numbers !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        If ComboBox4.Text = "Disabled" Then
            ComboBox5.ResetText()
            MessageBoxAdv.Show("If repeat task is disabled, then repeat duration will be disable", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub
    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        If ComboBox5.Text = "Disabled" Then
            ComboBox4.ResetText()
            ComboBox5.ResetText()
            MessageBoxAdv.Show("If repeat duration is disabled, then repeat task will be disable", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Button14.Visible = True
        Button15.Visible = True
        Button13.Visible = False
        SfDateTimeEdit2.Enabled = True
        DateTimePicker6.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Text = ""
        ComboBox6.Enabled = True
        ComboBox7.Enabled = True
        ComboBox6.ResetText()
        ComboBox7.ResetText()
        CheckBox1.Enabled = True
        CheckBox2.Enabled = True
        CheckBox3.Enabled = True
        CheckBox4.Enabled = True
        CheckBox5.Enabled = True
        CheckBox6.Enabled = True
        CheckBox7.Enabled = True
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
        SfDateTimeEdit2.Enabled = False
        DateTimePicker6.Enabled = False
        TextBox4.Enabled = False
        TextBox4.Text = ""
        ComboBox6.Enabled = False
        ComboBox7.Enabled = False
        ComboBox6.ResetText()
        ComboBox7.ResetText()
        CheckBox1.Enabled = False
        CheckBox2.Enabled = False
        CheckBox3.Enabled = False
        CheckBox4.Enabled = False
        CheckBox5.Enabled = False
        CheckBox6.Enabled = False
        CheckBox7.Enabled = False
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
            MessageBoxAdv.Show("Recurs week can not empty !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If CheckBox7.Checked = False And CheckBox1.Checked = False And CheckBox2.Checked = False And CheckBox3.Checked = False And CheckBox4.Checked = False And CheckBox5.Checked = False And CheckBox6.Checked = False Then
                MessageBoxAdv.Show("Recurs in day can not empty !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                Dim custdate As String = SfDateTimeEdit2.Value & " " & DateTimePicker6.Value.ToLongTimeString
                If ComboBox7.Text = "Disabled" Then
                    ComboBox6.ResetText()
                    MessageBoxAdv.Show("If repeat task is disabled, then repeat duration will be disable", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If
                If ComboBox6.Text = "Disabled" Then
                    ComboBox7.ResetText()
                    ComboBox6.ResetText()
                    MessageBoxAdv.Show("If repeat duration is disabled, then repeat task will be disable", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If
                WeeklyTrigger(custdate, 1, CInt(TextBox4.Text), CustRepDurValWeek(ComboBox6.Text), CustRepDurIntWeek(ComboBox7.Text), Cb1(CheckBox1.Checked), Cb2(CheckBox2.Checked), Cb3(CheckBox3.Checked), Cb4(CheckBox4.Checked), Cb5(CheckBox5.Checked), Cb6(CheckBox6.Checked), Cb7(CheckBox7.Checked), ComboBox6.Text, ComboBox7.Text)
                Button6.Visible = False
                Button7.Visible = True
                Button8.Visible = False
                CheckBox1.Enabled = False
                CheckBox2.Enabled = False
                CheckBox3.Enabled = False
                CheckBox4.Enabled = False
                CheckBox5.Enabled = False
                CheckBox6.Enabled = False
                CheckBox7.Enabled = False
            End If
        End If
    End Sub
    Private Sub Chk_Task_Settings_Button(sender As Object, e As EventArgs) Handles Button19.Click
        RichTextBox1.Text = ShowTask("Office Tools")
    End Sub
    Private Sub Remove_Task_Settings_Button(sender As Object, e As EventArgs) Handles Button20.Click
        Using tService As New TaskService()
            Dim tTask As Task = tService.GetTask("Office Tools")
            If tTask Is Nothing Then
                MessageBoxAdv.Show("Office Tools scheduler not exist !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                tService.RootFolder.DeleteTask("Office Tools")
                RichTextBox1.Text = ""
                MessageBoxAdv.Show("Office Tools scheduler removed !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Using
    End Sub
    Private Sub Run_Task_Settings_Button(sender As Object, e As EventArgs) Handles Button17.Click
        If File.Exists(confPath) Then
            Using tService As New TaskService()
                Dim tTask As Task = tService.GetTask("Office Tools")
                If tTask Is Nothing Then
                    MessageBoxAdv.Show("Office Tools scheduler not exist !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    MessageBoxAdv.Show("Please create new scheduler first !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    tTask.Run()
                    MessageBoxAdv.Show("Office Tools is running !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using
        Else
            MessageBoxAdv.Show("Config file is not exist !, Please configure directory first !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
    Private Sub Chk_Backup_Settings_Button(sender As Object, e As EventArgs) Handles Button22.Click
        RichTextBox2.Text = ShowLog("Config", confPath)
    End Sub
    Private Sub Remove_Backup_Settings_Button(sender As Object, e As EventArgs) Handles Button23.Click
        ClearLog(confPath, "Config", True)
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
    Private Sub GetBackPref()
        Dim SourceBackup As String = FindConfig(confPath, "Source Directory: ")
        Dim DestBackup As String = FindConfig(confPath, "Destination Directory: ")
        Dim BackupPreferences As String = FindConfig(confPath, "Backup Preferences: ")
        Dim PDFReaderConf As String = FindConfig(confPath, "PDF Reader Preferences: ")
        Dim PDFAutoConf As String = FindConfig(confPath, "Auto Open PDF: ")
        If SourceBackup = "null" Then
            TextBox1.Text = ""
        Else
            TextBox1.Text = SourceBackup.Remove(0, 18)
        End If
        If DestBackup = "null" Then
            TextBox2.Text = ""
        Else
            TextBox2.Text = DestBackup.Remove(0, 23)
        End If
        If BackupPreferences = "null" Then
            ComboBox1.Text = ""
        Else
            If BackupPreferences = "Backup Preferences: Anytime" Then
                ComboBox1.Text = "Anytime"
            ElseIf BackupPreferences = "Backup Preferences: Today" Then
                ComboBox1.Text = "Today"
            End If
        End If
    End Sub
End Class
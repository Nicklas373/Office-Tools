Imports Syncfusion.WinForms.Controls
Imports System.IO
Imports Microsoft.Win32.TaskScheduler
Imports Syncfusion.Windows.Forms

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
        DateTimePicker4.Format = DateTimePickerFormat.Time
        DateTimePicker4.ShowUpDown = True
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
        pdf_set_pnl.Visible = False
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        dir_bck_set.Visible = True
        daily_sched_pnl.Visible = True
        weekly_sched_pnl.Visible = False
        bck_info_pnl.Visible = False
        tsk_info_panel.Visible = False
        pdf_set_pnl.Visible = False
    End Sub
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        dir_bck_set.Visible = True
        daily_sched_pnl.Visible = True
        weekly_sched_pnl.Visible = True
        bck_info_pnl.Visible = False
        tsk_info_panel.Visible = False
        pdf_set_pnl.Visible = False
    End Sub
    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        dir_bck_set.Visible = True
        daily_sched_pnl.Visible = True
        weekly_sched_pnl.Visible = True
        bck_info_pnl.Visible = True
        tsk_info_panel.Visible = False
        pdf_set_pnl.Visible = False
    End Sub
    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        dir_bck_set.Visible = True
        daily_sched_pnl.Visible = True
        weekly_sched_pnl.Visible = True
        bck_info_pnl.Visible = True
        tsk_info_panel.Visible = True
        pdf_set_pnl.Visible = False
    End Sub
    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        dir_bck_set.Visible = True
        daily_sched_pnl.Visible = True
        weekly_sched_pnl.Visible = True
        bck_info_pnl.Visible = True
        tsk_info_panel.Visible = True
        pdf_set_pnl.Visible = True
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
                If File.Exists(confPath) Then
                    If File.ReadAllLines(confPath).Length < 1 Then
                        File.Create(confPath).Dispose()
                        Dim writer As New StreamWriter(confPath, True)
                        writer.WriteLine("Office Tools Config v1.2")
                        writer.WriteLine("Source Directory: " & TextBox1.Text)
                        writer.WriteLine("Destination Directory: " & TextBox2.Text)
                        writer.WriteLine("Backup Preferences: " & ComboBox1.Text)
                        writer.Close()
                        MessageBoxAdv.Show("Config created !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        WriteFile(confPath, 1, "Source Directory: " & TextBox1.Text)
                        WriteFile(confPath, 2, "Destination Directory: " & TextBox2.Text)
                        WriteFile(confPath, 3, "Backup Preferences: " & ComboBox1.Text)
                        MessageBoxAdv.Show("Config updated !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Else
                    File.Create(confPath).Dispose()
                    Dim writer As New StreamWriter(confPath, True)
                    writer.WriteLine("Office Tools Config v1.2")
                    writer.WriteLine("Source Directory: " & TextBox1.Text)
                    writer.WriteLine("Destination Directory: " & TextBox2.Text)
                    writer.WriteLine("Backup Preferences: " & ComboBox1.Text)
                    writer.Close()
                    MessageBoxAdv.Show("Config created !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
                WriteForAutoBackup(confPath, cliSrcPath, cliDestPath, cliDatePath, timePath)
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
        TextBox3.Text = ""
        ComboBox4.ResetText()
        ComboBox5.ResetText()
    End Sub
    Private Sub Cancel_Daily_Sched_Settings_Handler(sender As Object, e As EventArgs) Handles Button10.Click
        Button11.Visible = False
        Button9.Visible = True
        Button10.Visible = False
        TextBox3.Text = ""
        ComboBox4.ResetText()
        ComboBox5.ResetText()
    End Sub
    Private Sub Save_Daily_Sched_Settings_Handler(sender As Object, e As EventArgs) Handles Button11.Click
        If TextBox3.Text = "" Then
            MessageBoxAdv.Show("Recurs day can not empty !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Dim custdate As String = DateTimePicker3.Value.ToLongDateString & " " & DateTimePicker4.Value.ToLongTimeString
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
        End If
    End Sub
    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text = "0" Then
            MessageBoxAdv.Show("Can not set 0 days for recurs day !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBoxAdv.Show("Please set another value", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
            TextBox3.Text = ""
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
            MessageBoxAdv.Show("Recurs week can not empty !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If CheckBox7.Checked = False And CheckBox1.Checked = False And CheckBox2.Checked = False And CheckBox3.Checked = False And CheckBox4.Checked = False And CheckBox5.Checked = False And CheckBox6.Checked = False Then
                MessageBoxAdv.Show("Recurs in day can not empty !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                Dim custdate As String = DateTimePicker5.Value.ToLongDateString & " " & DateTimePicker6.Value.ToLongTimeString
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
    Private Sub Edit_PDF_Browse_Settings(sender As Object, e As EventArgs) Handles Button26.Click
        TextBox5.ReadOnly = False
        Button28.Visible = True
        Button27.Visible = True
        Button26.Visible = False
    End Sub
    Private Sub Cancel_PDF_Browse_Settings(sender As Object, e As EventArgs) Handles Button27.Click
        TextBox5.Text = ""
        TextBox5.ReadOnly = True
        Button27.Visible = False
        Button28.Visible = False
        Button26.Visible = True
    End Sub
    Private Sub Save_PDF_Browse_Settings(sender As Object, e As EventArgs) Handles Button28.Click
        If TextBox5.Text = "" Then
            MsgBox("PDF Path is empty, please choose PDF Reader !", MsgBoxStyle.Information, "Office Tools")
        Else
            If File.Exists(confPath) Then
                If File.ReadAllLines(confPath).Length < 5 Then
                    My.Computer.FileSystem.WriteAllText(confPath, "PDF Reader Preferences: " & TextBox5.Text, True)
                    My.Computer.FileSystem.WriteAllText(confPath, vbCrLf & "Auto Open PDF: " & CheckBox8.Checked, True)
                Else
                    WriteFile(confPath, 4, "PDF Reader Preferences: " & TextBox5.Text)
                    WriteFile(confPath, 5, "Auto Open PDF: " & CheckBox8.Checked)
                End If
                WriteForAutoBackup(confPath, cliSrcPath, cliDestPath, cliDatePath, timePath)
                MessageBoxAdv.Show("Config updated !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBoxAdv.Show("Config not found !, Please create on backup location settings !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
        Button27.Visible = False
        Button28.Visible = False
        TextBox5.ReadOnly = True
        Button26.Visible = True
    End Sub
    Private Sub Pdf_Settings_Browse_Button(sender As Object, e As EventArgs) Handles Button25.Click
        If TextBox5.ReadOnly = True Then
            MessageBoxAdv.Show("Configuration menu is locked, Please click edit !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            openfiledialog.InitialDirectory = Environment.SpecialFolder.UserProfile
            openfiledialog.DefaultExt = ".exe"
            openfiledialog.Filter = "Executables Files|*.exe"
            openfiledialog.ShowDialog()
        End If
    End Sub
    Private Sub Pdf_Settings_Browse_Handler(sender As Object, e As EventArgs) Handles Button25.Click
        If TextBox5.ReadOnly = False Then
            TextBox5.Text = openfiledialog.FileName.ToString
        End If
    End Sub
    Private Sub GetBackPref()
        If File.Exists(confPath) Then
            If File.ReadAllLines(confPath).Length < 5 Then
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
                If PathVal(confPath, 3).Equals("null") Then
                    If PathVal(confPath, 3).Replace("Backup Preferences: ", "").Equals("Anytime") Then
                        ComboBox1.Text = "Anytime"
                    ElseIf PathVal(confPath, 3).Replace("Backup Preferences: ", "").Equals("Today") Then
                        ComboBox1.Text = "Today"
                    End If
                End If
            Else
                If PathVal(confPath, 4).Equals("null") Then
                    TextBox5.Text = ""
                Else
                    TextBox5.Text = Replace(PathVal(confPath, 4), "PDF Reader Preferences: ", "")
                End If
                If PathVal(confPath, 5).Equals("null") Then
                    CheckBox8.Checked = False
                Else
                    If PathVal(confPath, 5).Replace("Auto Open PDF: ", "").Equals("False") Then
                        CheckBox8.Checked = False
                    Else
                        CheckBox8.Checked = True
                    End If
                End If
            End If
        End If
    End Sub
End Class
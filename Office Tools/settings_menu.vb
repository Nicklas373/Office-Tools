Imports System.IO
Imports Microsoft.Win32.TaskScheduler
Public Class settings_menu
    Dim openfiledialog As New OpenFileDialog
    Dim openfolderdialog As New FolderBrowserDialog
    Dim confPath As String = "conf/config"
    Dim pdfPath As String = "conf/pdf"
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
            MsgBox("Configuration menu is locked, Please click edit !", MsgBoxStyle.Information, "Office Tools")
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
            MsgBox("Configuration menu is locked, Please click edit !", MsgBoxStyle.Information, "Office Tools")
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
                MsgBox("Office Tools scheduler not exist !", MsgBoxStyle.Information, "Office Tools")
            Else
                tService.RootFolder.DeleteTask("Office Tools")
                RichTextBox1.Text = ""
                MsgBox("Office Tools scheduler removed !", MsgBoxStyle.Information, "Office Tools")
            End If
        End Using
    End Sub
    Private Sub Run_Task_Settings_Button(sender As Object, e As EventArgs) Handles Button17.Click
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
    Private Sub Chk_Backup_Settings_Button(sender As Object, e As EventArgs) Handles Button22.Click
        RichTextbox2.Text = ShowLog("Config", confPath)
    End Sub
    Private Sub Remove_Backup_Settings_Button(sender As Object, e As EventArgs) Handles Button23.Click
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
            CheckFileExist(pdfPath, TextBox5.Text)
        End If
        Button27.Visible = False
        Button28.Visible = False
        TextBox5.ReadOnly = True
        Button26.Visible = True
    End Sub
    Private Sub Pdf_Settings_Browse_Button(sender As Object, e As EventArgs) Handles Button25.Click
        If TextBox5.ReadOnly = True Then
            MsgBox("Configuration menu is locked, Please click edit !", MsgBoxStyle.Information, "Office Tools")
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
                    ComboBox1.SelectedIndex = 0
                ElseIf PathVal(confPath, 3).Replace("Backup Preferences: ", "").Equals("Today") Then
                    ComboBox1.SelectedIndex = 1
                End If
            End If
        End If
        If File.Exists(pdfPath) Then
            If PathVal(pdfPath, 0).Equals("null") Then
                TextBox5.Text = ""
            Else
                TextBox5.Text = PathVal(pdfPath, 0)
            End If
        End If
    End Sub
End Class
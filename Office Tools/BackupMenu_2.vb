Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Syncfusion.Windows.Forms
Imports Syncfusion.WinForms.Controls
Imports Syncfusion.WinForms.Input.Enums
Public Class BackupMenu_2
    Inherits SfForm
    Dim openfiledialog As New OpenFileDialog
    Dim openfolderdialog As New FolderBrowserDialog
    Dim savefiledialog As New SaveFileDialog
    Dim logPath As String = "log/log"
    Dim errPath As String = "log/err"
    Dim lastResult As String = "log/lastResult"
    Dim lastErr As String = "log/lastErr"
    Dim roboPath As String = "log/robolog"
    Dim advLogPath As String = "log/advlog"
    Dim resLogPath As String = "log/reslog"
    Dim advErrPath As String = "log/adverr"
    Dim resErrPath As String = "log/reserr"
    Dim resSrcPath As String = "conf/res_backup/resSrcPath"
    Dim resDestPath As String = "conf/res_backup/resDestPath"
    Dim resInstPath As String = "conf/res_backup/resInstPath"
    Dim resTempPass As String = "conf/res_backup/resTempKey"
    Dim resDecResult As String = "conf/res_backup/resDecResult"
    Dim resZipLog As String = "conf/res_backup/resZipLog"
    Dim resDecMtd As String = "conf/res_backup/resDecType"
    Dim advSrcPath As String = "conf/adv_backup/advSrcPath"
    Dim advDestPath As String = "conf/adv_backup/advDestPath"
    Dim advCompExt As String = "conf/adv_backup/advCompExt"
    Dim advCompLevel As String = "conf/adv_backup/advCompLvl"
    Dim advCompType As String = "conf/adv_backup/advCompType"
    Dim advEncType As String = "conf/adv_backup/advEncType"
    Dim advInstPath As String = "conf/adv_backup/advInstPath"
    Dim advRanStrg As String = "conf/adv_backup/advRandomStrg"
    Dim advTempPass As String = "conf/adv_backup/advTempPass"
    Dim uiProcessorCount As String = "conf/nrm_backup/nrmProcessor"
    Dim uiSpecFilePath As String = "conf/nrm_backup/nrmSpecfilePath"
    Dim actualDir As String
    Dim actualFile As String
    Dim compressLevel As String
    Dim compressType As String
    Dim compressExt As String
    Dim key As String = RandomString(10)
    Dim MyKey As String = "YOUR_KEY_HERE"
    Private Sub Backup_menu_load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim dt As Date = Today
        AllowTransparency = False
        Style.TitleBar.IconBackColor = Color.FromArgb(15, 161, 212)
        BackColor = Color.AliceBlue
        Style.TitleBar.TextHorizontalAlignment = HorizontalAlignment.Center
        Style.TitleBar.TextVerticalAlignment = VisualStyles.VerticalAlignment.Center
        Label1.Visible = False
        Label4.Visible = False
        Panel3.Visible = True
        Std_bck_pnl.Visible = False
        arc_bck_pnl.Visible = False
        arc_res_pnl.Visible = False
        TextBox1.ReadOnly = True
        TextBox2.ReadOnly = True
        TextBox3.ReadOnly = True
        TextBox6.ReadOnly = True
        SfDateTimeEdit1.Visible = False
        SfDateTimeEdit1.ShowUpDown = True
        SfDateTimeEdit1.DateTimeEditingMode = DateTimeEditingMode.Mask
        SfDateTimeEdit1.AllowNull = True
        SfDateTimeEdit1.Watermark = "Choose a date"
        SfDateTimeEdit2.Visible = False
        SfDateTimeEdit2.ShowUpDown = True
        SfDateTimeEdit2.DateTimeEditingMode = DateTimeEditingMode.Mask
        SfDateTimeEdit2.AllowNull = True
        SfDateTimeEdit2.Watermark = "Choose a date"
        WriteLogicalCount(uiProcessorCount)
        MessageBoxAdv.MessageBoxStyle = MessageBoxAdv.Style.Metro
    End Sub
    Private Sub Standard_Backup_Button(sender As Object, e As EventArgs) Handles Button1.Click
        Std_bck_pnl.Visible = True
        arc_bck_pnl.Visible = False
        arc_res_pnl.Visible = False
    End Sub
    Private Sub Archive_Backup_Button(sender As Object, e As EventArgs) Handles Button2.Click
        Std_bck_pnl.Visible = True
        arc_bck_pnl.Visible = True
        arc_res_pnl.Visible = False
    End Sub
    Private Sub Archive_Restore_Button(sender As Object, e As EventArgs) Handles Button3.Click
        Std_bck_pnl.Visible = True
        arc_bck_pnl.Visible = True
        arc_res_pnl.Visible = True
    End Sub
    Private Sub Close_Button(sender As Object, e As EventArgs) Handles Button4.Click
        Close()
    End Sub
    Private Sub Open_Folder_File_Std_Backup_Button(sender As Object, e As EventArgs) Handles Button5.Click
        If ComboBox2.Text = "Backup Folder" Then
            openfolderdialog.InitialDirectory = Environment.SpecialFolder.UserProfile
            If openfolderdialog.ShowDialog() = DialogResult.OK Then
                actualDir = openfolderdialog.SelectedPath
            End If
        ElseIf ComboBox2.Text = "Backup File" Then
            openfiledialog.InitialDirectory = Environment.SpecialFolder.UserProfile
            If openfiledialog.ShowDialog() = DialogResult.OK Then
                actualDir = Path.GetDirectoryName(openfiledialog.FileName)
                actualFile = openfiledialog.SafeFileName
            End If
        Else
            MessageBoxAdv.Show("Backup options was not selected !, Please select copy options first !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub
    Private Sub Save_Folder_Std_Backup_Button(sender As Object, e As EventArgs) Handles Button6.Click
        openfolderdialog.InitialDirectory = Environment.SpecialFolder.UserProfile
        openfolderdialog.ShowDialog()
    End Sub
    Private Sub Save_Std_Backup_Button(sender As Object, e As EventArgs) Handles Button7.Click
        If TextBox1.Text = "" Then
            MessageBoxAdv.Show("Source data is empty !, please fill source data location", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If TextBox2.Text = "" Then
                MessageBoxAdv.Show("Destination folder is empty !, Please fill destination folder location", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf ComboBox2.Text = "" Then
                MessageBoxAdv.Show("Backup options was not selected  !, Please select backup options first !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                Dim Dtcompare As Integer = DateTime.Compare(SfDateTimeEdit2.Value, SfDateTimeEdit1.Value)
                If Dtcompare < 0 Then
                    MessageBoxAdv.Show("To date can't be earlier than from date !, Please choose another date !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    If ComboBox2.Text = "Backup Folder" Then
                        ProgressBar1.Visible = True
                        ProgressBar1.Style = ProgressBarStyle.Marquee
                        ProgressBar1.MarqueeAnimationSpeed = 40
                        CheckFileExist(uiSpecFilePath, "*")
                        BeginCopy(ComboBox1.Text, ComboBox2.Text, TextBox1.Text, TextBox2.Text, SfDateTimeEdit1.Value, SfDateTimeEdit2.Value, actualDir, "")
                        ProgressBar1.Value = 100
                        ProgressBar1.Style = ProgressBarStyle.Blocks
                    ElseIf ComboBox2.Text = "Backup File" Then
                        ProgressBar1.Visible = True
                        ProgressBar1.Style = ProgressBarStyle.Marquee
                        ProgressBar1.MarqueeAnimationSpeed = 40
                        CheckFileExist(uiSpecFilePath, TextBox1.Text.ToString)
                        BeginCopy(ComboBox1.Text, ComboBox2.Text, TextBox1.Text, TextBox2.Text, SfDateTimeEdit1.Value, SfDateTimeEdit2.Value, actualDir, actualFile)
                        ProgressBar1.Value = 100
                        ProgressBar1.Style = ProgressBarStyle.Blocks
                    Else
                        MessageBoxAdv.Show("Backup options was not selected !, Please select backup options first !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                    ShowNotif("Backup", lastResult, lastErr)
                End If
            End If
        End If
    End Sub
    Private Sub Save_Archive_Backup_Button(sender As Object, e As EventArgs) Handles Button8.Click
        ProgressBar2.Visible = True
        ProgressBar2.Style = ProgressBarStyle.Marquee
        ProgressBar2.MarqueeAnimationSpeed = 40
        Dim uiTrimSrc As String
        Dim uiTrimDest As String
        If TextBox3.Text = "" Then
            MessageBoxAdv.Show("Source data is empty !, please fill source data location", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If TextBox6.Text = "" Then
                MessageBoxAdv.Show("Destination folder is empty !, Please fill destination folder location", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If ComboBox6.Text = "" Then
                    MessageBoxAdv.Show("Backup type is empty !, Please select backup type first !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ElseIf ComboBox6.Text = "Archive" Then
                    If ComboBox3.Text = "" Then
                        MessageBoxAdv.Show("Compression level is empty, please select compression type !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Else
                        InitComp()
                        If ComboBox4.Text = "" Then
                            MessageBoxAdv.Show("Compression type is empty, please select compression type !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Else
                            uiTrimSrc = TextBox3.Text
                            uiTrimDest = TextBox6.Text
                            compressType = CompType(ComboBox4.Text)
                            compressExt = CompExt(ComboBox4.Text)
                            compressLevel = CompLevel(ComboBox3.Text)
                            If Directory.Exists(uiTrimSrc) Then
                                Dim key As String = RandomString(10)
                                CheckFileExist(advEncType, "No Password")
                                CheckFileExist(advSrcPath, uiTrimSrc)
                                CheckFileExist(advDestPath, uiTrimDest)
                                CheckFileExist(advInstPath, My.Application.Info.DirectoryPath.ToString)
                                CheckFileExist(advRanStrg, key)
                                CheckFileExist(advCompExt, compressExt)
                                CheckFileExist(advCompType, compressType)
                                CheckFileExist(advCompLevel, compressLevel)
                                PrepareNotif(lastResult)
                                PrepareNotif(lastErr)
                                ManualBackup("bat\MigrateToGDrive_AR_NP.bat")
                            Else
                                CheckFileExist(lastResult, "err")
                                CheckFileExist(lastErr, "Source drive not exist !")
                            End If
                        End If
                        ShowNotif("Archive backup", lastResult, lastErr)
                    End If
                ElseIf ComboBox6.Text = "Archive + Password" Then
                    If ComboBox3.Text = "" Then
                        MessageBoxAdv.Show("Compression level is empty, please select compression type !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Else
                        If ComboBox4.Text = "" Then
                            MessageBoxAdv.Show("Compression type is empty, please select compression type !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Else
                            compressType = CompType(ComboBox4.Text)
                            If ComboBox5.Text = "" Then
                                MessageBoxAdv.Show("Password type is empty, please select password type !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Else
                                If Button14.Enabled = False Then
                                    InitComp()
                                    uiTrimSrc = TextBox3.Text
                                    uiTrimDest = TextBox6.Text
                                    compressType = CompType(ComboBox4.Text)
                                    compressExt = CompExt(ComboBox4.Text)
                                    compressLevel = CompLevel(ComboBox3.Text)
                                    Dim key As String = RandomString(10)
                                    If ComboBox5.SelectedIndex = 0 Then
                                        If Directory.Exists(uiTrimSrc) Then
                                            CheckFileExist(advEncType, "Password + No Encryption")
                                            CheckFileExist(advSrcPath, uiTrimSrc)
                                            CheckFileExist(advDestPath, uiTrimDest)
                                            CheckFileExist(advInstPath, My.Application.Info.DirectoryPath.ToString)
                                            CheckFileExist(advRanStrg, key)
                                            CheckFileExist(advCompExt, compressExt)
                                            CheckFileExist(advCompLevel, compressLevel)
                                            CheckFileExist(advCompType, compressType)
                                            CheckFileExist(advTempPass, TextBox8.Text)
                                            PrepareNotif(lastResult)
                                            PrepareNotif(lastErr)
                                            ManualBackup("bat\MigrateToGDrive_AR_P.bat")
                                        Else
                                            CheckFileExist(lastResult, "err")
                                            CheckFileExist(lastErr, "Source drive not exist !")
                                        End If
                                        PrepareNotif(advTempPass)
                                    ElseIf ComboBox5.SelectedIndex = 1 Then
                                        If Directory.Exists(uiTrimSrc) Then
                                            Dim sha_pass1 As String = AESEncryptStringToBase64(TextBox8.Text, MyKey)
                                            Dim sha_pass2 As String = AESEncryptStringToBase64(sha_pass1, MyKey)
                                            CheckFileExist(advEncType, "Password + SHA-256")
                                            CheckFileExist(advSrcPath, uiTrimSrc)
                                            CheckFileExist(advDestPath, uiTrimDest)
                                            CheckFileExist(advInstPath, My.Application.Info.DirectoryPath.ToString)
                                            CheckFileExist(advRanStrg, key)
                                            CheckFileExist(advCompExt, compressExt)
                                            CheckFileExist(advCompLevel, compressLevel)
                                            CheckFileExist(advCompType, compressType)
                                            CheckFileExist(advTempPass, sha_pass1)
                                            PrepareNotif(lastResult)
                                            PrepareNotif(lastErr)
                                            ManualBackup("bat\MigrateToGDrive_AR_PE.bat")
                                            CheckFileExist(advTempPass, sha_pass2)
                                            EncryptFile(TextBox8.Text, advTempPass, TextBox6.Text & "\Office_Tools" & key & "_KEY.ofk")
                                        Else
                                            CheckFileExist(lastResult, "err")
                                            CheckFileExist(lastErr, "Source drive not exist !")
                                        End If
                                        PrepareNotif(advTempPass)
                                    End If
                                    ShowNotif("Archive backup", lastResult, lastErr)
                                Else
                                    MessageBoxAdv.Show("Please save the password to proceed backup !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        ProgressBar2.Value = 100
        ProgressBar2.Style = ProgressBarStyle.Blocks
    End Sub
    Private Sub Extract_Restore_Backup_Button(sender As Object, e As EventArgs) Handles Button9.Click
        If Button20.Enabled = False Then
            MessageBoxAdv.Show("Password has been set, please cancel to change this option !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            openfolderdialog.InitialDirectory = Environment.SpecialFolder.UserProfile
            openfolderdialog.ShowDialog()
        End If
    End Sub
    Private Sub Destination_Archive_Backup_Button(sender As Object, e As EventArgs) Handles Button11.Click
        If ComboBox6.SelectedIndex = 0 Then
            openfolderdialog.InitialDirectory = Environment.SpecialFolder.UserProfile
            openfolderdialog.ShowDialog()
        Else
            If Button14.Enabled = False Then
                MessageBoxAdv.Show("Password has been set, please cancel to change this option !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
                openfolderdialog.InitialDirectory = Environment.SpecialFolder.UserProfile
                openfolderdialog.ShowDialog()
            End If
        End If
    End Sub
    Private Sub Source_Archive_Backup_Button(sender As Object, e As EventArgs) Handles Button12.Click
        If ComboBox6.SelectedIndex = 0 Then
            openfolderdialog.InitialDirectory = Environment.SpecialFolder.UserProfile
            openfolderdialog.ShowDialog()
        Else
            If Button14.Enabled = False Then
                MessageBoxAdv.Show("Password has been set, please cancel to change this option !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
                openfolderdialog.InitialDirectory = Environment.SpecialFolder.UserProfile
                openfolderdialog.ShowDialog()
            End If
        End If
    End Sub
    Private Sub Cancel_Archive_Pass_Button(sender As Object, e As EventArgs) Handles Button13.Click
        TextBox3.ReadOnly = False
        TextBox6.ReadOnly = False
        TextBox7.Text = ""
        TextBox7.ReadOnly = False
        TextBox8.Text = ""
        TextBox8.ReadOnly = False
        ComboBox3.Enabled = True
        ComboBox4.Enabled = True
        ComboBox5.Enabled = True
        ComboBox6.Enabled = True
        ComboBox3.ResetText()
        ComboBox4.ResetText()
        ComboBox5.ResetText()
        ComboBox6.ResetText()
        Button14.Enabled = True
    End Sub
    Private Sub Save_Archive_Pass_Button(sender As Object, e As EventArgs) Handles Button14.Click
        If TextBox8.Text = "" Then
            MessageBoxAdv.Show("Please fill your password !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            If TextBox7.Text = "" Then
                MessageBoxAdv.Show("Please complete your password !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                If TextBox8.Text = TextBox7.Text Then
                    TextBox3.ReadOnly = True
                    TextBox6.ReadOnly = True
                    TextBox7.ReadOnly = True
                    TextBox8.ReadOnly = True
                    Button14.Enabled = False
                    ComboBox3.Enabled = False
                    ComboBox4.Enabled = False
                    ComboBox5.Enabled = False
                    ComboBox6.Enabled = False
                    MessageBoxAdv.Show("Password locked !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBoxAdv.Show("Password not match !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        End If
    End Sub
    Private Sub Enc_Key_Restore_Backup_Button(sender As Object, e As EventArgs) Handles Button16.Click
        If Button20.Enabled = False Then
            MessageBoxAdv.Show("Password has been set, please cancel to change this option !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            If ComboBox7.Text = "Password (No Encryption)" Or ComboBox7.Text = "" Then
                MessageBoxAdv.Show("Encryption method set as no encryption, no key file needed !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
                openfiledialog.InitialDirectory = Environment.SpecialFolder.UserProfile
                openfiledialog.DefaultExt = ".ofk;.mtg"
                openfiledialog.Filter = "Office Tools Encrypted Key|*.ofk;*.mtg"
                openfiledialog.ShowDialog()
            End If
        End If
    End Sub
    Private Sub Archive_File_Restore_Backup(sender As Object, e As EventArgs) Handles Button18.Click
        If Button20.Enabled = False Then
            MessageBoxAdv.Show("Password has been set, please cancel to change this option !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            openfiledialog.InitialDirectory = Environment.SpecialFolder.UserProfile
            openfiledialog.DefaultExt = ".7z"
            openfiledialog.Filter = "7-ZIP Supported Format|*.7z;*.zip"
            openfiledialog.ShowDialog()
        End If
    End Sub
    Private Sub Cancel_Pass_Restore_Backup(sender As Object, e As EventArgs) Handles Button19.Click
        TextBox11.ReadOnly = False
        TextBox4.ReadOnly = False
        TextBox5.ReadOnly = False
        TextBox13.Text = ""
        TextBox13.ReadOnly = False
        TextBox12.ReadOnly = False
        TextBox12.Text = ""
        ComboBox7.Enabled = True
        ComboBox7.ResetText()
        Button20.Enabled = True
    End Sub
    Private Sub Save_Pass_Restore_Backup(sender As Object, e As EventArgs) Handles Button20.Click
        If TextBox13.Text = "" Then
            MessageBoxAdv.Show("Please fill your password !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            If TextBox12.Text = "" Then
                MessageBoxAdv.Show("Please complete your password !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                If TextBox13.Text = TextBox12.Text Then
                    TextBox11.ReadOnly = True
                    TextBox4.ReadOnly = True
                    TextBox5.ReadOnly = True
                    TextBox13.ReadOnly = True
                    TextBox12.ReadOnly = True
                    ComboBox7.Enabled = False
                    Button20.Enabled = False
                    MessageBoxAdv.Show("Password locked !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBoxAdv.Show("Password not match !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        End If
    End Sub
    Private Sub Restore_Archive_Button(sender As Object, e As EventArgs) Handles Button21.Click
        ProgressBar3.Visible = True
        ProgressBar3.Style = ProgressBarStyle.Marquee
        ProgressBar3.MarqueeAnimationSpeed = 40
        Dim srcFile As String
        Dim destFolder As String
        Dim encKeyPath As String
        Dim encPass As String
        Dim decKey As String
        Dim decIndKey As Integer
        Dim decKeyPass As String
        Dim decAesKey As String
        If TextBox11.Text = "" Then
            MessageBoxAdv.Show("Archive file path is empty !, please select file", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If TextBox4.Text = "" Then
                MessageBoxAdv.Show("Destination folder is empty !, please fill destination folder location", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                srcFile = TextBox11.Text
                destFolder = TextBox4.Text
                encKeyPath = TextBox5.Text
                encPass = TextBox13.Text
                InitComp()
                If ComboBox7.SelectedIndex = 0 Then
                    If File.Exists(srcFile) Then
                        If Directory.Exists(destFolder) Then
                            CheckFileExist(resSrcPath, srcFile)
                            CheckFileExist(resDestPath, destFolder)
                            CheckFileExist(resInstPath, My.Application.Info.DirectoryPath.ToString)
                            CheckFileExist(resDecMtd, "No Decryption")
                            CheckFileExist(resTempPass, TextBox8.Text)
                            PrepareNotif(lastResult)
                            PrepareNotif(lastErr)
                            ManualBackup("bat\MigrateToGDrive_RB.bat")
                            decKey = File.ReadAllText(resZipLog)
                            decIndKey = decKey.IndexOf("Archives with Errors: 1")
                            If decIndKey >= 0 Then
                                PrepareNotif(lastResult)
                                PrepareNotif(lastErr)
                                CheckFileExist(lastResult, "err")
                                CheckFileExist(lastErr, "Extract error, password not match !")
                                Dim writer As New StreamWriter(resLogPath, True)
                                writer.WriteLine("# Office Tools v1.2")
                                writer.WriteLine("Extract Result               : Error")
                                writer.WriteLine("Extract Time                 : " & DateTime.Now)
                                writer.WriteLine("Extract Filename          : " & TextBox11.Text)
                                writer.WriteLine("Extract Location          : " & TextBox4.Text)
                                writer.WriteLine("Error Reason                 : Password Error !")
                                writer.WriteLine("Decryption Method    : No Decryption")
                                writer.WriteLine(vbCrLf)
                                writer.Close()
                            Else
                                decIndKey = decKey.IndexOf("Everything is Ok")
                                If decIndKey >= 0 Then
                                    PrepareNotif(lastResult)
                                    PrepareNotif(lastErr)
                                    CheckFileExist(lastResult, "success")
                                    Dim writer As New StreamWriter(resLogPath, True)
                                    writer.WriteLine("# Office Tools v1.2")
                                    writer.WriteLine("Extract Result               : Success")
                                    writer.WriteLine("Extract Time                 : " & DateTime.Now)
                                    writer.WriteLine("Extract Filename          : " & TextBox11.Text)
                                    writer.WriteLine("Extract Location          : " & TextBox4.Text)
                                    writer.WriteLine("Decryption Method    : No Decryption")
                                    writer.WriteLine(vbCrLf)
                                    writer.Close()
                                End If
                            End If
                        Else
                            CheckFileExist(lastResult, "err")
                            CheckFileExist(lastErr, "Destination folder not exist !")
                        End If
                    Else
                        CheckFileExist(lastResult, "err")
                        CheckFileExist(lastErr, "Source file not exist !")
                    End If
                Else
                    If TextBox5.Text = "" Then
                        MessageBoxAdv.Show("Encryption key path is empty, please select encryption key file !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Else
                        If Button20.Visible = True Then
                            If ComboBox7.SelectedIndex = 1 Then
                                DecryptFile(encPass, encKeyPath, encKeyPath.Remove(0, encKeyPath.Length - 3))
                                If PathVal(resDecResult, 0).Equals("KEY_ERR") Then
                                    CheckFileExist(lastResult, "err")
                                    CheckFileExist(lastErr, "Key decryption failed, password not match !")
                                Else
                                    decKeyPass = PathVal(encKeyPath.Remove(0, encKeyPath.Length - 3), 0)
                                    If AESDecryptBase64ToString(decKeyPass, MyKey).Equals("SHA256_ERR") Then
                                        CheckFileExist(lastResult, "err")
                                        CheckFileExist(lastErr, "SHA256 decryption failure !")
                                    Else
                                        decAesKey = AESDecryptBase64ToString(decKeyPass, MyKey)
                                        If File.Exists(srcFile) Then
                                            If Directory.Exists(destFolder) Then
                                                CheckFileExist(resSrcPath, srcFile)
                                                CheckFileExist(resDestPath, destFolder)
                                                CheckFileExist(resInstPath, My.Application.Info.DirectoryPath.ToString)
                                                CheckFileExist(resTempPass, decAesKey)
                                                CheckFileExist(resDecMtd, "SHA-256")
                                                PrepareNotif(lastResult)
                                                PrepareNotif(lastErr)
                                                ManualBackup("bat\MigrateToGDrive_RB.bat")
                                                File.Delete(destFolder & "\" & encKeyPath.Remove(0, encKeyPath.Length - 3))
                                                File.Delete(My.Application.Info.DirectoryPath.ToString & "\" & encKeyPath.Remove(0, encKeyPath.Length - 3))
                                                decKey = File.ReadAllText(resZipLog)
                                                decIndKey = decKey.IndexOf("Archives with Errors: 1")
                                                If decIndKey >= 0 Then
                                                    PrepareNotif(lastResult)
                                                    PrepareNotif(lastErr)
                                                    CheckFileExist(lastResult, "err")
                                                    CheckFileExist(lastErr, "Extract error, password not match !")
                                                    Dim writer As New StreamWriter(resErrPath, True)
                                                    writer.WriteLine("# Office Tools v1.2")
                                                    writer.WriteLine("Extract Result               : Error")
                                                    writer.WriteLine("Extract Time                 : " & DateTime.Now)
                                                    writer.WriteLine("Extract Filename          : " & TextBox11.Text)
                                                    writer.WriteLine("Extract Location          : " & TextBox4.Text)
                                                    writer.WriteLine("Error Reason                 : Password Error !")
                                                    writer.WriteLine("Decryption Method    : SHA-256")
                                                    writer.WriteLine(vbCrLf)
                                                    writer.Close()
                                                Else
                                                    decIndKey = decKey.IndexOf("Everything is Ok")
                                                    If decIndKey >= 0 Then
                                                        PrepareNotif(lastResult)
                                                        PrepareNotif(lastErr)
                                                        CheckFileExist(lastResult, "success")
                                                        Dim writer As New StreamWriter(resLogPath, True)
                                                        writer.WriteLine("# Office Tools v1.2")
                                                        writer.WriteLine("Extract Result               : Success")
                                                        writer.WriteLine("Extract Time                 : " & DateTime.Now)
                                                        writer.WriteLine("Extract Filename          : " & TextBox11.Text)
                                                        writer.WriteLine("Extract Location          : " & TextBox4.Text)
                                                        writer.WriteLine("Decryption Method    : SHA-256")
                                                        writer.WriteLine(vbCrLf)
                                                        writer.Close()
                                                    End If
                                                End If
                                            Else
                                                CheckFileExist(lastResult, "err")
                                                CheckFileExist(lastErr, "Source drive not exist !")
                                            End If
                                        Else
                                            CheckFileExist(lastResult, "err")
                                            CheckFileExist(lastErr, "Source file not exist !")
                                        End If
                                        PrepareNotif(resTempPass)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                ShowNotif("Restore archive", lastResult, lastErr)
            End If
        End If
        ProgressBar3.Value = 100
        ProgressBar3.Style = ProgressBarStyle.Blocks
    End Sub
    Private Sub Open_Folder_File_Std_Backup_Handler(sender As Object, e As EventArgs) Handles Button5.Click
        If ComboBox2.Text = "Backup Folder" Then
            TextBox1.Text = openfolderdialog.SelectedPath.ToString
        ElseIf ComboBox2.Text = "Backup File" Then
            If openfiledialog.FileName.ToString = "" Then
                TextBox1.Text = ""
            Else
                TextBox1.Text = Path.GetFullPath(openfiledialog.FileName.ToString)
            End If
        End If
    End Sub
    Private Sub Save_Folder_Std_Backup_Handler(sender As Object, e As EventArgs) Handles Button6.Click
        TextBox2.Text = openfolderdialog.SelectedPath.ToString
    End Sub
    Private Sub Extract_Restore_Backup_Handler(sender As Object, e As EventArgs) Handles Button9.Click
        TextBox4.Text = openfolderdialog.SelectedPath.ToString
    End Sub
    Private Sub Destination_Archive_Backup_Handler(sender As Object, e As EventArgs) Handles Button11.Click
        If ComboBox6.SelectedIndex = 0 Then
            TextBox6.Text = openfolderdialog.SelectedPath.ToString
        Else
            If Button4.Enabled = True Then
                TextBox6.Text = openfolderdialog.SelectedPath.ToString
            End If
        End If
    End Sub
    Private Sub Source_Archive_Backup_Handler(sender As Object, e As EventArgs) Handles Button12.Click
        If ComboBox6.SelectedIndex = 0 Then
            TextBox3.Text = openfolderdialog.SelectedPath.ToString
        Else
            If Button4.Enabled = True Then
                TextBox3.Text = openfolderdialog.SelectedPath.ToString
            End If
        End If
    End Sub
    Private Sub Enc_Key_Restore_Backup_Handler(sender As Object, e As EventArgs) Handles Button16.Click
        If ComboBox7.SelectedIndex = 1 Then
            If openfiledialog.FileName.ToString.Remove(0, openfiledialog.FileName.ToString.Length - 3).Equals("ofk") Or openfiledialog.FileName.ToString.Remove(0, openfiledialog.FileName.ToString.Length - 3).Equals("mtg") Then
                TextBox5.Text = openfiledialog.FileName.ToString
            Else
                MessageBoxAdv.Show("Please select a valid Office Tools encrypted key file !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub
    Private Sub Archive_File_Restore_Handler(sender As Object, e As EventArgs) Handles Button18.Click
        TextBox11.Text = openfiledialog.FileName.ToString
    End Sub
    Private Sub Backup_Period_Handler(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "From Date" Then
            SfDateTimeEdit1.Visible = True
            SfDateTimeEdit2.Visible = True
            Label1.Visible = True
            Label4.Visible = True
        Else
            SfDateTimeEdit1.Visible = False
            SfDateTimeEdit2.Visible = False
            Label1.Visible = False
            Label4.Visible = False
        End If
    End Sub
    Private Sub Backup_Method_Handler(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        If ComboBox6.Text = "Archive" Then
            Label20.Visible = False
            Label19.Visible = False
            Button13.Visible = False
            Button14.Visible = False
            TextBox8.Visible = False
            TextBox7.Visible = False
            ComboBox5.Enabled = False
        Else
            Label20.Visible = True
            Label19.Visible = True
            Button13.Visible = True
            Button14.Visible = True
            TextBox8.Visible = True
            TextBox7.Visible = True
            ComboBox5.Enabled = True
        End If
    End Sub
    Private Sub Enc_Key_Method_Handler(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        If Button20.Enabled = False Then
            MessageBoxAdv.Show("Password has been set, please cancel to change this option !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            If ComboBox7.SelectedIndex = 0 Then
                TextBox5.Enabled = False
                TextBox5.Text = ""
                Button16.Enabled = False
                TextBox5.ReadOnly = True
            Else
                TextBox5.Enabled = True
                TextBox5.Text = ""
                Button16.Enabled = True
                TextBox5.ReadOnly = False
            End If
        End If
    End Sub
    Private Sub InitComp()
        PrepareNotif(advCompExt)
        PrepareNotif(advCompLevel)
        PrepareNotif(advCompType)
        PrepareNotif(advDestPath)
        PrepareNotif(advEncType)
        PrepareNotif(advRanStrg)
        PrepareNotif(advSrcPath)
        PrepareNotif(advTempPass)
        PrepareNotif(resDestPath)
        PrepareNotif(resInstPath)
        PrepareNotif(resSrcPath)
        PrepareNotif(resTempPass)
        PrepareNotif(resDecResult)
        PrepareNotif(resZipLog)
        PrepareNotif(resDecMtd)
    End Sub
End Class
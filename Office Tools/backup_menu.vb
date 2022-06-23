Imports System.IO
Public Class backup_menu
    Dim filedialog As New FolderBrowserDialog
    Dim openfiledialog As New OpenFileDialog
    Dim openfolderdialog As New FolderBrowserDialog
    Dim savefiledialog As New SaveFileDialog
    Dim logPath As String = "log/reslog"
    Dim errPath As String = "log/reserr"
    Dim lastResult As String = "log/lastResult"
    Dim lastErr As String = "log/lastErr"
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
    Dim roboPath As String = "log/robolog"
    Dim uiSrcPath As String = "conf/nrm_backup/nrmSrcPath"
    Dim uiDestPath As String = "conf/nrm_backup/nrmDestPath"
    Dim uiFrDatePath As String = "conf/nrm_backup/nrmFrDatePath"
    Dim uiReDatePath As String = "conf/nrm_backup/nrmReDatePath"
    Dim uiToDatePath As String = "conf/nrm_backup/nrmToDatePath"
    Dim uiProcessorCount As String = "conf/nrm_backup/nrmProcessor"
    Dim uiSpecFilePath As String = "conf/nrm_backup/nrmSpecfilePath"
    Dim actualDir As String
    Dim compressLevel As String
    Dim compressType As String
    Dim compressExt As String
    Dim key As String = RandomString(10)
    Dim MyKey As String = "YOUR_KEY_HERE"
    Private Sub backup_menu_load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim dt As Date = Today
        AllowTransparency = False
        Label4.Visible = False
        Std_bck_pnl.Visible = False
        arc_bck_pnl.Visible = False
        arc_res_pnl.Visible = False
        TextBox1.ReadOnly = True
        TextBox2.ReadOnly = True
        TextBox3.ReadOnly = True
        TextBox6.ReadOnly = True
        DateTimePicker1.Visible = False
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "MM-dd-yyyy"
        DateTimePicker2.Visible = False
        DateTimePicker2.Format = DateTimePickerFormat.Custom
        DateTimePicker2.CustomFormat = "MM-dd-yyyy"
        WriteLogicalCount(uiProcessorCount)
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Std_bck_pnl.Visible = True
        arc_bck_pnl.Visible = False
        arc_res_pnl.Visible = False
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Std_bck_pnl.Visible = True
        arc_bck_pnl.Visible = True
        arc_res_pnl.Visible = False
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Std_bck_pnl.Visible = True
        arc_bck_pnl.Visible = True
        arc_res_pnl.Visible = True
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If ComboBox2.Text = "Backup Folder" Then
            filedialog.InitialDirectory = Environment.SpecialFolder.UserProfile
            filedialog.ShowDialog()
        ElseIf ComboBox2.Text = "Backup File" Then
            openfiledialog.InitialDirectory = Environment.SpecialFolder.UserProfile
            openfiledialog.ShowDialog()
        Else
            MsgBox("Backup options was not selected !, Please select copy options first !", MsgBoxStyle.Critical, "Office Tools")
        End If
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        filedialog.InitialDirectory = Environment.SpecialFolder.UserProfile
        filedialog.ShowDialog()
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If TextBox1.Text = "" Then
            MsgBox("Source data is empty !, please fill source data location", MsgBoxStyle.Critical, "Office Tools")
        Else
            If TextBox2.Text = "" Then
                MsgBox("Destination folder is empty !, Please fill destination folder location", MsgBoxStyle.Critical, "Office Tools")
            ElseIf ComboBox2.Text = "" Then
                MsgBox("Backup options was not selected  !, Please select backup options first !", MsgBoxStyle.Critical, "Office Tools")
            Else
                If ComboBox2.Text = "Backup Folder" Then
                    ProgressBar1.Visible = True
                    ProgressBar1.Style = ProgressBarStyle.Marquee
                    ProgressBar1.MarqueeAnimationSpeed = 40
                    CheckFileExist(uiSpecFilePath, "*")
                    beginCopy()
                    ProgressBar1.Value = 100
                    ProgressBar1.Style = ProgressBarStyle.Blocks
                ElseIf ComboBox2.Text = "Backup File" Then
                    ProgressBar1.Visible = True
                    ProgressBar1.Style = ProgressBarStyle.Marquee
                    ProgressBar1.MarqueeAnimationSpeed = 40
                    CheckFileExist(uiSpecFilePath, TextBox1.Text.ToString)
                    beginCopy()
                    ProgressBar1.Value = 100
                    ProgressBar1.Style = ProgressBarStyle.Blocks
                Else
                    MsgBox("Backup options was not selected !, Please select backup options first !", MsgBoxStyle.Critical, "Office Tools")
                End If
                ShowNotif("Backup")
            End If
        End If
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        ProgressBar2.Visible = True
        ProgressBar2.Style = ProgressBarStyle.Marquee
        ProgressBar2.MarqueeAnimationSpeed = 40
        Dim uiTrimSrc As String
        Dim uiTrimDest As String
        If TextBox3.Text = "" Then
            MsgBox("Source data is empty !, please fill source data location", MsgBoxStyle.Critical, "Office Tools")
        Else
            If TextBox6.Text = "" Then
                MsgBox("Destination folder is empty !, Please fill destination folder location", MsgBoxStyle.Critical, "Office Tools")
            Else
                If ComboBox6.Text = "" Then
                    MsgBox("Backup type is empty !, Please select backup type first !", MsgBoxStyle.Critical, "Office Tools")
                ElseIf ComboBox6.Text = "Archive" Then
                    If ComboBox3.Text = "" Then
                        MsgBox("Compression level is empty, please select compression type !", MsgBoxStyle.Critical, "Office Tools")
                    Else
                        InitComp()
                        If ComboBox4.Text = "" Then
                            MsgBox("Compression type is empty, please select compression type !", MsgBoxStyle.Critical, "Office Tools")
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
                        ShowNotif("Archive backup")
                    End If
                ElseIf ComboBox6.Text = "Archive + Password" Then
                    If ComboBox3.Text = "" Then
                        MsgBox("Compression level is empty, please select compression type !", MsgBoxStyle.Critical, "Office Tools")
                    Else
                        If ComboBox4.Text = "" Then
                            MsgBox("Compression type is empty, please select compression type !", MsgBoxStyle.Critical, "Office Tools")
                        Else
                            compressType = CompType(ComboBox4.Text)
                            If ComboBox5.Text = "" Then
                                MsgBox("Password type is empty, please select password type !", MsgBoxStyle.Critical, "Office Tools")
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
                                    ShowNotif("Archive backup")
                                Else
                                    MsgBox("Please save the password to proceed backup !", vbInformation, "Office Tools")
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
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If Button20.Enabled = False Then
            MsgBox("Password has been set, please cancel to change this option !", vbExclamation, "Office Tools")
        Else
            openfolderdialog.InitialDirectory = Environment.SpecialFolder.UserProfile
            openfolderdialog.ShowDialog()
        End If
    End Sub
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If ComboBox6.SelectedIndex = 0 Then
            filedialog.InitialDirectory = Environment.SpecialFolder.UserProfile
            filedialog.ShowDialog()
        Else
            If Button14.Enabled = False Then
                MsgBox("Password has been set, please cancel to change this option !", vbExclamation, "Office Tools")
            Else
                filedialog.InitialDirectory = Environment.SpecialFolder.UserProfile
                filedialog.ShowDialog()
            End If
        End If
    End Sub
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If ComboBox6.SelectedIndex = 0 Then
            filedialog.InitialDirectory = Environment.SpecialFolder.UserProfile
            filedialog.ShowDialog()
        Else
            If Button14.Enabled = False Then
                MsgBox("Password has been set, please cancel to change this option !", vbExclamation, "Office Tools")
            Else
                filedialog.InitialDirectory = Environment.SpecialFolder.UserProfile
                filedialog.ShowDialog()
            End If
        End If
    End Sub
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
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
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        If TextBox8.Text = "" Then
            MsgBox("Please fill your password !", vbInformation, "Office Tools")
        Else
            If TextBox7.Text = "" Then
                MsgBox("Please complete your password !", vbInformation, "Office Tools")
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
                    MsgBox("Password locked !", vbInformation, "Office Tools")
                Else
                    MsgBox("Password not match !", vbExclamation, "Office Tools")
                End If
            End If
        End If
    End Sub
    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        If Button20.Enabled = False Then
            MsgBox("Password has been set, please cancel to change this option !", vbExclamation, "Office Tools")
        Else
            If ComboBox7.SelectedIndex = 0 Then
                MsgBox("Encryption method set as no encryption, no key file needed !", vbExclamation, "Office Tools")
            Else
                openfiledialog.InitialDirectory = Environment.SpecialFolder.UserProfile
                openfiledialog.DefaultExt = ".ofk"
                openfiledialog.Filter = "Office Tools Encrypted Key|*.ofk"
                openfiledialog.ShowDialog()
            End If
        End If
    End Sub
    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        If Button20.Enabled = False Then
            MsgBox("Password has been set, please cancel to change this option !", vbExclamation, "Office Tools")
        Else
            openfiledialog.InitialDirectory = Environment.SpecialFolder.UserProfile
            openfiledialog.DefaultExt = ".7z"
            openfiledialog.Filter = "7-ZIP Supported Format|*.7z;*.zip"
            openfiledialog.ShowDialog()
        End If
    End Sub
    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
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
    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        If TextBox13.Text = "" Then
            MsgBox("Please fill your password !", vbExclamation, "Office Tools")
        Else
            If TextBox12.Text = "" Then
                MsgBox("Please fill your password !", vbExclamation, "Office Tools")
            Else
                If TextBox13.Text = TextBox12.Text Then
                    TextBox11.ReadOnly = True
                    TextBox4.ReadOnly = True
                    TextBox5.ReadOnly = True
                    TextBox13.ReadOnly = True
                    TextBox12.ReadOnly = True
                    ComboBox7.Enabled = False
                    Button20.Enabled = False
                    MsgBox("Password locked !", vbInformation, "Office Tools")
                Else
                    MsgBox("Password not match !", vbCritical, "Office Tools")
                End If
            End If
        End If
    End Sub
    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
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
            MsgBox("Archive file path is empty !, please select file", MsgBoxStyle.Critical, "Office Tools")
        Else
            If TextBox4.Text = "" Then
                MsgBox("Destination folder is empty !, please fill destination folder location", MsgBoxStyle.Critical, "Office Tools")
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
                                Dim writer As New StreamWriter(errPath, True)
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
                                    Dim writer As New StreamWriter(logPath, True)
                                    writer.WriteLine("# Office Tools v1.1")
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
                        MsgBox("Encryption key path is empty, please select encryption key file !", MsgBoxStyle.Critical, "Office Tools")
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
                                                decKey = File.ReadAllText(resZipLog)
                                                decIndKey = decKey.IndexOf("Archives with Errors: 1")
                                                If decIndKey >= 0 Then
                                                    PrepareNotif(lastResult)
                                                    PrepareNotif(lastErr)
                                                    CheckFileExist(lastResult, "err")
                                                    CheckFileExist(lastErr, "Extract error, password not match !")
                                                    Dim writer As New StreamWriter(errPath, True)
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
                                                        Dim writer As New StreamWriter(logPath, True)
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
                ShowNotif("Restore archive")
            End If
        End If
        ProgressBar3.Value = 100
        ProgressBar3.Style = ProgressBarStyle.Blocks
    End Sub
    Private Sub FolderBrowserDialog1_Disposed(sender As Object, e As EventArgs) Handles Button5.Click
        If ComboBox2.Text = "Backup Folder" Then
            TextBox1.Text = filedialog.SelectedPath.ToString
            Label7.Text = filedialog.SelectedPath.ToString
        ElseIf ComboBox2.Text = "Backup File" Then
            TextBox1.Text = Path.GetFileName(openfiledialog.FileName.ToString)
            If (openfiledialog.FileName.ToString) = "" Then
                Label7.Text = ""
            Else
                actualDir = Path.GetFullPath(openfiledialog.FileName.ToString)
                Label7.Text = Path.GetDirectoryName(actualDir)
            End If
        End If
    End Sub
    Private Sub FolderBrowserDialog2_Disposed(sender As Object, e As EventArgs) Handles Button6.Click
        TextBox2.Text = filedialog.SelectedPath.ToString
    End Sub
    Private Sub FolderBrowserDialog9_Disposed(sender As Object, e As EventArgs) Handles Button9.Click
        TextBox4.Text = openfolderdialog.SelectedPath.ToString
    End Sub
    Private Sub FolderBrowserDialog11_Disposed(sender As Object, e As EventArgs) Handles Button11.Click
        If ComboBox6.SelectedIndex = 0 Then
            TextBox6.Text = filedialog.SelectedPath.ToString
        Else
            If Button4.Enabled = True Then
                TextBox6.Text = filedialog.SelectedPath.ToString
            End If
        End If
    End Sub
    Private Sub FolderBrowserDialog12_Disposed(sender As Object, e As EventArgs) Handles Button12.Click
        If ComboBox6.SelectedIndex = 0 Then
            TextBox3.Text = filedialog.SelectedPath.ToString
        Else
            If Button4.Enabled = True Then
                TextBox3.Text = filedialog.SelectedPath.ToString
            End If
        End If
    End Sub
    Private Sub FileSaveDialog16_Disposed(sender As Object, e As EventArgs) Handles Button16.Click
        If ComboBox7.SelectedIndex = 1 Then
            If openfiledialog.FileName.ToString.Remove(0, openfiledialog.FileName.ToString.Length - 3).Equals("ofk") Then
                TextBox5.Text = openfiledialog.FileName.ToString
            Else
                MsgBox("Please select a valid Office Tools encrypted key file !", vbCritical, "Office Tools")
            End If
        End If
    End Sub
    Private Sub FileBrowserDialog18_Disposed(sender As Object, e As EventArgs) Handles Button18.Click
        TextBox11.Text = openfiledialog.FileName.ToString
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "From Date" Then
            DateTimePicker1.Visible = True
            DateTimePicker2.Visible = True
            Label1.Visible = True
        Else
            DateTimePicker1.Visible = False
            DateTimePicker2.Visible = False
            Label1.Visible = False
        End If
    End Sub
    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
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
    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        If Button20.Enabled = False Then
            MsgBox("Password has been set, please cancel to change this option !", vbExclamation, "Office Tools")
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
    Private Sub beginCopy()
        If File.Exists(roboPath) Then
            PrepareNotif(roboPath)
        End If
        If ComboBox1.Text = "Anytime" Then
            Dim uiTrimSrc As String
            Dim uiTrimDest As String
            uiTrimSrc = Label7.Text
            uiTrimDest = TextBox2.Text
            If Directory.Exists(uiTrimSrc) Then
                If Directory.Exists(uiTrimDest) Then
                    CheckFileExist(uiSrcPath, uiTrimSrc)
                    CheckFileExist(uiDestPath, uiTrimDest)
                    PrepareNotif(lastResult)
                    PrepareNotif(lastErr)
                    ManualBackup("bat/MigrateToGDrive_AT_MN.bat")
                    WriteFrRobo()
                Else
                    CheckFileExist(lastResult, "err")
                    CheckFileExist(lastErr, "Destination folder not exist !")
                End If
            Else
                CheckFileExist(lastResult, "err")
                CheckFileExist(lastErr, "Source folder not exist !")
            End If
        ElseIf ComboBox1.Text = "Today" Then
            If File.Exists(uiReDatePath) Then
                GC.Collect()
                GC.WaitForPendingFinalizers()
                File.Delete(uiReDatePath)
                File.Create(uiReDatePath).Dispose()
                Dim destWriter As New StreamWriter(uiReDatePath, True)
                Dim dt As Date = Today
                destWriter.WriteLine(dt.ToString("yyyyMMdd"))
                destWriter.Close()
            Else
                File.Create(uiReDatePath).Dispose()
                Dim destWriter As New StreamWriter(uiReDatePath, True)
                Dim dt As Date = Today
                destWriter.WriteLine(dt.ToString("yyyyMMdd"))
                destWriter.Close()
            End If
            Dim uiTrimSrc As String
            Dim uiTrimDest As String
            uiTrimSrc = Label7.Text
            uiTrimDest = TextBox2.Text
            If Directory.Exists(uiTrimSrc) Then
                If Directory.Exists(uiTrimDest) Then
                    CheckFileExist(uiSrcPath, uiTrimSrc)
                    CheckFileExist(uiDestPath, uiTrimDest)
                    PrepareNotif(lastResult)
                    PrepareNotif(lastErr)
                    ManualBackup("bat/MigrateToGDrive_TD_MN.bat")
                    WriteFrRobo()
                Else
                    CheckFileExist(lastResult, "err")
                    CheckFileExist(lastErr, "Destination folder not exist !")
                End If
            Else
                CheckFileExist(lastResult, "err")
                CheckFileExist(lastErr, "Source folder not exist !")
            End If
        ElseIf ComboBox1.Text = "From Date" Then
            If File.Exists(uiFrDatePath) Then
                GC.Collect()
                GC.WaitForPendingFinalizers()
                File.Delete(uiFrDatePath)
                File.Create(uiFrDatePath).Dispose()
                Dim destWriter As New StreamWriter(uiFrDatePath, True)
                Dim dt As Date = DateTimePicker1.Value.ToShortDateString
                destWriter.WriteLine(dt.ToString("yyyyMMdd"))
                destWriter.Close()
            Else
                File.Create(uiFrDatePath).Dispose()
                Dim destWriter As New StreamWriter(uiFrDatePath, True)
                Dim dt As Date = DateTimePicker1.Value.ToShortDateString
                destWriter.WriteLine(dt.ToString("yyyyMMdd"))
                destWriter.Close()
            End If
            If File.Exists(uiToDatePath) Then
                GC.Collect()
                GC.WaitForPendingFinalizers()
                File.Delete(uiToDatePath)
                File.Create(uiToDatePath).Dispose()
                Dim destWriter As New StreamWriter(uiToDatePath, True)
                Dim dt As Date = DateTimePicker2.Value.ToShortDateString
                Dim newDate As Integer = Integer.Parse(dt.ToString("dd") + 1)
                If newDate < 10 Then
                    Dim newAffixDate = "0" + newDate.ToString
                    Dim newMonthYear As String = dt.ToString("yyyyMM")
                    destWriter.WriteLine(newMonthYear + newAffixDate.ToString)
                Else
                    Dim newMonthYear As String = dt.ToString("yyyyMM")
                    destWriter.WriteLine(newMonthYear + newDate.ToString)
                End If
                destWriter.Close()
            Else
                File.Create(uiToDatePath).Dispose()
                Dim destWriter As New StreamWriter(uiToDatePath, True)
                Dim dt As Date = DateTimePicker2.Value.ToShortDateString
                Dim newDate As Integer = Integer.Parse(dt.ToString("dd") + 1)
                If newDate < 10 Then
                    Dim newAffixDate = "0" + newDate.ToString
                    Dim newMonthYear As String = dt.ToString("yyyyMM")
                    destWriter.WriteLine(newMonthYear + newAffixDate.ToString)
                Else
                    Dim newMonthYear As String = dt.ToString("yyyyMM")
                    destWriter.WriteLine(newMonthYear + newDate.ToString)
                End If
                destWriter.Close()
            End If
            Dim uiTrimSrc As String
            Dim uiTrimDest As String
            uiTrimSrc = Label7.Text
            uiTrimDest = TextBox2.Text
            If Directory.Exists(uiTrimSrc) Then
                If Directory.Exists(uiTrimDest) Then
                    CheckFileExist(uiSrcPath, uiTrimSrc)
                    CheckFileExist(uiDestPath, uiTrimDest)
                    PrepareNotif(lastResult)
                    PrepareNotif(lastErr)
                    ManualBackup("bat/MigrateToGDrive_FD_MN.bat")
                    WriteFrRobo()
                Else
                    CheckFileExist(lastResult, "err")
                    CheckFileExist(lastErr, "Destination folder not exist !")
                End If
            Else
                CheckFileExist(lastResult, "err")
                CheckFileExist(lastErr, "Source folder not exist !")
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
    Private Sub ShowNotif(result As String)
        If File.Exists(lastResult) Then
            Dim lastRest As String = PathVal(lastResult, 0)
            If PathVal(lastResult, 0).Equals("success") Then
                MsgBox(result & " success !", MsgBoxStyle.Information, "Office Tools")
            ElseIf PathVal(lastResult, 0).Equals("err") Then
                MsgBox(result & " error !", MsgBoxStyle.Critical, "Office Tools")
                If File.Exists(lastErr) Then
                    If PathVal(lastErr, 0).Equals("") Then
                        MsgBox("Unknown error reason !", MsgBoxStyle.Critical, "Office Tools")
                    Else
                        MsgBox(PathVal(lastErr, 0), MsgBoxStyle.Critical, "Office Tools")
                    End If
                Else
                    MsgBox("Error file not found !", MsgBoxStyle.Critical, "Office Tools")
                End If
            Else
                MsgBox("Unknown result status !", MsgBoxStyle.Critical, "Office Tools")
            End If
        Else
            MsgBox("Result file not found !", MsgBoxStyle.Critical, "Office Tools")
        End If
    End Sub
    Private Sub ShowLog(log As String, path As String)
        If File.Exists(path) Then
            If New FileInfo(path).Length.Equals(0) Then
                MsgBox(log & " File is empty !", MsgBoxStyle.Critical, "Office Tools")
            End If
        Else
            MsgBox(log & " File does not exist !", MsgBoxStyle.Critical, "Office Tools")
        End If
    End Sub
    Private Sub WriteFrRobo()
        Dim destWriter As New StreamWriter(logPath, True)
        destWriter.WriteLine(PathVal(roboPath, 1))
        destWriter.WriteLine(PathVal(roboPath, 2))
        destWriter.WriteLine(PathVal(roboPath, 3))
        destWriter.WriteLine(PathVal(roboPath, CInt(File.ReadAllLines(roboPath).Length) - 11))
        destWriter.WriteLine(PathVal(roboPath, CInt(File.ReadAllLines(roboPath).Length) - 10))
        destWriter.WriteLine(PathVal(roboPath, CInt(File.ReadAllLines(roboPath).Length) - 9))
        destWriter.WriteLine(PathVal(roboPath, CInt(File.ReadAllLines(roboPath).Length) - 8))
        destWriter.WriteLine(PathVal(roboPath, CInt(File.ReadAllLines(roboPath).Length) - 7))
        destWriter.WriteLine(PathVal(roboPath, CInt(File.ReadAllLines(roboPath).Length) - 6))
        destWriter.WriteLine(PathVal(roboPath, CInt(File.ReadAllLines(roboPath).Length) - 5))
        destWriter.WriteLine(PathVal(roboPath, CInt(File.ReadAllLines(roboPath).Length) - 4))
        destWriter.WriteLine(PathVal(roboPath, CInt(File.ReadAllLines(roboPath).Length) - 3))
        destWriter.WriteLine(PathVal(roboPath, CInt(File.ReadAllLines(roboPath).Length) - 2))
        destWriter.Close()
    End Sub
End Class
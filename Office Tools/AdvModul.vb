Imports System.IO
Imports Syncfusion.Windows.Forms

Module AdvModul
    Dim logPath As String = "log/log"
    Dim lastResult As String = "log/lastResult"
    Dim lastErr As String = "log/lastErr"
    Dim roboPath As String = "log/robolog"
    Dim uiSrcPath As String = "conf/nrm_backup/nrmSrcPath"
    Dim uiDestPath As String = "conf/nrm_backup/nrmDestPath"
    Dim uiFrDatePath As String = "conf/nrm_backup/nrmFrDatePath"
    Dim uiReDatePath As String = "conf/nrm_backup/nrmReDatePath"
    Dim uiToDatePath As String = "conf/nrm_backup/nrmToDatePath"
    Public Sub BeginCopy(Cmbx1 As String, Lbl7 As String, Tbx2 As String, Dtp1 As Date, Dtp2 As Date)
        If File.Exists(roboPath) Then
            PrepareNotif(roboPath)
        End If
        If Cmbx1 = "Anytime" Then
            Dim uiTrimSrc As String
            Dim uiTrimDest As String
            uiTrimSrc = Lbl7
            uiTrimDest = Tbx2
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
        ElseIf Cmbx1 = "Today" Then
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
            uiTrimSrc = Lbl7
            uiTrimDest = Tbx2
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
        ElseIf Cmbx1 = "From Date" Then
            If File.Exists(uiFrDatePath) Then
                GC.Collect()
                GC.WaitForPendingFinalizers()
                File.Delete(uiFrDatePath)
                File.Create(uiFrDatePath).Dispose()
                Dim destWriter As New StreamWriter(uiFrDatePath, True)
                Dim dt As Date = Dtp1.ToShortDateString
                destWriter.WriteLine(dt.ToString("yyyyMMdd"))
                destWriter.Close()
            Else
                File.Create(uiFrDatePath).Dispose()
                Dim destWriter As New StreamWriter(uiFrDatePath, True)
                Dim dt As Date = Dtp1.ToShortDateString
                destWriter.WriteLine(dt.ToString("yyyyMMdd"))
                destWriter.Close()
            End If
            If File.Exists(uiToDatePath) Then
                GC.Collect()
                GC.WaitForPendingFinalizers()
                File.Delete(uiToDatePath)
                File.Create(uiToDatePath).Dispose()
                Dim destWriter As New StreamWriter(uiToDatePath, True)
                Dim dt As Date = Dtp2.ToShortDateString
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
                Dim dt As Date = Dtp2.ToShortDateString
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
            uiTrimSrc = Lbl7
            uiTrimDest = Tbx2
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
    Public Sub WriteFrRobo()
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
    Public Sub ShowNotif(result As String, lastResult As String, lastErr As String)
        MessageBoxAdv.MessageBoxStyle = MessageBoxAdv.Style.Metro
        If File.Exists(lastResult) Then
            Dim lastRest As String = PathVal(lastResult, 0)
            If PathVal(lastResult, 0).Equals("success") Then
                MessageBoxAdv.Show(result & " success !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ElseIf PathVal(lastResult, 0).Equals("err") Then
                MessageBoxAdv.Show(result & " error !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If File.Exists(lastErr) Then
                    If PathVal(lastErr, 0).Equals("") Then
                        MessageBoxAdv.Show("Unknown error reason !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Else
                        MessageBoxAdv.Show(PathVal(lastErr, 0), "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                Else
                    MessageBoxAdv.Show("Error file not found !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Else
                MessageBoxAdv.Show("Unknown result status !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Else
            MessageBoxAdv.Show("Result file not found !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
    Public Function RandomString(ByRef Length As String) As String
        Dim str As String = Nothing
        Dim rnd As New Random
        For i As Integer = 0 To Length
            Dim chrInt As Integer = 0
            Do
                chrInt = rnd.Next(30, 122)
                If (chrInt >= 48 And chrInt <= 57) Or (chrInt >= 65 And chrInt <= 90) Or (chrInt >= 97 And chrInt <= 122) Then
                    Exit Do
                End If
            Loop
            str &= Chr(chrInt)
        Next
        Return str
    End Function
    Public Function CompLevel(level As String) As String
        Dim compressionLevel As String
        If level = "No" Then
            compressionLevel = "-mx0"
            Return compressionLevel
        ElseIf level = "Low" Then
            compressionLevel = "-mx1"
            Return compressionLevel
        ElseIf level = "Fast" Then
            compressionLevel = "-mx3"
            Return compressionLevel
        ElseIf level = "Normal" Then
            compressionLevel = "-mx5"
            Return compressionLevel
        ElseIf level = "Maximum" Then
            compressionLevel = "-mx7"
            Return compressionLevel
        ElseIf level = "Ultra" Then
            compressionLevel = "-mx9"
            Return compressionLevel
        Else
            compressionLevel = "-mx5"
            Return compressionLevel
        End If
    End Function
    Public Function CompType(type As String) As String
        Dim compressionType As String
        If type = "7Z" Then
            compressionType = "-t7z"
            Return compressionType
        ElseIf type = "ZIP" Then
            compressionType = "-tzip"
            Return compressionType
        Else
            compressionType = "-tzip"
            Return compressionType
        End If
    End Function
    Public Function CompExt(type As String) As String
        Dim compressionExt As String
        If type = "7Z" Then
            compressionExt = "7z"
            Return compressionExt
        ElseIf type = "ZIP" Then
            compressionExt = "zip"
            Return compressionExt
        Else
            compressionExt = "zip"
            Return compressionExt
        End If
    End Function
    Public Function PassType(pass As String) As String
        Dim passwordType As String
        If pass = "Password (No Encrypt)" Then
            passwordType = "pass_noenc"
            Return passwordType
        ElseIf pass = "Password (SHA256)" Then
            passwordType = "pass_sha256"
            Return passwordType
        ElseIf pass = "Password (HMAC SHA256)" Then
            passwordType = "pass_hmac256"
            Return passwordType
        Else
            passwordType = "pass_enc"
            Return passwordType
        End If
    End Function
End Module

Imports System.IO
Imports Syncfusion.Windows.Forms

Module GlobalModul
    Dim confPath As String = "conf/config"
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
    Dim uiDatePath As String = "conf/nrm_backup/nrmDatePath"
    Dim uiDestPath As String = "conf/nrm_backup/nrmDestPath"
    Dim uiFrDatePath As String = "conf/nrm_backup/nrmFrDatePath"
    Dim uiReDatePath As String = "conf/nrm_backup/nrmReDatePath"
    Dim uiSrcPath As String = "conf/nrm_backup/nrmSrcPath"
    Dim uiToDatePath As String = "conf/nrm_backup/nrmToDatePath"
    Dim cliSrcPath As String = "conf/cli_backup/cliSrcPath"
    Dim cliDestPath As String = "conf/cli_backup/cliDestPath"
    Dim cliDatePath As String = "conf/cli_backup/cliDatePath"
    Dim cliProcessor As String = "conf/cli_backup/cliProcessor"
    Dim cliTimeInit As String = "conf/cli_backup/cliTimeInit"
    Public Sub InitCheck()
        Dim num As Integer() = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16,
                                17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29}
        Dim conf As String() = {confPath, resSrcPath, resDestPath, resInstPath, resTempPass, resDecResult,
                                resZipLog, resDecMtd, advSrcPath, advDestPath, advCompExt, advCompLevel,
                                advEncType, advInstPath, advRanStrg, advTempPass, uiProcessorCount, uiSpecFilePath,
                                uiDatePath, uiDestPath, uiFrDatePath, uiSrcPath, uiReDatePath, uiToDatePath, cliSrcPath,
                                cliDestPath, cliDatePath, cliProcessor, cliTimeInit}
        For Each number As Integer In num
            For Each letter As String In conf
                CheckCoreComponent(letter)
            Next
        Next
    End Sub
    Public Sub CheckCoreComponent(path As String)
        If File.Exists(path) Then
        Else
            PrepareNotif(path)
        End If
    End Sub
    Public Sub CheckFileExist(path As String, trim As String)
        If File.Exists(path) Then
            GC.Collect()
            GC.WaitForPendingFinalizers()
            File.Delete(path)
            File.Create(path).Dispose()
            Dim writer As New StreamWriter(path, True)
            writer.WriteLine(trim)
            writer.Close()
        Else
            File.Create(path).Dispose()
            Dim writer As New StreamWriter(path, True)
            writer.WriteLine(trim)
            writer.Close()
        End If
    End Sub
    Public Sub PrepareNotif(path As String)
        If File.Exists(path) Then
            GC.Collect()
            GC.WaitForPendingFinalizers()
            File.Delete(path)
            File.Create(path).Dispose()
        Else
            File.Create(path).Dispose()
        End If
    End Sub
    Public Sub ClearLog(log As String, log2 As String, loopNotif As Boolean)
        MessageBoxAdv.MessageBoxStyle = MessageBoxAdv.Style.Metro
        If File.Exists(log) Then
            File.Delete(log)
            File.Create(log).Dispose()
            If loopNotif Then
                MessageBoxAdv.Show(log2 & " file reset !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Else
            If loopNotif Then
                MessageBoxAdv.Show(log2 & " file is not exist !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub
    Public Sub ManualBackup(bat As String)
        Dim psi As New ProcessStartInfo(bat) With {
            .RedirectStandardError = False,
            .RedirectStandardOutput = False,
            .CreateNoWindow = False,
            .WindowStyle = ProcessWindowStyle.Normal,
            .UseShellExecute = False
        }
        Dim process As Process = Process.Start(psi)
        process.WaitForExit()
    End Sub
    Public Function PathVal(path As String, line As Integer) As String
        Dim value As String
        If New FileInfo(path).Length = 0 Then
            value = "null"
            Return value
        Else
            value = File.ReadAllLines(path).ElementAt(line).ToString
            Return value
        End If
    End Function
    Public Function GetFileSize(file As String) As String
        Dim srcFile As String
        If file = "" Then
            srcFile = ""
            Return srcFile
        Else
            Dim newFile As New FileInfo(file)
            If newFile.Exists Then
                Dim fileSize As Double = newFile.Length / 1024 / 1024
                If fileSize < 1.0 Then
                    Dim newFileSize As Double = (newFile.Length / 1024)
                    srcFile = Format(newFileSize, "###.##").ToString & " KB"
                Else
                    srcFile = Format(fileSize, "###.##").ToString & " MB"
                End If
                Return srcFile
            Else
                srcFile = "File not found"
                Return srcFile
            End If
        End If
    End Function
    Public Function GetSrcDriveSize(dir As String) As String
        Dim srcDir As String
        If dir = "" Then
            srcDir = ""
            Return srcDir
        Else
            Dim freeSpaceSrc As Double = (My.Computer.FileSystem.GetDriveInfo(dir.Remove(3)).TotalFreeSpace / 1024 / 1024 / 1024)
            Dim totalspaceSrc As Double = (My.Computer.FileSystem.GetDriveInfo(dir.Remove(3)).TotalSize / 1024 / 1024 / 1024)
            srcDir = "Source drive size : " & Format(freeSpaceSrc, "###.##").ToString & " GB" & " | " & Format(totalspaceSrc, "###.##").ToString & " GB"
            Return srcDir
        End If
    End Function
    Public Function GetDestDriveSize(dir As String) As String
        Dim destDir As String
        If dir = "" Then
            destDir = ""
            Return destDir
        Else
            Dim freeSpaceDest As Double = (My.Computer.FileSystem.GetDriveInfo(dir.Remove(3)).TotalFreeSpace / 1024 / 1024 / 1024)
            Dim totalspaceDest As Double = (My.Computer.FileSystem.GetDriveInfo(dir.Remove(3)).TotalSize / 1024 / 1024 / 1024)
            destDir = "Destination drive size : " & Format(freeSpaceDest, "###.##").ToString & " GB" & " | " & Format(totalspaceDest, "###.##").ToString & " GB"
            Return destDir
        End If
    End Function
    Public Sub WriteLogicalCount(proc As String)
        If File.Exists(proc) Then
            GC.Collect()
            GC.WaitForPendingFinalizers()
            File.Delete(proc)
            File.Create(proc).Dispose()
            Dim destWriter As New StreamWriter(proc, True)
            destWriter.WriteLine(Environment.ProcessorCount.ToString)
            destWriter.Close()
        Else
            File.Create(proc).Dispose()
            Dim destWriter As New StreamWriter(proc, True)
            destWriter.WriteLine(Environment.ProcessorCount.ToString)
            destWriter.Close()
        End If
    End Sub
    Public Sub ExportLog(logpath As String, filename As String, logname As String, logCountPath As String)
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
    Public Function CheckNull(curCount As String) As Integer
        Dim result As Integer
        If PathVal(curCount, 0).Equals("null") Then
            result = 0
            Return result
        Else
            result = CInt(PathVal(curCount, 0))
            Return result
        End If
    End Function
    Public Function ShowLog(log As String, path As String) As String
        Dim value As String
        MessageBoxAdv.MessageBoxStyle = MessageBoxAdv.Style.Metro
        If File.Exists(path) Then
            If New FileInfo(path).Length.Equals(0) Then
                MessageBoxAdv.Show(log & " is empty !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                value = ""
                Return value
            Else
                value = File.ReadAllText(path)
                Return value
            End If
        Else
            MessageBoxAdv.Show(log & " does not exist !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
            value = ""
            Return value
        End If
    End Function
End Module
Imports System.IO
Module GlobalModul
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
        If File.Exists(log) Then
            File.Delete(log)
            File.Create(log).Dispose()
            If loopNotif Then
                MsgBox(log2 & " file reset !", MsgBoxStyle.Information, "Office Tools")
            End If
        Else
            If loopNotif Then
                MsgBox(log2 & " file is not exist !", MsgBoxStyle.Critical, "Office Tools")
            End If
        End If
    End Sub
    Public Sub ManualBackup(bat As String)
        Dim psi As New ProcessStartInfo(bat) With {
            .RedirectStandardError = False,
            .RedirectStandardOutput = False,
            .CreateNoWindow = True,
            .WindowStyle = ProcessWindowStyle.Hidden,
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
        If File.Exists(path) Then
            If New FileInfo(path).Length.Equals(0) Then
                MsgBox(log & " is empty !", MsgBoxStyle.Critical, "Office Tools")
                value = ""
                Return value
            Else
                value = File.ReadAllText(path)
                Return value
            End If
        Else
            MsgBox(log & " does not exist !", MsgBoxStyle.Critical, "Office Tools")
            value = ""
            Return value
        End If
    End Function
End Module
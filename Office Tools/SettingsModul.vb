Imports System.IO
Imports Microsoft.Win32.TaskScheduler
Imports Syncfusion.Windows.Forms
Module SettingsModul
    Public Sub FindTask(strTaskName As String)
        MessageBoxAdv.MessageBoxStyle = MessageBoxAdv.Style.Metro
        Using tService As New TaskService()
            Dim tTask As Task = tService.GetTask(strTaskName)
            If tTask Is Nothing Then
                MessageBoxAdv.Show("Scheduler failed to create !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                MessageBoxAdv.Show("Scheduler successfully created !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Using
    End Sub
    Public Sub WriteForAutoBackup(confPath As String, cliSrcPath As String, cliDestPath As String, cliDatePath As String, timePath As String)
        Dim SourceBackup As String = FindConfig(confPath, "Source Directory: ")
        Dim DestBackup As String = FindConfig(confPath, "Destination Directory: ")
        Dim BackupPreferences As String = FindConfig(confPath, "Backup Preferences: ")
        MessageBoxAdv.MessageBoxStyle = MessageBoxAdv.Style.Metro
        CheckFileExist(cliSrcPath, SourceBackup)
        CheckFileExist(cliDestPath, DestBackup)
        If BackupPreferences = "null" Then
            MessageBoxAdv.Show("Please select backup options first !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If BackupPreferences = "Today" Then
                If File.Exists(cliDatePath) Then
                    GC.Collect()
                    GC.WaitForPendingFinalizers()
                    File.Delete(cliDatePath)
                    File.Create(cliDatePath).Dispose()
                    Dim destWriter As New StreamWriter(cliDatePath, True)
                    Dim dt As Date = Today
                    destWriter.WriteLine(dt.ToString("MM-dd-yyyy"))
                    destWriter.Close()
                    If File.Exists(timePath) Then
                        GC.Collect()
                        GC.WaitForPendingFinalizers()
                        File.Delete(timePath)
                        File.Create(timePath).Dispose()
                        Dim timeWriter As New StreamWriter(timePath, True)
                        timeWriter.WriteLine("Today")
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
                End If
            Else
                File.Create(cliDatePath).Dispose()
                Dim destWriter As New StreamWriter(cliDatePath, True)
                Dim dt As Date = Today
                destWriter.WriteLine("Anytime")
                destWriter.Close()
                If File.Exists(timePath) Then
                    GC.Collect()
                    GC.WaitForPendingFinalizers()
                    File.Delete(timePath)
                    File.Create(timePath).Dispose()
                    Dim timeWriter As New StreamWriter(timePath, True)
                    timeWriter.WriteLine("Anytime")
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
            End If
        End If
    End Sub
    Public Function CustRepDurIntDaily(cmbx4 As String) As Integer
        Dim task As Integer
        If cmbx4 = "Disabled" Then
            task = 0
            Return task
        ElseIf cmbx4 = "5 Minutes" Then
            task = 5
            Return task
        ElseIf cmbx4 = "10 Minutes" Then
            task = 10
            Return task
        ElseIf cmbx4 = "15 Minutes" Then
            task = 15
            Return task
        ElseIf cmbx4 = "30 Minutes" Then
            task = 30
            Return task
        ElseIf cmbx4 = "1 Hours" Then
            task = 1
            Return task
        Else
            task = 0
            Return task
        End If
    End Function
    Public Function CustRepDurValDaily(cmbx5 As String) As Integer
        Dim repDurValue As Integer
        If cmbx5 = "Disabled" Then
            repDurValue = 0
            Return repDurValue = 0
        ElseIf cmbx5 = "15 Minutes" Then
            repDurValue = 15
            Return repDurValue
        ElseIf cmbx5 = "30 Minutes" Then
            repDurValue = 30
            Return repDurValue
        ElseIf cmbx5 = "1 Hours" Then
            repDurValue = 1
            Return repDurValue
        ElseIf cmbx5 = "12 Hours" Then
            repDurValue = 12
            Return repDurValue
        ElseIf cmbx5 = "1 Day" Then
            repDurValue = 1
            Return repDurValue
        Else
            repDurValue = 0
            Return repDurValue
        End If
    End Function
    Public Function CustRepDurIntDecDaily(cmbx4 As String) As Integer
        Dim repDurIntDec As Integer
        If cmbx4 = "Disabled" Then
            repDurIntDec = 0
            Return repDurIntDec
        ElseIf cmbx4 = "5 Minutes" Then
            repDurIntDec = 1
            Return repDurIntDec
        ElseIf cmbx4 = "10 Minutes" Then
            repDurIntDec = 1
            Return repDurIntDec
        ElseIf cmbx4 = "15 Minutes" Then
            repDurIntDec = 1
            Return repDurIntDec
        ElseIf cmbx4 = "30 Minutes" Then
            repDurIntDec = 1
            Return repDurIntDec
        ElseIf cmbx4 = "1 Hours" Then
            repDurIntDec = 2
            Return repDurIntDec
        Else
            repDurIntDec = 0
            Return repDurIntDec
        End If
    End Function
    Public Function CustRepDurValDecDaily(cmbx5 As String) As Integer
        Dim repDurValDec As Integer
        If cmbx5 = "Disabled" Then
            repDurValDec = 0
            Return repDurValDec
        ElseIf cmbx5 = "15 Minutes" Then
            repDurValDec = 1
            Return repDurValDec
        ElseIf cmbx5 = "30 Minutes" Then
            repDurValDec = 1
            Return repDurValDec
        ElseIf cmbx5 = "1 Hours" Then
            repDurValDec = 2
            Return repDurValDec
        ElseIf cmbx5 = "12 Hours" Then
            repDurValDec = 2
            Return repDurValDec
        ElseIf cmbx5 = "1 Day" Then
            repDurValDec = 3
            Return repDurValDec
        Else
            repDurValDec = 0
            Return repDurValDec
        End If
    End Function
    Public Function CustRepDurIntWeek(cmbx7 As String) As Integer
        Dim task As Integer
        If cmbx7 = "Disabled" Then
            task = 0
            Return task
        ElseIf cmbx7 = "5 Minutes" Then
            task = 5
            Return task
        ElseIf cmbx7 = "10 Minutes" Then
            task = 10
            Return task
        ElseIf cmbx7 = "15 Minutes" Then
            task = 15
            Return task
        ElseIf cmbx7 = "30 Minutes" Then
            task = 30
            Return task
        ElseIf cmbx7 = "1 Hours" Then
            task = 1
            Return task
        Else
            task = 0
            Return task
        End If
    End Function
    Public Function CustRepDurValWeek(cmbx6 As String) As Integer
        Dim repDurValue As Integer
        If cmbx6 = "Disabled" Then
            repDurValue = 0
            Return repDurValue = 0
        ElseIf cmbx6 = "15 Minutes" Then
            repDurValue = 15
            Return repDurValue
        ElseIf cmbx6 = "30 Minutes" Then
            repDurValue = 30
            Return repDurValue
        ElseIf cmbx6 = "1 Hours" Then
            repDurValue = 1
            Return repDurValue
        ElseIf cmbx6 = "12 Hours" Then
            repDurValue = 12
            Return repDurValue
        ElseIf cmbx6 = "1 Day" Then
            repDurValue = 1
            Return repDurValue
        Else
            repDurValue = 0
            Return repDurValue
        End If
    End Function
    Public Function CustRepDurIntDecWeek(cmbx7 As String) As Integer
        Dim repDurIntDec As Integer
        If cmbx7 = "Disabled" Then
            repDurIntDec = 0
            Return repDurIntDec
        ElseIf cmbx7 = "5 Minutes" Then
            repDurIntDec = 1
            Return repDurIntDec
        ElseIf cmbx7 = "10 Minutes" Then
            repDurIntDec = 1
            Return repDurIntDec
        ElseIf cmbx7 = "15 Minutes" Then
            repDurIntDec = 1
            Return repDurIntDec
        ElseIf cmbx7 = "30 Minutes" Then
            repDurIntDec = 1
            Return repDurIntDec
        ElseIf cmbx7 = "1 Hours" Then
            repDurIntDec = 2
            Return repDurIntDec
        Else
            repDurIntDec = 0
            Return repDurIntDec
        End If
    End Function
    Public Function CustRepDurValDecWeek(cmbx6 As String) As Integer
        Dim repDurValDec As Integer
        If cmbx6 = "Disabled" Then
            repDurValDec = 0
            Return repDurValDec
        ElseIf cmbx6 = "15 Minutes" Then
            repDurValDec = 1
            Return repDurValDec
        ElseIf cmbx6 = "30 Minutes" Then
            repDurValDec = 1
            Return repDurValDec
        ElseIf cmbx6 = "1 Hours" Then
            repDurValDec = 2
            Return repDurValDec
        ElseIf cmbx6 = "12 Hours" Then
            repDurValDec = 2
            Return repDurValDec
        ElseIf cmbx6 = "1 Day" Then
            repDurValDec = 3
            Return repDurValDec
        Else
            repDurValDec = 0
            Return repDurValDec
        End If
    End Function
    Public Sub DailyTrigger(custDate As String, repDayInt As Integer, custInt As Integer, custrepdur As Integer, custrepint As Integer, cmbx5 As String, cmbx4 As String)
        Dim appPath As String = Application.StartupPath()
        MessageBoxAdv.MessageBoxStyle = MessageBoxAdv.Style.Metro
        Using tService As New TaskService()
            Dim tTask As Task = tService.GetTask("Office Tools")
            Dim repDurVal As Integer = CustRepDurValDecDaily(cmbx5)
            Dim repDurInt As Integer = CustRepDurIntDecDaily(cmbx4)
            Dim tDefinition As TaskDefinition = tService.NewTask
            Dim tTrigger As New DailyTrigger()
            If tTask Is Nothing Then
                tDefinition.RegistrationInfo.Description = "Office Tools - Daily Task"
                tTrigger.StartBoundary = custDate
                If repDayInt = 1 Then
                    tTrigger.DaysInterval = custInt
                End If
                If repDurInt = 1 Then
                    tTrigger.Repetition.Interval = TimeSpan.FromMinutes(custrepint)
                ElseIf repDurInt = 2 Then
                    tTrigger.Repetition.Interval = TimeSpan.FromHours(custrepint)
                End If
                If repDurVal = 1 Then
                    tTrigger.Repetition.Duration = TimeSpan.FromMinutes(custrepdur)
                ElseIf repDurVal = 2 Then
                    tTrigger.Repetition.Duration = TimeSpan.FromHours(custrepdur)
                ElseIf repDurVal = 3 Then
                    tTrigger.Repetition.Duration = TimeSpan.FromDays(custrepdur)
                End If
                tDefinition.Triggers.Add(tTrigger)
                tDefinition.Actions.Add(New ExecAction(appPath & "bat\MigrateToGDrive_Init.bat", "", appPath & "bat"))
                tService.RootFolder.RegisterTaskDefinition("Office Tools", tDefinition)
                FindTask("Office Tools")
            Else
                Dim validation As Integer
                validation = MsgBox("Old scheduler already exist !, Create a new scheduler ?", vbExclamation + vbYesNo + vbDefaultButton2, "Office Tools")
                If validation = vbYes Then
                    tService.RootFolder.DeleteTask("Office Tools")
                    tDefinition.RegistrationInfo.Description = "Office Tools - Daily Task"
                    tTrigger.StartBoundary = custDate
                    If repDayInt = 1 Then
                        tTrigger.DaysInterval = custInt
                    End If
                    If repDurInt = 1 Then
                        tTrigger.Repetition.Interval = TimeSpan.FromMinutes(custrepint)
                    ElseIf repDurInt = 2 Then
                        tTrigger.Repetition.Interval = TimeSpan.FromHours(custrepint)
                    End If
                    If repDurVal = 1 Then
                        tTrigger.Repetition.Duration = TimeSpan.FromMinutes(custrepdur)
                    ElseIf repDurVal = 2 Then
                        tTrigger.Repetition.Duration = TimeSpan.FromHours(custrepdur)
                    ElseIf repDurVal = 3 Then
                        tTrigger.Repetition.Duration = TimeSpan.FromDays(custrepdur)
                    End If
                    tDefinition.Triggers.Add(tTrigger)
                    tDefinition.Actions.Add(New ExecAction(appPath & "bat\MigrateToGDrive_Init.bat", "", appPath & "bat"))
                    tService.RootFolder.RegisterTaskDefinition("Office Tools", tDefinition)
                    FindTask("Office Tools")
                Else
                    MessageBoxAdv.Show("Cancel To create scheduler !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If
        End Using
    End Sub
    Public Sub WeeklyTrigger(custdate As String, repDayInt As Integer, custInt As Integer, custrepdur As Integer, custrepint As Integer, cb1 As DaysOfTheWeek, cb2 As DaysOfTheWeek, cb3 As DaysOfTheWeek, cb4 As DaysOfTheWeek, cb5 As DaysOfTheWeek, cb6 As DaysOfTheWeek, cb7 As DaysOfTheWeek, cmbx6 As String, cmbx7 As String)
        Dim appPath As String = Application.StartupPath()
        Dim repDurVal As Integer = CustRepDurValDecWeek(cmbx6)
        Dim repDurInt As Integer = CustRepDurIntDecWeek(cmbx7)
        MessageBoxAdv.MessageBoxStyle = MessageBoxAdv.Style.Metro
        Using tService As New TaskService()
            Dim tTask As Task = tService.GetTask("Office Tools")
            Dim tDefinition As TaskDefinition = tService.NewTask
            Dim tTrigger As New WeeklyTrigger()
            If tTask Is Nothing Then
                tDefinition.RegistrationInfo.Description = "Office Tools - Weekly Task"
                tTrigger.StartBoundary = custdate
                tTrigger.DaysOfWeek = cb1 Or cb2 Or cb3 Or cb4 Or cb5 Or cb6 Or cb7
                If repDayInt = 1 Then
                    tTrigger.WeeksInterval = custInt
                End If
                If repDurInt = 1 Then
                    tTrigger.Repetition.Interval = TimeSpan.FromMinutes(custrepint)
                ElseIf repDurInt = 2 Then
                    tTrigger.Repetition.Interval = TimeSpan.FromHours(custrepint)
                End If
                If repDurVal = 1 Then
                    tTrigger.Repetition.Duration = TimeSpan.FromMinutes(custrepdur)
                ElseIf repDurVal = 2 Then
                    tTrigger.Repetition.Duration = TimeSpan.FromHours(custrepdur)
                ElseIf repDurVal = 3 Then
                    tTrigger.Repetition.Duration = TimeSpan.FromDays(custrepdur)
                End If
                tDefinition.Triggers.Add(tTrigger)
                tDefinition.Actions.Add(New ExecAction(appPath & "bat\MigrateToGDrive_Init.bat", "", appPath & "bat"))
                tService.RootFolder.RegisterTaskDefinition("Office Tools", tDefinition)
                FindTask("Office Tools")
            Else
                Dim validation As Integer
                validation = MsgBox("Old scheduler already exist !, Create a new scheduler ?", vbExclamation + vbYesNo + vbDefaultButton2, "Office Tools")
                If validation = vbYes Then
                    tService.RootFolder.DeleteTask("Office Tools")
                    tDefinition.RegistrationInfo.Description = "Office Tools - Weekly Task"
                    tTrigger.StartBoundary = custdate
                    tTrigger.DaysOfWeek = cb1 Or cb2 Or cb3 Or cb4 Or cb5 Or cb6 Or cb7
                    If repDayInt = 1 Then
                        tTrigger.WeeksInterval = custInt
                    End If
                    If repDurInt = 1 Then
                        tTrigger.Repetition.Interval = TimeSpan.FromMinutes(custrepint)
                    ElseIf repDurInt = 2 Then
                        tTrigger.Repetition.Interval = TimeSpan.FromHours(custrepint)
                    End If
                    If repDurVal = 1 Then
                        tTrigger.Repetition.Duration = TimeSpan.FromMinutes(custrepdur)
                    ElseIf repDurVal = 2 Then
                        tTrigger.Repetition.Duration = TimeSpan.FromHours(custrepdur)
                    ElseIf repDurVal = 3 Then
                        tTrigger.Repetition.Duration = TimeSpan.FromDays(custrepdur)
                    End If
                    tDefinition.Triggers.Add(tTrigger)
                    tDefinition.Actions.Add(New ExecAction(appPath & "bat\MigrateToGDrive_Init.bat", "", appPath & "bat"))
                    tService.RootFolder.RegisterTaskDefinition("Office Tools", tDefinition)
                    FindTask("Office Tools")
                Else
                    MessageBoxAdv.Show("Cancel To create scheduler !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If
        End Using
    End Sub
    Public Function Cb1(Cmbx1 As Boolean) As DaysOfTheWeek
        Dim monday As DaysOfTheWeek
        If Cmbx1 Then
            monday = DaysOfTheWeek.Monday
            Return monday
        Else
            monday = DaysOfTheWeek.Sunday
            Return monday
        End If
    End Function
    Public Function Cb2(Cmbx2 As Boolean) As DaysOfTheWeek
        Dim tuesday As DaysOfTheWeek
        If Cmbx2 Then
            tuesday = DaysOfTheWeek.Tuesday
            Return tuesday
        Else
            tuesday = DaysOfTheWeek.Sunday
            Return tuesday
        End If
    End Function
    Public Function Cb3(Cmbx3 As Boolean) As DaysOfTheWeek
        Dim wednesday As DaysOfTheWeek
        If Cmbx3 Then
            wednesday = DaysOfTheWeek.Wednesday
            Return wednesday
        Else
            wednesday = DaysOfTheWeek.Sunday
            Return wednesday
        End If
    End Function
    Public Function Cb4(Cmbx4 As Boolean) As DaysOfTheWeek
        Dim thursday As DaysOfTheWeek
        If Cmbx4 Then
            thursday = DaysOfTheWeek.Thursday
            Return thursday
        Else
            thursday = DaysOfTheWeek.Sunday
            Return thursday
        End If
    End Function
    Public Function Cb5(cmbx5 As Boolean) As DaysOfTheWeek
        Dim friday As DaysOfTheWeek
        If cmbx5 Then
            friday = DaysOfTheWeek.Friday
            Return friday
        Else
            friday = DaysOfTheWeek.Sunday
            Return friday
        End If
    End Function
    Public Function Cb6(cmbx6 As Boolean) As DaysOfTheWeek
        Dim saturday As DaysOfTheWeek
        If cmbx6 Then
            saturday = DaysOfTheWeek.Saturday
            Return saturday
        Else
            saturday = DaysOfTheWeek.Sunday
            Return saturday
        End If
    End Function
    Public Function Cb7(cmbx7 As Boolean) As DaysOfTheWeek
        Dim sunday As DaysOfTheWeek
        If cmbx7 Then
            sunday = DaysOfTheWeek.Sunday
            Return sunday
        Else
            sunday = DaysOfTheWeek.Sunday
            Return sunday
        End If
    End Function
    Public Function ShowTask(taskName As String) As String
        Dim value As String
        MessageBoxAdv.MessageBoxStyle = MessageBoxAdv.Style.Metro
        Using tService As New TaskService()
            Dim tTask As Task = tService.GetTask(taskName)
            If tTask Is Nothing Then
                MessageBoxAdv.Show("Office Task scheduler not exist !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                MessageBoxAdv.Show("Please create new scheduler first !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                value = ""
                Return value
            Else
                value = "Task Name: " & tTask.Name & vbCrLf & "Task State: " & tTask.State.ToString & vbCrLf &
                                    "Task Path: " & tTask.Path.ToString & vbCrLf &
                                    "Next Runtime: " & tTask.NextRunTime.ToLongDateString & " " & tTask.NextRunTime.ToLongTimeString & vbCrLf &
                                    "Last Runtime: " & tTask.LastRunTime.ToLongDateString & " " & tTask.LastRunTime.ToLongTimeString & vbCrLf &
                                    "Last Task Result: " & tTask.LastTaskResult.ToString & vbCrLf &
                                    "Total Failed Task: " & tTask.NumberOfMissedRuns.ToString & vbCrLf
                Return value
            End If
        End Using
    End Function
    Public Sub WriteFile(filePath As String, noLine As Integer, newText As String)
        Dim lines() As String = File.ReadAllLines(filePath)
        lines(noLine) = newText
        File.WriteAllLines(filePath, lines)
    End Sub
    Public Function FindConfig(confpath As String, contains As String) As String
        Dim value As String
        Using sReader As New StreamReader(confpath)
            While Not sReader.EndOfStream
                Dim line As String = sReader.ReadLine()
                If line.Contains(contains) Then
                    value = line
                    Return value
                End If
            End While
            value = "null"
            Return value
        End Using
    End Function
End Module
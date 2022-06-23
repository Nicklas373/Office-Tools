Imports System.IO
Public Class about_menu
    Dim changelog As String = "changelog.txt"
    Private Sub About_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AllowTransparency = False
        desc_pnl.Visible = False
        info_pnl.Visible = False
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        desc_pnl.Visible = True
        info_pnl.Visible = False
        history_pnl.Visible = False
        RichTextBox1.Text = "Description: " & vbCrLf & vbCrLf & My.Application.Info.Description
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        desc_pnl.Visible = True
        info_pnl.Visible = True
        history_pnl.Visible = False
        TextBox1.Text = My.Application.Info.ProductName
        TextBox2.Text = My.Application.Info.Version.ToString
        TextBox3.Text = "June, 23 2022"
        TextBox4.Text = My.Application.Info.Copyright
        TextBox5.Text = My.Application.Info.DirectoryPath
        TextBox6.Text = "PT Elwilis Mitra Sejahtera"
        TextBox7.Text = My.Computer.Name.ToString
        TextBox8.Text = My.Computer.Info.OSFullName.ToString
        TextBox9.Text = My.Computer.Info.OSVersion.ToString
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        desc_pnl.Visible = True
        info_pnl.Visible = True
        history_pnl.Visible = True
        ReadLog("changelog", changelog)
    End Sub
    Private Sub ReadLog(log As String, path As String)
        RichTextBox2.Text = ""
        If File.Exists(path) Then
            If New FileInfo(path).Length.Equals(0) Then
                MsgBox(log & " is empty !", MsgBoxStyle.Critical, "MigrateToGDrive")
            Else
                RichTextBox2.Text = "Changelog: " & vbCrLf & vbCrLf & File.ReadAllText(path)
            End If
        Else
            MsgBox(log & " does not exist !", MsgBoxStyle.Critical, "MigrateToGDrive")
        End If
    End Sub
End Class
Imports Syncfusion.WinForms.Controls
Public Class AboutMenu
    Inherits SfForm
    Dim changelog As String = "changelog.txt"
    Private Sub About_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AllowTransparency = False
        Style.TitleBar.IconBackColor = Color.FromArgb(15, 161, 212)
        BackColor = Color.AliceBlue
        Style.TitleBar.TextHorizontalAlignment = HorizontalAlignment.Center
        Style.TitleBar.TextVerticalAlignment = VisualStyles.VerticalAlignment.Center
        desc_pnl.Visible = False
        info_pnl.Visible = False
    End Sub
    Private Sub Description_Button(sender As Object, e As EventArgs) Handles Button5.Click
        desc_pnl.Visible = True
        info_pnl.Visible = False
        history_pnl.Visible = False
        RichTextBox1.Text = "Description: " & vbCrLf & vbCrLf & My.Application.Info.Description
    End Sub
    Private Sub Close_Button(sender As Object, e As EventArgs) Handles Button4.Click
        Close()
    End Sub
    Private Sub Information_Button(sender As Object, e As EventArgs) Handles Button1.Click
        desc_pnl.Visible = True
        info_pnl.Visible = True
        history_pnl.Visible = False
        TextBox1.Text = My.Application.Info.ProductName
        TextBox2.Text = My.Application.Info.Version.ToString
        TextBox3.Text = "November, 18 2022"
        TextBox4.Text = My.Application.Info.Copyright
        TextBox5.Text = My.Application.Info.DirectoryPath
        TextBox6.Text = "PLACEHOLDER"
        TextBox7.Text = My.Computer.Name.ToString
        TextBox8.Text = My.Computer.Info.OSFullName.ToString
        TextBox9.Text = My.Computer.Info.OSVersion.ToString
    End Sub
    Private Sub Changelog_Button(sender As Object, e As EventArgs) Handles Button2.Click
        desc_pnl.Visible = True
        info_pnl.Visible = True
        history_pnl.Visible = True
        RichTextBox2.Text = "Changelog: " & vbCrLf & vbCrLf & ShowLog("Changelog", changelog)
    End Sub
End Class
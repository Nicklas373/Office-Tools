Imports Syncfusion.WinForms.Controls

Public Class MainMenu
    Inherits SfForm
    Private Sub MainMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AllowTransparency = False
        InitCheck()
        Style.TitleBar.IconBackColor = Color.FromArgb(15, 161, 212)
        BackColor = Color.AliceBlue
        Style.TitleBar.TextHorizontalAlignment = HorizontalAlignment.Center
        Style.TitleBar.TextVerticalAlignment = VisualStyles.VerticalAlignment.Center
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Close()
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim menu_backup = New BackupMenu_2
        menu_backup.Show()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim menu_history = New HistoryMenu
        menu_history.Show()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim menu_pdf = New PDFMenu
        menu_pdf.Show()
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim menu_settings = New SettingsMenu
        menu_settings.Show()
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim menu_about = New AboutMenu
        menu_about.Show()
    End Sub
End Class
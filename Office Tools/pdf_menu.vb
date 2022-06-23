Imports System.IO
Imports Syncfusion.Pdf
Imports Syncfusion.Pdf.Parsing

Public Class pdf_menu
    Dim fileDialog As New OpenFileDialog
    Dim saveDialog As New SaveFileDialog
    Private Sub pdf_compress_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AllowTransparency = False
        pdf_com_pnl.Visible = False
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If TextBox1.ReadOnly = True Then
            MsgBox("Configuration menu is locked, Please click edit !", MsgBoxStyle.Information, "Office Tools")
        Else
            fileDialog.DefaultExt = ".pdf"
            fileDialog.Filter = "PDF File | *.pdf"
            fileDialog.InitialDirectory = Environment.SpecialFolder.UserProfile
            fileDialog.ShowDialog()
        End If
    End Sub
    Private Sub OpenFileDialog_Disposed(sender As Object, e As EventArgs) Handles Button8.Click
        If TextBox1.ReadOnly = False Then
            If fileDialog.FileName.ToString = "" Then
                TextBox1.Text = ""
            Else
                TextBox1.Text = Path.GetFullPath(fileDialog.FileName.ToString)
                Label8.Text = getFileSize(TextBox1.Text)
            End If
        End If
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If TextBox2.ReadOnly = True Then
            MsgBox("Configuration menu is locked, Please click edit !", MsgBoxStyle.Information, "Office Tools")
        Else
            saveDialog.DefaultExt = ".pdf"
            saveDialog.Filter = "PDF File | *.pdf"
            saveDialog.InitialDirectory = Environment.SpecialFolder.UserProfile
            saveDialog.ShowDialog()
        End If
    End Sub
    Private Sub SaveFileDialog_Disposed(sender As Object, e As EventArgs) Handles Button7.Click
        If TextBox2.ReadOnly = False Then
            If saveDialog.FileName.ToString = "" Then
                TextBox2.Text = ""
            Else
                TextBox2.Text = Path.GetFullPath(saveDialog.FileName.ToString)
            End If
        End If
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If TextBox1.Text = "" Then
            MsgBox("No PDF file was selected !, please select PDF file first !", MsgBoxStyle.Critical, "MigrateToGDrive")
        Else
            If TextBox2.Text = "" Then
                MsgBox("Destination PDF file location was not selected !, please select destination location first !", MsgBoxStyle.Critical, "MigrateToGDrive")
            Else
                If ComboBox2.Text = "" Then
                    MsgBox("Please select compression level !", MsgBoxStyle.Critical, "MigrateToGDrive")
                Else
                    CompressPDF(TextBox1.Text, TextBox2.Text, True, imgCompLvlVal, pdfFoOptVal, pdfOpcOptVal, pdfMtOptVal, pdfIncUpdVal)
                End If
            End If
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        pdf_com_pnl.Visible = True
    End Sub
    Private Async Sub CompressPDF(pdfPathIn As String, pdfPathOut As String, pdfCompOpt As Boolean, pdfImgQtyOpt As Integer, pdfOfOpt As Boolean, pdfOpcOpt As Boolean, pdfRmOpt As Boolean, pdfIncUpd As Boolean)
        ProgressBar1.Visible = True
        ProgressBar1.Style = ProgressBarStyle.Marquee
        ProgressBar1.MarqueeAnimationSpeed = 40
        ProgressBar1.Refresh()
        Dim ldoc As New PdfLoadedDocument(pdfPathIn)
        Dim options As PdfCompressionOptions = New PdfCompressionOptions
        options.CompressImages = pdfCompOpt
        options.ImageQuality = pdfImgQtyOpt
        options.OptimizeFont = pdfOfOpt
        options.OptimizePageContents = pdfOpcOpt
        options.RemoveMetadata = pdfRmOpt
        ldoc.FileStructure.IncrementalUpdate = pdfIncUpd
        ldoc.CompressionOptions = options
        Await Task.Run(Sub() ldoc.Save(pdfPathOut))
        ProgressBar1.Style = ProgressBarStyle.Blocks
        ProgressBar1.Value = 100
        ldoc.Close(True)
        If File.Exists(TextBox2.Text) Then
            MsgBox("Compress PDF success !", MsgBoxStyle.Information, "MigrateToGDrive")
            Label10.Text = getFileSize(TextBox2.Text)
        Else
            MsgBox("Compress PDF failed !", MsgBoxStyle.Critical, "MigrateToGDrive")
        End If
    End Sub
    Private Function imgCompLvlVal() As Integer
        Dim value As Integer
        If ComboBox2.Text = "Highest" Then
            value = 20
        ElseIf ComboBox2.Text = "High" Then
            value = 40
        ElseIf ComboBox2.Text = "Normal" Then
            value = 60
        ElseIf ComboBox2.Text = "Low" Then
            value = 80
        ElseIf ComboBox2.Text = "Lowest" Then
            value = 100
        Else
            value = 0
        End If

        Return value
    End Function
    Private Function pdfIncUpdVal() As Boolean
        Dim value As Boolean
        If CheckBox5.Checked Then
            value = True
        Else
            value = False
        End If

        Return value
    End Function
    Private Function pdfMtOptVal() As Boolean
        Dim value As Boolean
        If CheckBox4.Checked Then
            value = True
        Else
            value = False
        End If

        Return value
    End Function
    Private Function pdfFoOptVal() As Boolean
        Dim value As Boolean
        If CheckBox2.Checked Then
            value = True
        Else
            value = False
        End If

        Return value
    End Function
    Private Function pdfOpcOptVal() As Boolean
        Dim value As Boolean
        If CheckBox3.Checked Then
            value = True
        Else
            value = False
        End If

        Return value
    End Function
End Class
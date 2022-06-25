Imports System.IO
Imports Syncfusion.Pdf
Imports Syncfusion.Pdf.Parsing
Imports Syncfusion.Windows.Forms
Imports Syncfusion.WinForms.Controls
Public Class PDFMenu
    Inherits SfForm
    Dim fileDialog As New OpenFileDialog
    Dim saveDialog As New SaveFileDialog
    Dim confPath As String = "conf/config"
    Private Sub PDF_Compress_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AllowTransparency = False
        Style.TitleBar.IconBackColor = Color.FromArgb(15, 161, 212)
        BackColor = Color.AliceBlue
        Style.TitleBar.TextHorizontalAlignment = HorizontalAlignment.Center
        Style.TitleBar.TextVerticalAlignment = VisualStyles.VerticalAlignment.Center
        pdf_com_pnl.Visible = False
        MessageBoxAdv.MessageBoxStyle = MessageBoxAdv.Style.Metro
    End Sub
    Private Sub SourcePDF_folder(sender As Object, e As EventArgs) Handles Button8.Click
        If TextBox1.ReadOnly = True Then
            MessageBoxAdv.Show("Configuration menu is locked, Please click edit !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            fileDialog.DefaultExt = ".pdf"
            fileDialog.Filter = "PDF File | *.pdf"
            fileDialog.InitialDirectory = Environment.SpecialFolder.UserProfile
            fileDialog.ShowDialog()
        End If
    End Sub
    Private Sub OpenFileDialog_SourcePDF(sender As Object, e As EventArgs) Handles Button8.Click
        If TextBox1.ReadOnly = False Then
            If fileDialog.FileName.ToString = "" Then
                TextBox1.Text = ""
            Else
                TextBox1.Text = Path.GetFullPath(fileDialog.FileName.ToString)
                Label8.Text = GetFileSize(TextBox1.Text)
            End If
        End If
    End Sub
    Private Sub SavePDF_folder(sender As Object, e As EventArgs) Handles Button7.Click
        If TextBox2.ReadOnly = True Then
            MessageBoxAdv.Show("Configuration menu is locked, Please click edit !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            saveDialog.DefaultExt = ".pdf"
            saveDialog.Filter = "PDF File | *.pdf"
            saveDialog.InitialDirectory = Environment.SpecialFolder.UserProfile
            saveDialog.ShowDialog()
        End If
    End Sub
    Private Sub SaveFileDialog_file(sender As Object, e As EventArgs) Handles Button7.Click
        If TextBox2.ReadOnly = False Then
            If saveDialog.FileName.ToString = "" Then
                TextBox2.Text = ""
            Else
                TextBox2.Text = Path.GetFullPath(saveDialog.FileName.ToString)
            End If
        End If
    End Sub
    Private Sub Compress_Button(sender As Object, e As EventArgs) Handles Button6.Click
        If TextBox1.Text = "" Then
            MessageBoxAdv.Show("No PDF file was selected !, please select PDF file first !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If TextBox2.Text = "" Then
                MessageBoxAdv.Show("Destination PDF file location was not selected !, please select destination location first !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If ComboBox2.Text = "" Then
                    MessageBoxAdv.Show("Please select compression level !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    CompressPDF(TextBox1.Text, TextBox2.Text, True, ImgCompLvlVal(ComboBox2.Text), PdfFoOptVal(CheckBox5.Checked), PdfOpcOptVal(CheckBox4.Checked), PdfMtOptVal(CheckBox3.Checked), PdfIncUpdVal(CheckBox2.Checked))
                End If
            End If
        End If
    End Sub
    Private Sub Close_Button(sender As Object, e As EventArgs) Handles Button4.Click
        Close()
    End Sub
    Private Sub PDF_Panel_Button(sender As Object, e As EventArgs) Handles Button5.Click
        pdf_com_pnl.Visible = True
    End Sub
    Private Async Sub CompressPDF(pdfPathIn As String, pdfPathOut As String, pdfCompOpt As Boolean, pdfImgQtyOpt As Integer, pdfOfOpt As Boolean, pdfOpcOpt As Boolean, pdfRmOpt As Boolean, pdfIncUpd As Boolean)
        ProgressBar1.Visible = True
        ProgressBar1.Style = ProgressBarStyle.Marquee
        ProgressBar1.MarqueeAnimationSpeed = 40
        ProgressBar1.Refresh()
        Dim ldoc As New PdfLoadedDocument(pdfPathIn)
        Dim options As New PdfCompressionOptions With {
            .CompressImages = pdfCompOpt,
            .ImageQuality = pdfImgQtyOpt,
            .OptimizeFont = pdfOfOpt,
            .OptimizePageContents = pdfOpcOpt,
            .RemoveMetadata = pdfRmOpt
        }
        ldoc.FileStructure.IncrementalUpdate = pdfIncUpd
        ldoc.CompressionOptions = options
        Await Task.Run(Sub() ldoc.Save(pdfPathOut))
        ProgressBar1.Style = ProgressBarStyle.Blocks
        ProgressBar1.Value = 100
        ldoc.Close(True)
        If File.Exists(TextBox2.Text) Then
            MessageBoxAdv.Show("Compress PDF success !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Label10.Text = GetFileSize(TextBox2.Text)
            If File.ReadAllLines(confPath).Length > 5 Then
                If PathVal(confPath, 5).Replace("Auto Open PDF: ", "").Equals("True") Then
                    Process.Start(PathVal(confPath, 4).Replace("PDF Reader Preferences: ", ""), TextBox2.Text)
                End If
            End If
        Else
            MessageBoxAdv.Show("Compress PDF failed !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
End Class

Imports System.IO
Imports Syncfusion.Pdf
Imports Syncfusion.Pdf.Parsing
Imports Syncfusion.Windows.Forms
Imports Syncfusion.WinForms.Controls
Public Class PDFMenu
    Inherits SfForm
    Dim fileDialog As New OpenFileDialog
    Dim saveDialog As New SaveFileDialog
    Dim savefldDialog As New FolderBrowserDialog
    Dim confPath As String = "conf/config"
    Dim pdfMergeList As New List(Of String)()
    Dim pdfSplit As String
    Dim pdfSplit2 As String
    Private Sub PDF_Compress_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AllowTransparency = False
        Style.TitleBar.IconBackColor = Color.FromArgb(15, 161, 212)
        BackColor = Color.AliceBlue
        Style.TitleBar.TextHorizontalAlignment = HorizontalAlignment.Center
        Style.TitleBar.TextVerticalAlignment = VisualStyles.VerticalAlignment.Center
        pdf_com_pnl.Visible = False
        pdf_merge_pnl.Visible = False
        MessageBoxAdv.MessageBoxStyle = MessageBoxAdv.Style.Metro
    End Sub
    Private Sub SourcePDF_folder(sender As Object, e As EventArgs) Handles Button8.Click
        fileDialog.DefaultExt = ".pdf"
        fileDialog.Filter = "PDF File | *.pdf"
        fileDialog.Title = "Choose PDF File"
        fileDialog.InitialDirectory = Environment.SpecialFolder.UserProfile
        fileDialog.ShowDialog()
    End Sub
    Private Sub OpenFileDialog_SourcePDF(sender As Object, e As EventArgs) Handles Button8.Click
        If fileDialog.FileName.ToString = "" Then
            TextBox1.Text = ""
        Else
            TextBox1.Text = Path.GetFullPath(fileDialog.FileName.ToString)
            Label8.Text = GetFileSize(TextBox1.Text)
        End If
    End Sub
    Private Sub SavePDF_folder(sender As Object, e As EventArgs) Handles Button7.Click
        saveDialog.DefaultExt = ".pdf"
        saveDialog.Filter = "PDF File | *.pdf"
        saveDialog.Title = "Save PDF File"
        saveDialog.InitialDirectory = Environment.SpecialFolder.UserProfile
        saveDialog.ShowDialog()
    End Sub
    Private Sub SaveFileDialog_file(sender As Object, e As EventArgs) Handles Button7.Click
        If saveDialog.FileName.ToString = "" Then
            TextBox2.Text = ""
        Else
            TextBox2.Text = Path.GetFullPath(saveDialog.FileName.ToString)
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
        pdf_merge_pnl.Visible = False
        split_pdf_pnl.Visible = False
        cnv_pdf_pnl.Visible = False
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        pdf_com_pnl.Visible = True
        pdf_merge_pnl.Visible = True
        split_pdf_pnl.Visible = False
        cnv_pdf_pnl.Visible = False
    End Sub
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        pdf_com_pnl.Visible = True
        pdf_merge_pnl.Visible = True
        split_pdf_pnl.Visible = True
        cnv_pdf_pnl.Visible = False
    End Sub
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        pdf_com_pnl.Visible = True
        pdf_merge_pnl.Visible = True
        split_pdf_pnl.Visible = True
        cnv_pdf_pnl.Visible = True
    End Sub
    Private Sub OpenFileDialog_MergePDF(sender As Object, e As EventArgs) Handles Button9.Click
        Dim tempList As String
        fileDialog.DefaultExt = ".pdf"
        fileDialog.Filter = "PDF File | *.pdf"
        fileDialog.Title = "Choose PDF File"
        fileDialog.Multiselect = True
        fileDialog.InitialDirectory = Environment.SpecialFolder.UserProfile
        If fileDialog.ShowDialog() = DialogResult.OK Then
            For Each mfile As String In fileDialog.FileNames
                tempList &= mfile & vbCrLf
                pdfMergeList.Add(mfile)
            Next
            RichTextBox1.Text = tempList
        Else
            RichTextBox1.Text = ""
        End If
    End Sub
    Private Sub SaveFileDialog_MergePDF(sender As Object, e As EventArgs) Handles Button3.Click
        saveDialog.DefaultExt = ".pdf"
        saveDialog.Filter = "PDF File | *.pdf"
        saveDialog.Title = "Save PDF File"
        saveDialog.InitialDirectory = Environment.SpecialFolder.UserProfile
        saveDialog.ShowDialog()
    End Sub
    Private Sub SaveFileDialog_MergePDF_Handler(sender As Object, e As EventArgs) Handles Button3.Click
        If saveDialog.FileName.ToString = "" Then
            TextBox4.Text = ""
        Else
            TextBox4.Text = Path.GetFullPath(saveDialog.FileName.ToString)
        End If
    End Sub
    Private Sub MergePDF_Button(sender As Object, e As EventArgs) Handles Button2.Click
        If RichTextBox1.Text = "" Then
            MessageBoxAdv.Show("No PDF file was selected !, please select PDF file first !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If TextBox4.Text = "" Then
                MessageBoxAdv.Show("Destination PDF file location was not selected !, please select destination location first !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                PDFMerge(pdfMergeList.ToArray, TextBox4.Text)
            End If
        End If
    End Sub
    Private Sub OpenFileDialog_SplitPDF(sender As Object, e As EventArgs) Handles Button13.Click
        fileDialog.DefaultExt = ".pdf"
        fileDialog.Filter = "PDF File | *.pdf"
        fileDialog.Title = "Choose PDF File"
        fileDialog.InitialDirectory = Environment.SpecialFolder.UserProfile
        If fileDialog.ShowDialog() = DialogResult.OK Then
            pdfSplit = "\" + Path.GetFileNameWithoutExtension(fileDialog.FileName)
            TextBox3.Text = fileDialog.FileName
        Else
            TextBox3.Text = ""
        End If
    End Sub
    Private Sub SavePDF_folder_SplitPDF(sender As Object, e As EventArgs) Handles Button12.Click
        savefldDialog.InitialDirectory = Environment.SpecialFolder.UserProfile
        savefldDialog.ShowDialog()
    End Sub
    Private Sub SaveFileDialog_file_SplitPDF(sender As Object, e As EventArgs) Handles Button12.Click
        If savefldDialog.SelectedPath.ToString = "" Then
            TextBox5.Text = ""
        Else
            TextBox5.Text = savefldDialog.SelectedPath.ToString
        End If
    End Sub
    Private Sub SplitPDF_Button(sender As Object, e As EventArgs) Handles Button11.Click
        If TextBox3.Text = "" Then
            MessageBoxAdv.Show("No PDF file was selected !, please select PDF file first !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If TextBox5.Text = "" Then
                MessageBoxAdv.Show("Destination PDF file location was not selected !, please select destination location first !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                pdfSplit2 = TextBox5.Text + pdfSplit
                SplitPDF(TextBox3.Text, pdfSplit2, TextBox5.Text)
            End If
        End If
    End Sub
    Private Sub OpenFileDialog_Cnv_PDF_Button(sender As Object, e As EventArgs) Handles Button17.Click
        fileDialog.DefaultExt = ".pdf"
        fileDialog.Filter = "PDF File | *.pdf"
        fileDialog.Title = "Choose PDF File"
        fileDialog.InitialDirectory = Environment.SpecialFolder.UserProfile
        fileDialog.ShowDialog()
    End Sub
    Private Sub OpenFileDialog_Cnv_PDF(sender As Object, e As EventArgs) Handles Button17.Click
        If fileDialog.FileName.ToString = "" Then
            TextBox6.Text = ""
        Else
            TextBox6.Text = Path.GetFullPath(fileDialog.FileName.ToString)
            Label24.Text = GetFileSize(TextBox6.Text)
        End If
    End Sub
    Private Sub SaveFileDialog_Cnv_PDF(sender As Object, e As EventArgs) Handles Button16.Click
        saveDialog.DefaultExt = ".docx|.xlsx"
        saveDialog.Filter = "DOCX File (.docx) | *.docx|XLSX File (.xlsx)|*.xlsx"
        saveDialog.Title = "Save To"
        saveDialog.InitialDirectory = Environment.SpecialFolder.UserProfile
        saveDialog.ShowDialog()
    End Sub
    Private Sub SaveFileDialog_CnvPDF_Handler(sender As Object, e As EventArgs) Handles Button16.Click
        If saveDialog.FileName.ToString = "" Then
            TextBox7.Text = ""
        Else
            TextBox7.Text = Path.GetFullPath(saveDialog.FileName.ToString)
        End If
    End Sub
    Private Sub CnvPDF_Button(sender As Object, e As EventArgs) Handles Button15.Click
        If TextBox6.Text = "" Then
            MessageBoxAdv.Show("No PDF file was selected !, please select PDF file first !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If TextBox7.Text = "" Then
                MessageBoxAdv.Show("Destination file location was not selected !, please select destination location first !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If ComboBox1.Text = "" Then
                    MessageBoxAdv.Show("Convert option was not selected !, please select destination location first !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    If ComboBox1.Text = "Documents (*.DOCX)" Then
                        CnvPDF(TextBox6.Text, TextBox7.Text, "DOCX")
                    ElseIf ComboBox1.Text = "Spreadsheets (*.XLSX)" Then
                        CnvPDF(TextBox6.Text, TextBox7.Text, "XLSX")
                    End If
                End If
            End If
        End If
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
            Dim PDFReader As String = FindConfig(confPath, "PDF Reader Preferences: ")
            Dim PDFAutoConf As String = FindConfig(confPath, "Auto Open PDF: ")
            MessageBoxAdv.Show("Compress PDF success !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Label10.Text = GetFileSize(TextBox2.Text)
            If PDFReader = "null" Then
            Else
                If PDFAutoConf = "Auto Open PDF: True" Then
                    Process.Start(PDFReader.Remove(0, 24), TextBox2.Text)
                End If
            End If
        Else
            MessageBoxAdv.Show("Compress PDF failed !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
    Private Async Sub PDFMerge(pdfarray As String(), destname As String)
        Dim finalDoc As New PdfDocument()
        Dim source As String() = pdfarray
        ProgressBar2.Visible = True
        ProgressBar2.Style = ProgressBarStyle.Marquee
        ProgressBar2.MarqueeAnimationSpeed = 40
        ProgressBar2.Refresh()
        PdfDocument.Merge(finalDoc, source)
        Await Task.Run(Sub() finalDoc.Save(destname))
        ProgressBar2.Style = ProgressBarStyle.Blocks
        ProgressBar2.Value = 100
        finalDoc.Close(True)
        If File.Exists(TextBox4.Text) Then
            MessageBoxAdv.Show("Merge PDF success !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Dim PDFReader As String = FindConfig(confPath, "PDF Reader Preferences: ")
            Dim PDFAutoConf As String = FindConfig(confPath, "Auto Open PDF: ")
            Label15.Text = GetFileSize(TextBox4.Text)
            If PDFReader = "null" Then
            Else
                If PDFAutoConf = "Auto Open PDF: True" Then
                    Process.Start(PDFReader.Remove(0, 24), TextBox4.Text)
                End If
            End If
        Else
            MessageBoxAdv.Show("Merge PDF failed !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
    Private Async Sub SplitPDF(loadPDF As String, outSplitPDF As String, outSplitFldr As String)
        ProgressBar3.Visible = True
        ProgressBar3.Style = ProgressBarStyle.Marquee
        ProgressBar3.MarqueeAnimationSpeed = 40
        ProgressBar3.Refresh()
        Dim loadedDocument As New PdfLoadedDocument(loadPDF)
        Dim destinationFilePattern As String = outSplitPDF + "_{0}.pdf"
        loadedDocument.FileStructure.IncrementalUpdate = False
        loadedDocument.Compression = PdfCompressionLevel.Best
        Await Task.Run(Sub() loadedDocument.Split(destinationFilePattern))
        loadedDocument.Close(True)
        ProgressBar3.Style = ProgressBarStyle.Blocks
        ProgressBar3.Value = 100
        If File.Exists(outSplitPDF + "_1.pdf") Then
            MessageBoxAdv.Show("Split PDF success !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Process.Start("explorer.exe", String.Format("/n, /e, {0}", outSplitFldr))
        Else
            MessageBoxAdv.Show("Split PDF failed !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
    Private Async Sub CnvPDF(pdfFile As String, cnvFile As String, cnvExt As String)
        ProgressBar4.Visible = True
        ProgressBar4.Style = ProgressBarStyle.Marquee
        ProgressBar4.MarqueeAnimationSpeed = 40
        ProgressBar4.Refresh()
        Dim doc As New Spire.Pdf.PdfDocument()
        doc.LoadFromFile(pdfFile)
        If cnvExt = "DOCX" Then
            Await Task.Run(Sub() doc.SaveToFile(cnvFile, Spire.Pdf.FileFormat.DOCX))
        ElseIf cnvExt = "XLSX" Then
            Await Task.Run(Sub() doc.SaveToFile(cnvFile, Spire.Pdf.FileFormat.XLSX))
        End If
        doc.Close()
        ProgressBar4.Style = ProgressBarStyle.Blocks
        ProgressBar4.Value = 100
        If File.Exists(cnvFile) Then
            MessageBoxAdv.Show("Convert PDF success !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Label22.Text = GetFileSize(cnvFile)
        Else
            MessageBoxAdv.Show("Convert PDF failed !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
End Class
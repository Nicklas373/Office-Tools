Imports System.Drawing.Imaging
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
            pdfSplit = "\" + Path.GetFileNameWithoutExtension(fileDialog.SafeFileName)
            TextBox3.Text = fileDialog.FileName
            Label20.Text = "Total Pages: " & PDFGetPage(TextBox3.Text)
        Else
            TextBox3.Text = ""
        End If
    End Sub
    Private Sub SavePDF_folder_SplitPDF(sender As Object, e As EventArgs) Handles Button12.Click
        If ComboBox1.Text = "Split All" Then
            savefldDialog.InitialDirectory = Environment.SpecialFolder.UserProfile
            savefldDialog.ShowDialog()
        ElseIf ComboBox1.Text = "Custom range" Or ComboBox1.Text = "Fixed range" Then
            saveDialog.DefaultExt = ".pdf"
            saveDialog.Filter = "PDF File | *.pdf"
            saveDialog.Title = "Save PDF File"
            saveDialog.InitialDirectory = Environment.SpecialFolder.UserProfile
            saveDialog.ShowDialog()
        Else
            MessageBoxAdv.Show("Please choose split type first !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
    Private Sub SaveFileDialog_file_SplitPDF(sender As Object, e As EventArgs) Handles Button12.Click
        If ComboBox1.Text = "Split All" Then
            If savefldDialog.SelectedPath.ToString = "" Then
                TextBox5.Text = ""
            Else
                TextBox5.Text = savefldDialog.SelectedPath.ToString
            End If
        ElseIf ComboBox1.Text = "Custom range" Or ComboBox1.Text = "Fixed range" Then
            If saveDialog.FileName.ToString = "" Then
                TextBox5.Text = ""
            Else
                TextBox5.Text = Path.GetFullPath(saveDialog.FileName.ToString)
            End If
        End If
    End Sub
    Private Sub SplitPDF_Button(sender As Object, e As EventArgs) Handles Button11.Click
        If TextBox3.Text = "" Then
            MessageBoxAdv.Show("No PDF file was selected !, please select PDF file first !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If TextBox5.Text = "" Then
                MessageBoxAdv.Show("Destination PDF file location was not selected !, please select destination location first !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If ComboBox1.Text = "Split All" Then
                    If Strings.Right(TextBox5.Text, 4) = ".pdf" Then
                        SplitPDF(TextBox3.Text, Path.GetDirectoryName(TextBox5.Text) + pdfSplit, Path.GetDirectoryName(TextBox5.Text))
                    Else
                        SplitPDF(TextBox3.Text, TextBox5.Text + pdfSplit, TextBox5.Text)
                    End If
                ElseIf ComboBox1.Text = "Custom range" Then
                    If Strings.Right(TextBox5.Text, 4) = ".pdf" Then
                        PDFSplitRange(TextBox3.Text, Convert.ToInt32(TextBox6.Text), Convert.ToInt32(TextBox7.Text), TextBox5.Text)
                    Else
                        MessageBoxAdv.Show("Destination PDF filename was missing !, Please choose button 'Save Location' and fill PDF filename !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                ElseIf ComboBox1.Text = "Fixed range" Then
                    If Strings.Right(TextBox5.Text, 4) = ".pdf" Then
                        PDFSplitRange(TextBox3.Text, Convert.ToInt32(TextBox6.Text), Convert.ToInt32(TextBox6.Text), TextBox5.Text)
                    Else
                        MessageBoxAdv.Show("Destination PDF filename was missing !, Please choose button 'Save Location' and fill PDF filename !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                Else
                    MessageBoxAdv.Show("Please choose split type first !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
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
            TextBox8.Text = ""
        Else
            TextBox8.Text = Path.GetFullPath(fileDialog.FileName.ToString)
            Label24.Text = GetFileSize(TextBox8.Text)
        End If
    End Sub
    Private Sub SaveFileDialog_Cnv_PDF(sender As Object, e As EventArgs) Handles Button16.Click
        saveDialog.DefaultExt = ".docx|.xlsx|.jpeg"
        saveDialog.Filter = "DOCX File | *.docx|XLS File |*.xls|JPEG File | *.jpeg|PNG File | *.png"
        saveDialog.Title = "Save To"
        saveDialog.InitialDirectory = Environment.SpecialFolder.UserProfile
        If saveDialog.ShowDialog() = DialogResult.OK Then
            pdfSplit = "\" + Path.GetFileNameWithoutExtension(saveDialog.FileName)
            TextBox9.Text = saveDialog.FileName
            If Strings.Right(saveDialog.FileName, 4) = "docx" Then
                ComboBox1.Text = "Documents (*.DOCX)"
            ElseIf Strings.Right(saveDialog.FileName, 4) = ".xls" Then
                ComboBox1.Text = "Spreadsheets (*.XLS)"
            ElseIf Strings.Right(saveDialog.FileName, 4) = "jpeg" Then
                ComboBox1.Text = "Image (*.JPEG)"
                CheckBox1.Visible = True
            ElseIf Strings.Right(saveDialog.FileName, 4) = ".png" Then
                ComboBox1.Text = "Image (*.PNG)"
                CheckBox1.Visible = True
            End If
        Else
            TextBox9.Text = ""
        End If
    End Sub
    Private Sub CnvPDF_Button(sender As Object, e As EventArgs) Handles Button15.Click
        If TextBox8.Text = "" Then
            MessageBoxAdv.Show("No PDF file was selected !, please select PDF file first !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If TextBox9.Text = "" Then
                MessageBoxAdv.Show("Destination file location was not selected !, please select destination location first !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If ComboBox1.Text = "" Then
                    MessageBoxAdv.Show("Convert option was not selected !, please select destination location first !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    If ComboBox1.Text = "Documents (*.DOCX)" Then
                        If Strings.Right(TextBox9.Text, 4) = "docx" Then
                            CnvPDF(TextBox8.Text, TextBox9.Text, "DOCX")
                        Else
                            MessageBoxAdv.Show("Saved filename extensions was not same with extract to settings !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            MessageBoxAdv.Show("Current ext: " & Strings.Right(TextBox7.Text, 4) & " Current export settings: docx", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                    ElseIf ComboBox1.Text = "Spreadsheets (*.XLSX)" Then
                        If Strings.Right(TextBox9.Text, 4) = ".xls" Then
                            CnvPDF(TextBox8.Text, TextBox9.Text, "XLS")
                        Else
                            MessageBoxAdv.Show("Saved filename extensions was not same with extract to settings !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            MessageBoxAdv.Show("Current ext: " & Strings.Right(TextBox7.Text, 4) & " Current export settings: xls", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                    ElseIf ComboBox1.Text = "Image (*.JPEG)" Then
                        If Strings.Right(TextBox9.Text, 4) = "jpeg" Then
                            If CheckBox1.Checked Then
                                ExpImagesFrPDF(TextBox8.Text, TextBox9.Text, "JPEG")
                            Else
                                CnvPDF(TextBox8.Text, TextBox9.Text.Substring(0, TextBox9.Text.Length - 5), "JPEG")
                            End If
                            Label22.Visible = False
                            Label23.Visible = False
                        Else
                            MessageBoxAdv.Show("Saved filename extensions was not same with extract to settings !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            MessageBoxAdv.Show("Current ext: " & Strings.Right(TextBox9.Text, 4) & " Current export settings: jpeg", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                    ElseIf ComboBox1.Text = "Image (*.PNG)" Then
                        If Strings.Right(TextBox9.Text, 4) = "png" Then
                            If CheckBox1.Checked Then
                                ExpImagesFrPDF(TextBox8.Text, TextBox9.Text, "PNG")
                            Else
                                CnvPDF(TextBox8.Text, TextBox9.Text.Substring(0, TextBox9.Text.Length - 5), "PNG")
                            End If
                            Label22.Visible = False
                            Label23.Visible = False
                        Else
                            MessageBoxAdv.Show("Saved filename extensions was not same with extract to settings !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            MessageBoxAdv.Show("Current ext: " & Strings.Right(TextBox9.Text, 4) & " Current export settings: png", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "Split All" Then
            Label27.Visible = False
            Label19.Visible = False
            Label20.Visible = False
            TextBox6.Visible = False
            TextBox7.Visible = False
        ElseIf ComboBox1.Text = "Custom range" Then
            Label27.Visible = True
            Label19.Visible = True
            Label20.Visible = True
            TextBox6.Visible = True
            TextBox7.Visible = True
        ElseIf ComboBox1.Text = "Fixed range" Then
            Label27.Visible = True
            Label19.Visible = False
            Label20.Visible = True
            TextBox6.Visible = True
            TextBox7.Visible = False
        End If
    End Sub
    Private Sub startPagePDFSplit_Btn(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles TextBox6.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBoxAdv.Show("Please enter numbers only !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Handled = True
        End If
    End Sub
    Private Sub endPagePDFSplit_Btn(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles TextBox7.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBoxAdv.Show("Please enter numbers only !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        If TextBox3.Text = "" Then
            TextBox6.Text = ""
            MessageBoxAdv.Show("Please select source PDF file before input pages !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Dim sIndex As Integer = Convert.ToInt32(Val(TextBox6.Text))
            Dim tIndex As Integer = PDFGetPage(TextBox3.Text)
            If TextBox6.Text = "" Then

            Else
                If sIndex = 0 Then
                    MessageBoxAdv.Show("Page number start from page 1 !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    TextBox6.Text = "1"
                Else
                    If sIndex > tIndex Then
                        TextBox6.Text = ""
                        MessageBoxAdv.Show("Selected page can not more than total pages !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        If TextBox3.Text = "" Then
            TextBox7.Text = ""
            MessageBoxAdv.Show("Please select source PDF file before input pages !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Dim sIndex As Integer = Convert.ToInt32(Val(TextBox7.Text))
            Dim tIndex As Integer = PDFGetPage(TextBox3.Text)
            If TextBox7.Text = "" Then

            Else
                If sIndex = 0 Then
                    MessageBoxAdv.Show("Page number start from page 1 !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    TextBox7.Text = "1"
                Else
                    If sIndex > tIndex Then
                        TextBox7.Text = ""
                        MessageBoxAdv.Show("Selected page can not more than total pages !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
            Dim pdfViewer = New PDFViewer(TextBox2.Text)
            pdfViewer.Show()
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
            Dim pdfViewer = New PDFViewer(TextBox4.Text)
            pdfViewer.Show()
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
        loadedDocument.FileStructure.IncrementalUpdate = True
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
    Private Async Sub PDFSplitRange(pdfIn As String, sIndex As Integer, eIndex As Integer, pdfOut As String)
        ProgressBar3.Visible = True
        ProgressBar3.Style = ProgressBarStyle.Marquee
        ProgressBar3.MarqueeAnimationSpeed = 40
        ProgressBar3.Refresh()
        Dim loadedDocument As New PdfLoadedDocument(pdfIn)
        Dim document As New PdfDocument()
        document.ImportPageRange(loadedDocument, sIndex - 1, eIndex - 1)
        Await Task.Run(Sub() document.Save(pdfOut))
        loadedDocument.Close(True)
        document.Close(True)
        ProgressBar3.Style = ProgressBarStyle.Blocks
        ProgressBar3.Value = 100
        If File.Exists(pdfOut) Then
            MessageBoxAdv.Show("Split PDF success !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Dim pdfViewer = New PDFViewer(pdfOut)
            pdfViewer.Show()
        Else
            MessageBoxAdv.Show("Split PDF failed !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
    Private Function PDFGetPage(pdfIn As String) As Integer
        Dim value As Integer
        Dim loadedDocument As New PdfLoadedDocument(pdfIn)
        value = loadedDocument.Pages.Count
        loadedDocument.Close(True)
        Return value
    End Function
    Private Sub CnvPDF(pdfFile As String, cnvFile As String, cnvExt As String)
        ProgressBar4.Visible = True
        ProgressBar4.Style = ProgressBarStyle.Marquee
        ProgressBar4.MarqueeAnimationSpeed = 40
        ProgressBar4.Refresh()
        If cnvExt = "DOCX" Then
            Dim f As New SautinSoft.PdfFocus()
            f.OpenPdf(pdfFile)
            If f.PageCount > 0 Then
                f.WordOptions.Format = SautinSoft.PdfFocus.CWordOptions.eWordDocument.Docx
                Dim result As Integer = f.ToWord(cnvFile)
                ProgressBar4.Style = ProgressBarStyle.Blocks
                ProgressBar4.Value = 100
                If result = 0 Then
                    MessageBoxAdv.Show("Convert PDF success !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Label22.Text = GetFileSize(cnvFile)
                Else
                    MessageBoxAdv.Show("Convert PDF failed !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Else
                MessageBoxAdv.Show("Convert PDF failed !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        ElseIf cnvExt = "XLS" Then
            Dim f As New SautinSoft.PdfFocus()
            f.OpenPdf(pdfFile)
            If f.PageCount > 0 Then
                f.ToExcel(cnvFile)
                If File.Exists(cnvFile) Then
                    ProgressBar4.Style = ProgressBarStyle.Blocks
                    ProgressBar4.Value = 100
                    MessageBoxAdv.Show("Convert PDF success !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Label22.Text = GetFileSize(cnvFile)
                Else
                    MessageBoxAdv.Show("Convert PDF failed !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Else
                MessageBoxAdv.Show("Convert PDF failed !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        ElseIf cnvExt = "JPEG" Then
            Dim f As New SautinSoft.PdfFocus
            f.OpenPdf(pdfFile)
            If f.PageCount > 0 Then
                f.ImageOptions.ImageFormat = ImageFormat.Jpeg
                f.ImageOptions.Dpi = 200
                f.ToImage(Path.GetDirectoryName(TextBox7.Text), Path.GetFileNameWithoutExtension(cnvFile))
            End If
            If File.Exists(Path.GetDirectoryName(TextBox7.Text) & "\" & Path.GetFileNameWithoutExtension(cnvFile) & "1.jpg") Then
                ProgressBar4.Style = ProgressBarStyle.Blocks
                ProgressBar4.Value = 100
                MessageBoxAdv.Show("Convert PDF success !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBoxAdv.Show("Convert PDF failed !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        ElseIf cnvExt = "PNG" Then
            Dim f As New SautinSoft.PdfFocus
            f.OpenPdf(pdfFile)
            If f.PageCount > 0 Then
                f.ImageOptions.ImageFormat = ImageFormat.Png
                f.ImageOptions.Dpi = 200
                f.ToImage(Path.GetDirectoryName(TextBox7.Text), Path.GetFileNameWithoutExtension(cnvFile))
            End If
            If File.Exists(Path.GetDirectoryName(TextBox7.Text) & "\" & Path.GetFileNameWithoutExtension(cnvFile) & "1.png") Then
                ProgressBar4.Style = ProgressBarStyle.Blocks
                ProgressBar4.Value = 100
                MessageBoxAdv.Show("Convert PDF success !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBoxAdv.Show("Convert PDF failed !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub
    Private Async Sub ExpImagesFrPDF(pdffile As String, cnvfile As String, cnvExt As String)
        ProgressBar4.Visible = True
        ProgressBar4.Style = ProgressBarStyle.Marquee
        ProgressBar4.MarqueeAnimationSpeed = 40
        ProgressBar4.Refresh()
        Dim ldoc As New PdfLoadedDocument(pdffile)
        Dim loadedPages As PdfLoadedPageCollection = ldoc.Pages
        Dim n As Integer
        For Each lpage As PdfLoadedPage In loadedPages
            n += 1
            Dim img As Image() = lpage.ExtractImages()
            If img Is Nothing Then Continue For
            For Each img1 As Image In img
                If cnvExt = "JPEG" Then
                    Await Task.Run(Sub() img1.Save(cnvfile & "_" & n & cnvExt, ImageFormat.Png))
                ElseIf cnvExt = "PNG" Then
                    Await Task.Run(Sub() img1.Save(cnvfile & "_" & n & cnvExt, ImageFormat.Jpeg))
                End If
            Next img1
        Next lpage
        ldoc.Dispose()
        If File.Exists(cnvfile & "_1" & cnvExt) Then
            ProgressBar4.Style = ProgressBarStyle.Blocks
            ProgressBar4.Value = 100
            MessageBoxAdv.Show("Extract Image success !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            ProgressBar4.Style = ProgressBarStyle.Blocks
            ProgressBar4.Value = 100
            MessageBoxAdv.Show("Extract Image failed !", "Office Tools", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
End Class
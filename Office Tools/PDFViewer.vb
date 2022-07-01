Imports System.IO
Imports Syncfusion.WinForms.Controls

Public Class PDFViewer
    Inherits SfForm

    Private Sub PDF_Viewer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AllowTransparency = False
        Style.TitleBar.IconBackColor = Color.FromArgb(15, 161, 212)
        BackColor = Color.AliceBlue
        Style.TitleBar.TextHorizontalAlignment = HorizontalAlignment.Center
        Style.TitleBar.TextVerticalAlignment = VisualStyles.VerticalAlignment.Center
    End Sub
    Public Sub New(pdfTitle As String)
        InitializeComponent()
        Me.Text = "Office Tools | PDFViewer - " & Path.GetFileName(pdfTitle)
        PdfViewerControl1.Load(pdfTitle)
    End Sub
End Class
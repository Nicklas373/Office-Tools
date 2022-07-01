<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PDFViewer
    Inherits Syncfusion.WinForms.Controls.SfForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim MessageBoxSettings1 As Syncfusion.Windows.Forms.PdfViewer.MessageBoxSettings = New Syncfusion.Windows.Forms.PdfViewer.MessageBoxSettings()
        Dim PdfViewerPrinterSettings1 As Syncfusion.Windows.PdfViewer.PdfViewerPrinterSettings = New Syncfusion.Windows.PdfViewer.PdfViewerPrinterSettings()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PDFViewer))
        Dim TextSearchSettings1 As Syncfusion.Windows.Forms.PdfViewer.TextSearchSettings = New Syncfusion.Windows.Forms.PdfViewer.TextSearchSettings()
        Me.PdfViewerControl1 = New Syncfusion.Windows.Forms.PdfViewer.PdfViewerControl()
        Me.SuspendLayout()
        '
        'PdfViewerControl1
        '
        Me.PdfViewerControl1.CursorMode = Syncfusion.Windows.Forms.PdfViewer.PdfViewerCursorMode.SelectTool
        Me.PdfViewerControl1.EnableContextMenu = True
        Me.PdfViewerControl1.EnableNotificationBar = True
        Me.PdfViewerControl1.HorizontalScrollOffset = 0
        Me.PdfViewerControl1.IsBookmarkEnabled = True
        Me.PdfViewerControl1.IsTextSearchEnabled = True
        Me.PdfViewerControl1.IsTextSelectionEnabled = True
        Me.PdfViewerControl1.Location = New System.Drawing.Point(-9, 5)
        MessageBoxSettings1.EnableNotification = True
        Me.PdfViewerControl1.MessageBoxSettings = MessageBoxSettings1
        Me.PdfViewerControl1.MinimumZoomPercentage = 50
        Me.PdfViewerControl1.Name = "PdfViewerControl1"
        Me.PdfViewerControl1.PageBorderThickness = 1
        PdfViewerPrinterSettings1.Copies = 1
        PdfViewerPrinterSettings1.PageOrientation = Syncfusion.Windows.PdfViewer.PdfViewerPrintOrientation.[Auto]
        PdfViewerPrinterSettings1.PageSize = Syncfusion.Windows.PdfViewer.PdfViewerPrintSize.ActualSize
        PdfViewerPrinterSettings1.PrintLocation = CType(resources.GetObject("PdfViewerPrinterSettings1.PrintLocation"), System.Drawing.PointF)
        PdfViewerPrinterSettings1.ShowPrintStatusDialog = True
        Me.PdfViewerControl1.PrinterSettings = PdfViewerPrinterSettings1
        Me.PdfViewerControl1.ReferencePath = Nothing
        Me.PdfViewerControl1.ScrollDisplacementValue = 0
        Me.PdfViewerControl1.ShowHorizontalScrollBar = True
        Me.PdfViewerControl1.ShowToolBar = True
        Me.PdfViewerControl1.ShowVerticalScrollBar = True
        Me.PdfViewerControl1.Size = New System.Drawing.Size(961, 681)
        Me.PdfViewerControl1.SpaceBetweenPages = 8
        Me.PdfViewerControl1.TabIndex = 0
        Me.PdfViewerControl1.Text = "PdfViewerControl1"
        TextSearchSettings1.CurrentInstanceColor = System.Drawing.Color.FromArgb(CType(CType(127, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(171, Byte), Integer), CType(CType(64, Byte), Integer))
        TextSearchSettings1.HighlightAllInstance = True
        TextSearchSettings1.OtherInstanceColor = System.Drawing.Color.FromArgb(CType(CType(127, Byte), Integer), CType(CType(254, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.PdfViewerControl1.TextSearchSettings = TextSearchSettings1
        Me.PdfViewerControl1.ThemeName = "Office2016Colorful"
        Me.PdfViewerControl1.VerticalScrollOffset = 0
        Me.PdfViewerControl1.VisualStyle = Syncfusion.Windows.Forms.PdfViewer.VisualStyle.Office2016Colorful
        Me.PdfViewerControl1.ZoomMode = Syncfusion.Windows.Forms.PdfViewer.ZoomMode.[Default]
        '
        'PDFViewer
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Inherit
        Me.BackColor = System.Drawing.Color.AliceBlue
        Me.ClientSize = New System.Drawing.Size(944, 681)
        Me.Controls.Add(Me.PdfViewerControl1)
        Me.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "PDFViewer"
        Me.Style.BackColor = System.Drawing.Color.AliceBlue
        Me.Style.MdiChild.IconHorizontalAlignment = System.Windows.Forms.HorizontalAlignment.Center
        Me.Style.MdiChild.IconVerticalAlignment = System.Windows.Forms.VisualStyles.VerticalAlignment.Center
        Me.Text = "Office Tools | PDFViewer"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents PdfViewerControl1 As Syncfusion.Windows.Forms.PdfViewer.PdfViewerControl
End Class

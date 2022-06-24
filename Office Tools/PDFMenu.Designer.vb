<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PDFMenu
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PDFMenu))
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.CheckBox4 = New System.Windows.Forms.CheckBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.CheckBox5 = New System.Windows.Forms.CheckBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.CheckBox3 = New System.Windows.Forms.CheckBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.pdf_com_pnl = New System.Windows.Forms.Panel()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Panel1.SuspendLayout()
        Me.pdf_com_pnl.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button5
        '
        Me.Button5.FlatAppearance.BorderSize = 0
        Me.Button5.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button5.Font = New System.Drawing.Font("Segoe UI Semibold", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.Button5.ForeColor = System.Drawing.Color.AliceBlue
        Me.Button5.Image = CType(resources.GetObject("Button5.Image"), System.Drawing.Image)
        Me.Button5.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button5.Location = New System.Drawing.Point(3, 2)
        Me.Button5.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(228, 43)
        Me.Button5.TabIndex = 5
        Me.Button5.Text = "  PDF Compress"
        Me.Button5.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.Location = New System.Drawing.Point(240, 2)
        Me.Panel2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(706, 425)
        Me.Panel2.TabIndex = 1
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(6, 418)
        Me.ProgressBar1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(694, 22)
        Me.ProgressBar1.TabIndex = 35
        Me.ProgressBar1.Visible = False
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Label10.Font = New System.Drawing.Font("Segoe UI Semibold", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.Label10.ForeColor = System.Drawing.Color.AliceBlue
        Me.Label10.Location = New System.Drawing.Point(177, 341)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(0, 19)
        Me.Label10.TabIndex = 28
        '
        'CheckBox4
        '
        Me.CheckBox4.AutoSize = True
        Me.CheckBox4.Font = New System.Drawing.Font("Segoe UI Semibold", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.CheckBox4.ForeColor = System.Drawing.Color.AliceBlue
        Me.CheckBox4.Location = New System.Drawing.Point(12, 223)
        Me.CheckBox4.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CheckBox4.Name = "CheckBox4"
        Me.CheckBox4.Size = New System.Drawing.Size(141, 23)
        Me.CheckBox4.TabIndex = 28
        Me.CheckBox4.Text = "Remove Metadata"
        Me.CheckBox4.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Label9.Font = New System.Drawing.Font("Segoe UI Semibold", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.Label9.ForeColor = System.Drawing.Color.AliceBlue
        Me.Label9.Location = New System.Drawing.Point(13, 341)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(115, 19)
        Me.Label9.TabIndex = 27
        Me.Label9.Text = "Compressed Size"
        '
        'ComboBox2
        '
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Items.AddRange(New Object() {"Highest", "High", "Normal", "Low", "Lowest"})
        Me.ComboBox2.Location = New System.Drawing.Point(177, 88)
        Me.ComboBox2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(155, 25)
        Me.ComboBox2.TabIndex = 24
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Label8.Font = New System.Drawing.Font("Segoe UI Semibold", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.Label8.ForeColor = System.Drawing.Color.AliceBlue
        Me.Label8.Location = New System.Drawing.Point(177, 310)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(0, 19)
        Me.Label8.TabIndex = 26
        '
        'CheckBox5
        '
        Me.CheckBox5.AutoSize = True
        Me.CheckBox5.Font = New System.Drawing.Font("Segoe UI Semibold", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.CheckBox5.ForeColor = System.Drawing.Color.AliceBlue
        Me.CheckBox5.Location = New System.Drawing.Point(12, 149)
        Me.CheckBox5.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CheckBox5.Name = "CheckBox5"
        Me.CheckBox5.Size = New System.Drawing.Size(232, 23)
        Me.CheckBox5.TabIndex = 29
        Me.CheckBox5.Text = "Enable Incremental Compression"
        Me.CheckBox5.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Label2.Font = New System.Drawing.Font("Segoe UI Semibold", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.Label2.ForeColor = System.Drawing.Color.AliceBlue
        Me.Label2.Location = New System.Drawing.Point(13, 310)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(90, 19)
        Me.Label2.TabIndex = 25
        Me.Label2.Text = "Original Size"
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.Font = New System.Drawing.Font("Segoe UI Semibold", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.CheckBox3.ForeColor = System.Drawing.Color.AliceBlue
        Me.CheckBox3.Location = New System.Drawing.Point(12, 199)
        Me.CheckBox3.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(180, 23)
        Me.CheckBox3.TabIndex = 27
        Me.CheckBox3.Text = "Optimize Page Contents"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Label4.Font = New System.Drawing.Font("Segoe UI Semibold", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.Label4.ForeColor = System.Drawing.Color.AliceBlue
        Me.Label4.Location = New System.Drawing.Point(7, 92)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(155, 19)
        Me.Label4.TabIndex = 22
        Me.Label4.Text = "PDF Compression Level"
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Font = New System.Drawing.Font("Segoe UI Semibold", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.CheckBox2.ForeColor = System.Drawing.Color.AliceBlue
        Me.CheckBox2.Location = New System.Drawing.Point(12, 174)
        Me.CheckBox2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(124, 23)
        Me.CheckBox2.TabIndex = 26
        Me.CheckBox2.Text = "Optimize Fonts"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.SteelBlue
        Me.Panel1.Controls.Add(Me.Button5)
        Me.Panel1.Controls.Add(Me.Button4)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.ImeMode = System.Windows.Forms.ImeMode.Off
        Me.Panel1.Location = New System.Drawing.Point(-1, -3)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(234, 512)
        Me.Panel1.TabIndex = 7
        '
        'Button4
        '
        Me.Button4.FlatAppearance.BorderSize = 0
        Me.Button4.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button4.Font = New System.Drawing.Font("Segoe UI Semibold", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.Button4.ForeColor = System.Drawing.Color.AliceBlue
        Me.Button4.Image = CType(resources.GetObject("Button4.Image"), System.Drawing.Image)
        Me.Button4.Location = New System.Drawing.Point(0, 469)
        Me.Button4.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(234, 42)
        Me.Button4.TabIndex = 4
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.BackColor = System.Drawing.Color.SteelBlue
        Me.Button6.FlatAppearance.BorderSize = 0
        Me.Button6.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button6.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.Button6.ForeColor = System.Drawing.Color.AliceBlue
        Me.Button6.Image = CType(resources.GetObject("Button6.Image"), System.Drawing.Image)
        Me.Button6.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button6.Location = New System.Drawing.Point(294, 446)
        Me.Button6.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(135, 45)
        Me.Button6.TabIndex = 34
        Me.Button6.Text = "  Compress"
        Me.Button6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button6.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.Button6.UseVisualStyleBackColor = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Label6.Font = New System.Drawing.Font("Segoe UI Semibold", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.Label6.ForeColor = System.Drawing.Color.AliceBlue
        Me.Label6.Location = New System.Drawing.Point(7, 52)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(96, 19)
        Me.Label6.TabIndex = 18
        Me.Label6.Text = "Save Location"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Label5.Font = New System.Drawing.Font("Segoe UI Semibold", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.Label5.ForeColor = System.Drawing.Color.AliceBlue
        Me.Label5.Location = New System.Drawing.Point(7, 14)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(81, 19)
        Me.Label5.TabIndex = 17
        Me.Label5.Text = "Source PDF"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(177, 15)
        Me.TextBox1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.TextBox1.Size = New System.Drawing.Size(320, 25)
        Me.TextBox1.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.Label3.Location = New System.Drawing.Point(531, 28)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(0, 15)
        Me.Label3.TabIndex = 20
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(177, 52)
        Me.TextBox2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.TextBox2.Size = New System.Drawing.Size(320, 25)
        Me.TextBox2.TabIndex = 4
        '
        'Button7
        '
        Me.Button7.BackColor = System.Drawing.Color.Transparent
        Me.Button7.BackgroundImage = CType(resources.GetObject("Button7.BackgroundImage"), System.Drawing.Image)
        Me.Button7.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.Button7.FlatAppearance.BorderColor = System.Drawing.Color.LightSteelBlue
        Me.Button7.FlatAppearance.BorderSize = 0
        Me.Button7.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button7.Location = New System.Drawing.Point(503, 49)
        Me.Button7.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(42, 34)
        Me.Button7.TabIndex = 3
        Me.Button7.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.Label1.Location = New System.Drawing.Point(531, 74)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(0, 15)
        Me.Label1.TabIndex = 10
        '
        'Button8
        '
        Me.Button8.BackColor = System.Drawing.Color.Transparent
        Me.Button8.BackgroundImage = CType(resources.GetObject("Button8.BackgroundImage"), System.Drawing.Image)
        Me.Button8.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.Button8.FlatAppearance.BorderColor = System.Drawing.Color.LightSteelBlue
        Me.Button8.FlatAppearance.BorderSize = 0
        Me.Button8.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button8.Location = New System.Drawing.Point(503, 11)
        Me.Button8.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(42, 34)
        Me.Button8.TabIndex = 0
        Me.Button8.UseVisualStyleBackColor = False
        '
        'pdf_com_pnl
        '
        Me.pdf_com_pnl.Controls.Add(Me.ProgressBar1)
        Me.pdf_com_pnl.Controls.Add(Me.Label10)
        Me.pdf_com_pnl.Controls.Add(Me.CheckBox4)
        Me.pdf_com_pnl.Controls.Add(Me.Label9)
        Me.pdf_com_pnl.Controls.Add(Me.ComboBox2)
        Me.pdf_com_pnl.Controls.Add(Me.Label8)
        Me.pdf_com_pnl.Controls.Add(Me.CheckBox5)
        Me.pdf_com_pnl.Controls.Add(Me.Label2)
        Me.pdf_com_pnl.Controls.Add(Me.CheckBox3)
        Me.pdf_com_pnl.Controls.Add(Me.Label4)
        Me.pdf_com_pnl.Controls.Add(Me.CheckBox2)
        Me.pdf_com_pnl.Controls.Add(Me.Button6)
        Me.pdf_com_pnl.Controls.Add(Me.Label7)
        Me.pdf_com_pnl.Controls.Add(Me.Label6)
        Me.pdf_com_pnl.Controls.Add(Me.Label5)
        Me.pdf_com_pnl.Controls.Add(Me.TextBox1)
        Me.pdf_com_pnl.Controls.Add(Me.Label3)
        Me.pdf_com_pnl.Controls.Add(Me.TextBox2)
        Me.pdf_com_pnl.Controls.Add(Me.Button7)
        Me.pdf_com_pnl.Controls.Add(Me.Label1)
        Me.pdf_com_pnl.Controls.Add(Me.Button8)
        Me.pdf_com_pnl.Location = New System.Drawing.Point(3, 0)
        Me.pdf_com_pnl.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.pdf_com_pnl.Name = "pdf_com_pnl"
        Me.pdf_com_pnl.Size = New System.Drawing.Size(713, 504)
        Me.pdf_com_pnl.TabIndex = 0
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.Label7.Location = New System.Drawing.Point(531, 69)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(0, 15)
        Me.Label7.TabIndex = 19
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.DimGray
        Me.Panel3.Controls.Add(Me.pdf_com_pnl)
        Me.Panel3.Location = New System.Drawing.Point(227, -3)
        Me.Panel3.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(728, 512)
        Me.Panel3.TabIndex = 8
        '
        'PDFMenu
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Inherit
        Me.AutoSize = True
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.BackColor = System.Drawing.Color.AliceBlue
        Me.ClientSize = New System.Drawing.Size(944, 501)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel3)
        Me.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "PDFMenu"
        Me.Style.BackColor = System.Drawing.Color.AliceBlue
        Me.Style.MdiChild.IconHorizontalAlignment = System.Windows.Forms.HorizontalAlignment.Center
        Me.Style.MdiChild.IconVerticalAlignment = System.Windows.Forms.VisualStyles.VerticalAlignment.Center
        Me.Text = "PDF Menu | Office Tools"
        Me.Panel1.ResumeLayout(False)
        Me.pdf_com_pnl.ResumeLayout(False)
        Me.pdf_com_pnl.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Button5 As Button
    Friend WithEvents Panel2 As Panel
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents Label10 As Label
    Friend WithEvents CheckBox4 As CheckBox
    Friend WithEvents Label9 As Label
    Friend WithEvents ComboBox2 As ComboBox
    Friend WithEvents Label8 As Label
    Friend WithEvents CheckBox5 As CheckBox
    Friend WithEvents Label2 As Label
    Friend WithEvents CheckBox3 As CheckBox
    Friend WithEvents Label4 As Label
    Friend WithEvents CheckBox2 As CheckBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Button4 As Button
    Friend WithEvents Button6 As Button
    Friend WithEvents Label6 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents Button7 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Button8 As Button
    Friend WithEvents pdf_com_pnl As Panel
    Friend WithEvents Label7 As Label
    Friend WithEvents Panel3 As Panel
End Class

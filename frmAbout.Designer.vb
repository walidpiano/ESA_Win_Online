<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAbout
    Inherits DevExpress.XtraSplashScreen.SplashScreen

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
        Me.meData = New DevExpress.XtraEditors.MemoEdit()
        Me.HyperlinkLabelControl1 = New DevExpress.XtraEditors.HyperlinkLabelControl()
        Me.pbLogo = New System.Windows.Forms.PictureBox()
        CType(Me.meData.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pbLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'meData
        '
        Me.meData.Cursor = System.Windows.Forms.Cursors.Default
        Me.meData.EditValue = "This application is copyrighted to ESA 2018." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "The application was developed by: W" & _
    "alid Zakaria." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Phone:  +012 292 77250" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "E-mail:  walidpianooo@gmail.com"
        Me.meData.Location = New System.Drawing.Point(293, 14)
        Me.meData.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.meData.Name = "meData"
        Me.meData.Properties.AllowFocused = False
        Me.meData.Properties.Appearance.BackColor = System.Drawing.Color.Transparent
        Me.meData.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.meData.Properties.Appearance.Options.UseBackColor = True
        Me.meData.Properties.Appearance.Options.UseFont = True
        Me.meData.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.meData.Properties.ReadOnly = True
        Me.meData.Properties.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.meData.Size = New System.Drawing.Size(337, 92)
        Me.meData.TabIndex = 14
        Me.meData.TabStop = False
        '
        'HyperlinkLabelControl1
        '
        Me.HyperlinkLabelControl1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.HyperlinkLabelControl1.Location = New System.Drawing.Point(12, 114)
        Me.HyperlinkLabelControl1.Name = "HyperlinkLabelControl1"
        Me.HyperlinkLabelControl1.Size = New System.Drawing.Size(119, 13)
        Me.HyperlinkLabelControl1.TabIndex = 16
        Me.HyperlinkLabelControl1.Text = "walidpianooo@gmail.com"
        '
        'pbLogo
        '
        Me.pbLogo.Image = Global.ESA_Online.My.Resources.Resources.ESASplash
        Me.pbLogo.Location = New System.Drawing.Point(12, 14)
        Me.pbLogo.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.pbLogo.Name = "pbLogo"
        Me.pbLogo.Size = New System.Drawing.Size(273, 92)
        Me.pbLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pbLogo.TabIndex = 15
        Me.pbLogo.TabStop = False
        '
        'frmAbout
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(653, 134)
        Me.ControlBox = True
        Me.Controls.Add(Me.HyperlinkLabelControl1)
        Me.Controls.Add(Me.pbLogo)
        Me.Controls.Add(Me.meData)
        Me.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "frmAbout"
        Me.Text = "Form1"
        CType(Me.meData.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pbLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents meData As DevExpress.XtraEditors.MemoEdit
    Friend WithEvents pbLogo As System.Windows.Forms.PictureBox
    Friend WithEvents HyperlinkLabelControl1 As DevExpress.XtraEditors.HyperlinkLabelControl
End Class

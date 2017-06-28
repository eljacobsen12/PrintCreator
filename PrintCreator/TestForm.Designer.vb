<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class TestForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.txtString = New System.Windows.Forms.TextBox()
        Me.btnExecute = New System.Windows.Forms.Button()
        Me.TestPictureBox = New System.Windows.Forms.PictureBox()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.lblFont = New System.Windows.Forms.Label()
        CType(Me.TestPictureBox, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtString
        '
        Me.txtString.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtString.Location = New System.Drawing.Point(102, 479)
        Me.txtString.Name = "txtString"
        Me.txtString.Size = New System.Drawing.Size(375, 29)
        Me.txtString.TabIndex = 0
        '
        'btnExecute
        '
        Me.btnExecute.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExecute.Location = New System.Drawing.Point(213, 527)
        Me.btnExecute.Name = "btnExecute"
        Me.btnExecute.Size = New System.Drawing.Size(156, 43)
        Me.btnExecute.TabIndex = 1
        Me.btnExecute.Text = "Execute"
        Me.btnExecute.UseVisualStyleBackColor = True
        '
        'TestPictureBox
        '
        Me.TestPictureBox.Location = New System.Drawing.Point(92, 49)
        Me.TestPictureBox.Name = "TestPictureBox"
        Me.TestPictureBox.Size = New System.Drawing.Size(797, 400)
        Me.TestPictureBox.TabIndex = 2
        Me.TestPictureBox.TabStop = False
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(538, 485)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(238, 24)
        Me.ComboBox1.TabIndex = 7
        '
        'lblFont
        '
        Me.lblFont.AutoSize = True
        Me.lblFont.Location = New System.Drawing.Point(499, 488)
        Me.lblFont.Name = "lblFont"
        Me.lblFont.Size = New System.Drawing.Size(34, 16)
        Me.lblFont.TabIndex = 8
        Me.lblFont.Text = "Font"
        '
        'TestForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(968, 590)
        Me.Controls.Add(Me.lblFont)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.TestPictureBox)
        Me.Controls.Add(Me.btnExecute)
        Me.Controls.Add(Me.txtString)
        Me.Name = "TestForm"
        Me.Text = "TestForm"
        CType(Me.TestPictureBox, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtString As TextBox
    Friend WithEvents btnExecute As Button
    Friend WithEvents TestPictureBox As PictureBox
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents lblFont As Label
End Class

Public Class CustomRadioButton
    Inherits RadioButton
    Public Sub New()
        Dim myFont As System.Drawing.Font
        myFont = New Font("Lucida Bright", 18, FontStyle.Bold)
        Me.Font = myFont
        Me.ForeColor = Color.Black
        Me.AutoSize = True
    End Sub
End Class

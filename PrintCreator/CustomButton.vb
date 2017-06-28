Public Class CustomButton
    Inherits Button
    Public Sub New()
        Dim myFont As System.Drawing.Font
        myFont = New Font("Lucida Bright", 20)
        Me.Font = myFont
        Me.ForeColor = Color.Black
        Me.Width = 117
        Me.Height = 43
        Me.AutoSize = False
    End Sub

End Class

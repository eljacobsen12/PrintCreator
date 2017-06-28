
Class TestForm

    'Dim myBit1 As New Bitmap(My.Resources.Album_A_Tribe_Called_Quest_Midnight_Marauders)

    Private Sub btnExecute_Click(sender As Object, e As EventArgs) Handles btnExecute.Click

        'Dim image1 As Image = Image.FromFile(getImagePath("Desktop"))
        Dim image1 As Image = Image.FromFile("E:\EJ\MyPrograms\PrintCreator\Images\Stock Photos\Album-Blank-Blank.png")
        image1 = resizeBitmapPercent(image1, percShrink)
        image1 = rotateBitmap(image1, 10)
        Dim image2 As Image = Nothing
        image2 = funDrawImageOnImage(image1, TestPictureBox.Image, New Point(0, 0))
        'image2 = addImageBackground(image1, Color.Blue, 5)
        TestPictureBox.Image = image2
        TestPictureBox.Invalidate()
    End Sub

    Private Sub TestForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TestPictureBox.Image = fitImageToPictureBox(Image.FromFile(getImagePath("Desktop")), TestPictureBox)
        BindCombo()
    End Sub

    Private Sub ComboBox1_DrawItem(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles ComboBox1.DrawItem
        e.DrawBackground()
        If (e.State And DrawItemState.Focus) <> 0 Then
            e.DrawFocusRectangle()
        End If
        Dim objBrush As Brush = Nothing
        Try
            objBrush = New SolidBrush(e.ForeColor)
            Dim _FontName As String = ComboBox1.Items(e.Index)
            Dim _font As Font = Nothing
            Dim _fontfamily = New FontFamily(_FontName)
            If _fontfamily.IsStyleAvailable(FontStyle.Regular) Then
                _font = New Font(_fontfamily, 14, FontStyle.Regular)
            ElseIf _fontfamily.IsStyleAvailable(FontStyle.Bold) Then
                _font = New Font(_fontfamily, 14, FontStyle.Bold)
            ElseIf _fontfamily.IsStyleAvailable(FontStyle.Italic) Then
                _font = New Font(_fontfamily, 14, FontStyle.Italic)
            End If
            e.Graphics.DrawString(_FontName, _font, objBrush, e.Bounds)
        Finally
            If objBrush IsNot Nothing Then
                objBrush.Dispose()
            End If
            objBrush = Nothing
        End Try
        e.Graphics.Dispose()
    End Sub

    Private Sub BindCombo()
        ComboBox1.DrawMode = DrawMode.OwnerDrawFixed
        ComboBox1.Font = New Font("Microsoft Sans Serif, 11.25pt", 11.25)
        ComboBox1.ItemHeight = 20
        Dim objFontFamily As FontFamily
        Dim objFontCollection As System.Drawing.Text.FontCollection
        Dim tempFont As Font = Nothing
        objFontCollection = New System.Drawing.Text.InstalledFontCollection()
        For Each objFontFamily In objFontCollection.Families
            ComboBox1.Items.Add(objFontFamily.Name)
        Next
    End Sub
End Class
Public Class ImageSplicingExample1

    'Made the Function To take an array Of the Bitmap images To splice
    'And the second parameter To take 1 Of the 4 (StripDirection) Enum values (Horizontal, Vertical, TopLeftToBottomRight, And BottomLeftToTopRight)
    'To Set the direction Of the splices. It will Return a Bitmap image Of the spliced images.

    Private Imgs As New List(Of Bitmap)

    Public Enum StripDirection As Integer
        Horizontal = 0
        Vertical = 1
        TopLeftToBottomRight = 2
        BottomLeftToTopRight = 3
    End Enum

    Private Sub ImageSplicingExample_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For Each fn As String In IO.Directory.GetFiles("E:\EJ\MyPrograms\PrintCreator\Images\Albums", "*.jpg")
            Imgs.Add(New Bitmap(fn))
        Next
        If Imgs(0) IsNot Nothing AndAlso Imgs(1) IsNot Nothing AndAlso Imgs(2) IsNot Nothing AndAlso Imgs(3) IsNot Nothing AndAlso Imgs(4) IsNot Nothing Then
            PictureBox2.Image = resizeImage(Imgs(0), 158, 122)
            PictureBox3.Image = resizeImage(Imgs(1), 158, 122)
            PictureBox4.Image = resizeImage(Imgs(2), 158, 122)
            PictureBox5.Image = resizeImage(Imgs(3), 158, 122)
            PictureBox6.Image = resizeImage(Imgs(4), 158, 122)
            PictureBox1.Image = CreateStripImage(Imgs.ToArray, StripDirection.TopLeftToBottomRight)
        End If
    End Sub

    Private Function CreateStripImage(ByVal images As Bitmap(), ByVal direction As StripDirection) As Bitmap
        Dim ttx As Single = CSng(images(0).Width / 2)
        Dim tty As Single = CSng(images(0).Height / 2)
        Dim stpx As Single = CSng((ttx / images.Count) * 2)
        Dim sx As Single = CSng(If(direction = StripDirection.Horizontal, 0, stpx))
        Dim stpy As Single = CSng((tty / images.Count) * 2)
        Dim sy As Single = CSng(If(direction = StripDirection.Vertical, 0, stpy))
        Dim cbm As New Bitmap(images(0).Width, images(0).Height)
        Dim pnts() As PointF = {}
        Using g As Graphics = Graphics.FromImage(cbm)
            g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            For i As Integer = 0 To images.Count - 1
                Using bm As New Bitmap(images(i), CInt(cbm.Width + sx), CInt(cbm.Height + sy))
                    Using gp As New Drawing2D.GraphicsPath
                        If direction = StripDirection.Horizontal Then
                            g.TranslateTransform(1, -(tty) + (stpy * i))
                            pnts = New PointF() {New PointF(0, CSng((bm.Height / 2) - (stpy / 2))), New PointF(bm.Width, CSng((bm.Height / 2) - (stpy / 2))), New PointF(bm.Width, CSng((bm.Height / 2) + (stpy / 2))), New PointF(0, CSng((bm.Height / 2) + (stpy / 2)))}
                        ElseIf direction = StripDirection.Vertical Then
                            g.TranslateTransform(-(ttx) + (stpx * i), 1)
                            pnts = New PointF() {New PointF(CSng((bm.Width / 2) - (stpx / 2)), 0), New PointF(CSng((bm.Width / 2) - (stpx / 2)), bm.Height), New PointF(CSng((bm.Width / 2) + (stpx / 2)), bm.Height), New PointF(CSng((bm.Width / 2) + (stpx / 2)), 0)}
                        ElseIf direction = StripDirection.BottomLeftToTopRight Then
                            g.TranslateTransform(-ttx + (stpx * i), -tty + (stpy * i))
                            pnts = New PointF() {New PointF(bm.Width - stpx, 0), New PointF(bm.Width, 0), New PointF(bm.Width, stpy), New PointF(stpx, bm.Height), New PointF(0, bm.Height), New PointF(0, bm.Height - stpy)}
                        Else
                            g.TranslateTransform(-ttx + (stpx * i), tty - (stpy * (i + 1)))
                            pnts = New PointF() {New PointF(0, 0), New PointF(stpx, 0), New PointF(bm.Width, bm.Height - stpy), New PointF(bm.Width, bm.Height), New PointF(bm.Width - stpx, bm.Height), New PointF(0, stpy)}
                        End If
                        gp.AddPolygon(pnts)
                        Using tb As New TextureBrush(bm)
                            g.FillPath(tb, gp)
                        End Using
                    End Using
                End Using
                g.ResetTransform()
            Next
        End Using
        Return cbm
    End Function
End Class
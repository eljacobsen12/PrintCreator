Imports System.Drawing.Drawing2D

Public Class ImageSplicingExample2
    Private path1 As New Drawing2D.GraphicsPath
    Private SliceWidth As Integer = 150
        Private OffsetX As Integer = -350
        Dim MouseDownnType As Integer = -1
        Dim MouseDownX, MouseDownY, MouseDownBmpX, MouseDownBmpY As Integer

        Structure bmpTom
            Public x As Integer
            Public y As Integer
            Public bmp As Bitmap
        End Structure
        Private bmpList As New List(Of bmpTom)

    Private Sub ImageSplicingExample2_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
        'determine which slice was clicked
        Using path2 As Drawing2D.GraphicsPath = CType(path1.Clone, GraphicsPath)
            MouseDownnType = -1

            Dim Matrix1, matrix2 As New Matrix
            Matrix1.Translate(OffsetX, 0)
            path2.Transform(Matrix1)
            matrix2.Translate(SliceWidth, 0)

            For i = 0 To 3
                path2.Transform(matrix2)
                If path2.IsVisible(e.X, e.Y) Then
                    MouseDownnType = i
                    MouseDownX = e.X
                    MouseDownY = e.Y
                    MouseDownBmpX = bmpList(MouseDownnType).x
                    MouseDownBmpY = bmpList(MouseDownnType).y
                    Exit Sub
                End If
            Next
        End Using
    End Sub

    Private Sub ImageSplicingExample2_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        If MouseDownnType >= 0 Then
            'move the image if dragging
            Dim thisbmptom As New bmpTom
            thisbmptom = bmpList(MouseDownnType)
            thisbmptom.x = MouseDownBmpX - (MouseDownX - e.X)
            thisbmptom.y = MouseDownBmpY - (MouseDownY - e.Y)
            bmpList(MouseDownnType) = thisbmptom
            thisbmptom = Nothing
            Me.Invalidate()
        End If
    End Sub

    Private Sub Form3_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp
            MouseDownnType = -1
        End Sub

    Private Sub ImageSplicingExample2_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint

        Dim B As New Bitmap(Me.ClientRectangle.Width, Me.ClientRectangle.Height + 5)
        Using G As Graphics = Graphics.FromImage(B)
            With G
                .TranslateTransform(OffsetX, 0)
                For i = 0 To 3
                    .TranslateTransform(SliceWidth, 0)
                    .SetClip(path1)
                    .DrawImage(bmpList(i).bmp, bmpList(i).x, bmpList(i).y, bmpList(0).bmp.Width, bmpList(0).bmp.Height)
                Next
            End With
        End Using

        Dim TopSliceOffset As Integer = OffsetX + 350 \ 2 + SliceWidth
        Dim BotSliceOffset As Double = OffsetX + 350 / 2 + 28
        Dim OffsetPerPixel As Double = BotSliceOffset / bmpList(0).bmp.Height
        For S As Integer = 0 To 2
            TopSliceOffset += SliceWidth
            For x = 1 To Me.ClientRectangle.Height - 1
                For y = -2 To 2
                    Dim Gc As Integer = 0
                    Dim Rc As Integer = 0
                    Dim Bc As Integer = 0
                    For z = -1 To 1
                        For zz = -2 To 2
                            Dim H As Integer = CInt(IIf(x + y + zz < 0, x, x + y + zz))
                            Dim Col As Color = B.GetPixel(TopSliceOffset + CInt(x * OffsetPerPixel) + y + z, H)
                            Gc += Col.G
                            Rc += Col.R
                            Bc += Col.B
                        Next
                        B.SetPixel(TopSliceOffset + CInt(x * OffsetPerPixel) + y, x, Color.FromArgb(255, Rc \ 15, Gc \ 15, Bc \ 15))
                    Next
                Next
            Next
        Next

        e.Graphics.DrawImage(B, New Point(0, 0))
        e.Graphics.Dispose()
    End Sub

    Private Sub ImageSplicingExample2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.DoubleBuffered = True
        Me.Height = 450
        Me.Width = 450
        'load the images
        Dim thisbmpTom As New bmpTom
        thisbmpTom.bmp = New Bitmap("E:\EJ\MyPrograms\PrintCreator\Images\Albums\Album-Joey_Badass-All_Amerikkkan_Badass.JPG")
        bmpList.Add(thisbmpTom)
        thisbmpTom.bmp = New Bitmap("E:\EJ\MyPrograms\PrintCreator\Images\Albums\Album-Notorious_B.I.G.-ReadY_To_Die.JPG")
        bmpList.Add(thisbmpTom)
        thisbmpTom.bmp = New Bitmap("E:\EJ\MyPrograms\PrintCreator\Images\Albums\Album-Kendrick_Lamar-Good.Kid.Mad.City.JPG")
        bmpList.Add(thisbmpTom)
        thisbmpTom.bmp = New Bitmap("E:\EJ\MyPrograms\PrintCreator\Images\Albums\Album-Outkast-Stankonia.JPG")
        bmpList.Add(thisbmpTom)

        'create the path
        Dim w1 As Single = 350 'bmpList(0).bmp.Width
        Dim h1 As Single = bmpList(0).bmp.Height
        Dim thePolygon() As PointF = {New PointF(w1 / 2, 0), New PointF((w1 / 2) + SliceWidth, 0), New PointF(w1 / 2, h1), New PointF((w1 / 2) - SliceWidth, h1), New PointF(w1 / 2, 0)}
        path1.AddLines(thePolygon)

    End Sub

    Public Shared Function AlphaBlend(ByVal ForeGround As Color, ByVal BackGround As Color) As Color
        If ForeGround.A = 0 Then Return BackGround
        If BackGround.A = 0 Then Return ForeGround
        If ForeGround.A = 255 Then Return ForeGround
        Dim Alpha As Integer = CInt(ForeGround.A) + 1
        Dim B As Integer = Alpha * ForeGround.B + (255 - Alpha) * BackGround.B >> 8
        Dim G As Integer = Alpha * ForeGround.G + (255 - Alpha) * BackGround.G >> 8
        Dim R As Integer = Alpha * ForeGround.R + (255 - Alpha) * BackGround.R >> 8
        Dim A As Integer = ForeGround.A
        If BackGround.A = 255 Then A = 255
        If A > 255 Then A = 255
        If R > 255 Then R = 255
        If G > 255 Then G = 255
        If B > 255 Then B = 255
        Return Color.FromArgb(Math.Abs(A), Math.Abs(R), Math.Abs(G), Math.Abs(B))
    End Function
End Class
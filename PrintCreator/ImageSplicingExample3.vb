Public Class ImageSplicingExample3
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Height = 400
    End Sub

    Private Sub ImageSplicingExample3_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint

        Dim BmpWidth As Integer = 100
        Dim BmpHeight As Integer = 500
        Dim Bpp As Integer = 4 '   <-- bytes per pixel
        Dim BmpStride As Integer = BmpWidth * Bpp ' < --- Is a multiple of 4 : no padding needed

        'Get an array for the bytmap data
        Dim Pixels(BmpHeight * BmpWidth) As Integer

        'Get a Ptr that will be used as Scan0
        Dim handle As Runtime.InteropServices.GCHandle
        handle = Runtime.InteropServices.GCHandle.Alloc(Pixels, Runtime.InteropServices.GCHandleType.Pinned)
        Dim Scan0 As IntPtr = handle.AddrOfPinnedObject()

        'Get a bitmap 32bpp - ARGB from the Scan0
        Dim Bmp As New Bitmap(BmpWidth, BmpHeight, BmpStride, Imaging.PixelFormat.Format32bppArgb, Scan0)

        'Get a graphics from the bitmap
        Dim G As Graphics = Graphics.FromImage(Bmp)

        'Draw a polygone on the bitmap
        G.FillPolygon(Brushes.Red, New Point() {New Point(50, 10), New Point(150, 10), New Point(100, 500), New Point(0, 500)})

        'Top-Left corner of the polygone
        Dim x As Integer = 50
        Dim y As Integer = 10

        Dim sw As New Stopwatch
        sw.Start()

        'Draw transparency
        Dim Offset As Integer = BmpWidth * y + x

        For yy As Integer = y To y + 399
            Offset -= 1
            While Pixels(Offset) <> &HFFFF0000
                Offset += 1
            End While
            For i As Integer = 0 To 19
                Dim Alpha As Integer = (CInt(255 / 20) * i) << 24
                Dim Px As Integer = Pixels(Offset + i)
                Dim ClrAlpha As Integer = Px And &HFFFFFF
                Dim NewColor As Integer = ClrAlpha Or Alpha
                Pixels(Offset + i) = NewColor
            Next
            Offset += BmpWidth
        Next

        sw.Stop()
        Me.Text = sw.Elapsed.TotalMilliseconds.ToString & " Milliseconds"

        'draw the bitmap to the form
        e.Graphics.DrawImage(Bmp, 0, 0)

        'Clean up
        handle.Free()
        G.Dispose()
        Bmp.Dispose()
    End Sub
End Class
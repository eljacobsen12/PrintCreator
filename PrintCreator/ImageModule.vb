Imports System.IO

Module ImageModule
    Dim currentLayout As String
    Private userPercent As Decimal
    Property percShrink() As Decimal
        Get
            Return userPercent
        End Get
        Private Set(ByVal value As Decimal)
            userPercent = value
        End Set
    End Property

    '******************************
    '     IMAGE MANIPULATION
    '******************************

    'Create & Draw Bitmap from FileName.
    Public Sub CreateBitmapFromFileNameAndDraw(ByVal e As PaintEventArgs, ByRef fileName As String)
        If fileName IsNot Nothing Then
            Dim myBitmap As New Bitmap(fileName) 'Creates Bitmap.
            Dim ulCorner As New Point(100, 100) 'Creates point for upper-left corner of form.
            e.Graphics.DrawImage(myBitmap, ulCorner) 'Draws image to screen.
            e.Graphics.Dispose()
        Else
            MsgBox("ERROR: createBitmapFromFileNameAndDraw()", Title:="ERROR")
        End If
    End Sub

    'Create & Draw Bitmap From FilePath.
    Public Sub CreateBitmapFromFilePathAndDraw(ByVal e As PaintEventArgs, ByRef fileName As String)
        Dim myBitmap As New Bitmap(fileName) 'Creates Bitmap.
        Dim ulCorner As New Point(100, 100) 'Creates point for upper-left corner of form.      
        e.Graphics.DrawImage(myBitmap, ulCorner) 'Draws image to screen.
        e.Graphics.Dispose()
    End Sub

    'Convert from Image to Bitmap.
    Public Function convertToBitmap(ByRef myImage As Image)
        Dim myBitmap As New Bitmap(myImage, myImage.Width, myImage.Height)
        Dim myGraphics As Graphics = Graphics.FromImage(myBitmap)
        With myGraphics
            .InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic
            .SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            .CompositingQuality = Drawing2D.CompositingQuality.HighQuality
            .DrawImage(myBitmap, 0, 0, myImage.Width, myImage.Height)
            .Dispose()
        End With
        Return myBitmap
    End Function

    'Convert From Bitmap to Image.
    Public Function convertToImage(ByRef myImage As Bitmap)
        Dim myBitmap As New Bitmap(myImage, myImage.Width, myImage.Height)
        Dim myGraphics As Graphics = Graphics.FromImage(myBitmap)
        With myGraphics
            .InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic
            .SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            .CompositingQuality = Drawing2D.CompositingQuality.HighQuality
            .DrawImage(myBitmap, 0, 0, myImage.Width, myImage.Height)
            .Dispose()
        End With
        Return myBitmap
    End Function


    'Resize Image, preserve quality.    //PASSED TESTING
    Public Function resizeImage(ByVal myImage As Image, ByVal width As Integer, ByVal height As Integer) As Bitmap
        Dim newBit As New Bitmap(width, height) 'New blank bitmap
        Dim g As Graphics = Graphics.FromImage(newBit)
        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic 'Change interpolation for reduction quality; .NearestNeighbor for less cpu usage
        g.DrawImage(myImage, 0, 0, width, height)
        g.Dispose()
        Return newBit
    End Function

    'Resize Bitmap, preserve quality.   //PASSED TESTING
    Public Function resizeBitmap(ByVal myImage As Bitmap, ByVal width As Integer, ByVal height As Integer) As Bitmap
        Dim newBit As New Bitmap(width, height) 'new blank bitmap
        Dim g As Graphics = Graphics.FromImage(newBit)
        'change interpolation for reduction quality
        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic    '.NearestNeighbor for less cpu usage
        g.DrawImage(myImage, 0, 0, width, height)
        g.Dispose()
        Return newBit
    End Function

    'Resize Bitmap by Percent, preserve quality.
    Public Function resizeBitmapPercent(ByRef myImage As Bitmap, ByVal perc As Decimal)
        Dim returnImage As Bitmap = Nothing
        Dim percWidth As Integer = myImage.Width * perc
        Dim percHeight As Integer = myImage.Height * perc
        returnImage = resizeBitmap(myImage, percWidth, percHeight)
        Return returnImage
    End Function

    'Resize Image if Image.Width/Height >= GroupBox.Width/Height    //PASSED TESTING
    Public Function fitImageToGroupBox(ByRef myImage As Bitmap, ByRef myGroupBox As GroupBox)
        Dim newWidth As Integer = myImage.Width
        Dim newHeight As Integer = myImage.Height
        Dim wDiff As Decimal = Nothing
        Dim hDiff As Decimal = Nothing
        percShrink = 0

        If myImage Is Nothing Then Return Nothing
        While myImage.Width >= myGroupBox.Width Or myImage.Height >= myGroupBox.Height
            If myImage.Width >= myGroupBox.Width Then
                wDiff = myImage.Width - myGroupBox.Width
            Else
                wDiff = myImage.Width
            End If
            If myImage.Height >= myGroupBox.Height Then
                hDiff = myImage.Height - myGroupBox.Height
            End If
            If wDiff <> Nothing Or hDiff <> Nothing Then
                If wDiff >= hDiff Then
                    'Shrink Image by slightly more than ratio ImageWidth/GroupBoxWidth
                    percShrink = 1 - (Math.Round(wDiff / myImage.Width, 2, MidpointRounding.AwayFromZero) + 0.01)
                Else
                    'Shrink Image by slightly more than ratio ImageHeight/GroupBoxHeight
                    percShrink = 1 - (Math.Round(hDiff / myImage.Height, 2, MidpointRounding.AwayFromZero) + 0.01)
                End If
                newWidth = CInt(myImage.Width * percShrink)
                newHeight = CInt(myImage.Height * percShrink)
            End If
            myImage = resizeBitmap(myImage, newWidth, newHeight)
        End While
        Return myImage
    End Function

    'Resize Image if Image.Width/Height >= to PictureBox.Width/Height   //PASSED TESTING
    Public Function fitImageToPictureBox(ByRef myImage As Bitmap, ByRef myPictureBox As PictureBox)
        Dim newWidth As Integer = myImage.Width
        Dim newHeight As Integer = myImage.Height
        Dim wDiff As Decimal = Nothing
        Dim hDiff As Decimal = Nothing
        percShrink = 0

        If myImage Is Nothing Then Return Nothing
        While myImage.Width >= myPictureBox.Width Or myImage.Height >= myPictureBox.Height
            If myImage.Width >= myPictureBox.Width Then
                wDiff = myImage.Width - myPictureBox.Width
            Else
                wDiff = myImage.Width
            End If
            If myImage.Height >= myPictureBox.Height Then
                hDiff = myImage.Height - myPictureBox.Height
            End If
            If wDiff <> Nothing Or hDiff <> Nothing Then
                If wDiff >= hDiff Then
                    'Shrink Image by slightly more than ratio ImageWidth/GroupBoxWidth
                    percShrink = 1 - (Math.Round(wDiff / myImage.Width, 2, MidpointRounding.AwayFromZero) + 0.01)
                Else
                    'Shrink Image by slightly more than ratio ImageHeight/GroupBoxHeight
                    percShrink = 1 - (Math.Round(hDiff / myImage.Height, 2, MidpointRounding.AwayFromZero) + 0.01)
                End If
                newWidth = CInt(myImage.Width * percShrink)
                newHeight = CInt(myImage.Height * percShrink)
            End If
            myImage = resizeBitmap(myImage, newWidth, newHeight)
        End While
        Return myImage
    End Function

    Public Function fitImageToControl(ByRef myImage As Bitmap, ByRef myControl As Control)
        Dim newWidth As Integer = myImage.Width
        Dim newHeight As Integer = myImage.Height
        Dim wDiff As Decimal = Nothing
        Dim hDiff As Decimal = Nothing
        percShrink = 0

        If myImage Is Nothing Then Return Nothing
        While myImage.Width >= myControl.Width Or myImage.Height >= myControl.Height
            If myImage.Width >= myControl.Width Then
                wDiff = myImage.Width - myControl.Width
            Else
                wDiff = myImage.Width
            End If
            If myImage.Height >= myControl.Height Then
                hDiff = myImage.Height - myControl.Height
            End If
            If wDiff <> Nothing Or hDiff <> Nothing Then
                If wDiff >= hDiff Then
                    'Shrink Image by slightly more than ratio ImageWidth/GroupBoxWidth
                    percShrink = 1 - (Math.Round(wDiff / myImage.Width, 2, MidpointRounding.AwayFromZero) + 0.01)
                Else
                    'Shrink Image by slightly more than ratio ImageHeight/GroupBoxHeight
                    percShrink = 1 - (Math.Round(hDiff / myImage.Height, 2, MidpointRounding.AwayFromZero) + 0.01)
                End If
                newWidth = CInt(myImage.Width * percShrink)
                newHeight = CInt(myImage.Height * percShrink)
            End If
            myImage = resizeBitmap(myImage, newWidth, newHeight)
        End While
        Return myImage
    End Function

    'Overlays Text of specific Font and Color on Image
    Public Sub AddTextToImage(ByRef myImage As Bitmap, ByVal strText As String, ByVal font As Font, ByVal color As Color, ByVal x As Single, ByVal y As Single)
        Dim gr As Graphics = Graphics.FromImage(myImage)
        gr.DrawString(strText, font, New SolidBrush(color), x, y)
        gr.Dispose()
    End Sub

    'Rotate Bitmap.     //PASSED TESTING *added .dispose()
    Public Function rotateBitmap(myBitmap As Bitmap, angle As Single) As Bitmap
        'create a new empty bitmap to hold rotated image
        Dim newSize As Integer = getLargestSide(myBitmap)
        Dim returnBitmap As New Bitmap(newSize, newSize)  'Multiply by 2 so new bitmap is bigger, has room to rotate.
        'make a graphics object from the empty bitmap
        Dim g As Graphics = Graphics.FromImage(returnBitmap)
        'move rotation point to center of image
        g.TranslateTransform(CSng(myBitmap.Width) / 2, CSng(myBitmap.Height) / 2)
        'rotate
        g.RotateTransform(angle)
        'move image back
        g.TranslateTransform(-CSng(myBitmap.Width) / 2, -CSng(myBitmap.Height) / 2)
        'draw passed in image onto graphics object
        Dim centerPoint As Point = Nothing
        centerPoint.X = (returnBitmap.Width / 2) - (myBitmap.Width / 2)
        centerPoint.Y = (returnBitmap.Height / 2) - (myBitmap.Height / 2)
        g.DrawImage(myBitmap, centerPoint)
        g.Dispose()
        Return returnBitmap     'Use Bresenhams Line Algorithm or Simple DDA to get exact shift
    End Function

    'Rotate Image.     //PASSED TESTING *added .dispose()
    Public Function rotateImage(myImage As Image, angle As Single) As Bitmap
        'create a new empty bitmap to hold rotated image
        Dim returnBitmap As New Bitmap(myImage.Width * 2, myImage.Height * 2)   'Multiply by 2 so new bitmap is bigger, has room to rotate.
        'make a graphics object from the empty bitmap
        Dim g As Graphics = Graphics.FromImage(returnBitmap)
        'move rotation point to center of image
        g.TranslateTransform(CSng(myImage.Width) / 2, CSng(myImage.Height) / 2)
        'rotate
        g.RotateTransform(angle)
        'move image back
        g.TranslateTransform(-CSng(myImage.Width) / 2, -CSng(myImage.Height) / 2)
        'draw passed in image onto graphics object
        g.DrawImage(myImage, New Point(0, 0))
        g.Dispose()
        Return returnBitmap     'Use Bresenhams Line Algorithm or Simple DDA to get exact shift
    End Function

    'SUB Draw Image on another Image, specify location.     //PASSED TESTING
    Public Sub SubDrawImageOnImage(ByRef topImage As Bitmap, ByRef backImage As Bitmap, ByVal myPoint As Point)
        Dim myGraphics As Graphics = Graphics.FromImage(backImage)
        With myGraphics
            .InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic
            .SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            .CompositingQuality = Drawing2D.CompositingQuality.HighQuality
            .DrawImage(topImage, myPoint)
            .Dispose()
        End With
    End Sub

    'FUNCTION Draw Image on another Image, specify location.
    Public Function funDrawImageOnImage(ByRef topImage As Bitmap, ByRef backImage As Bitmap, ByVal myPoint As Point)
        Dim returnImage As Image = Nothing
        Dim myGraphics As Graphics = Graphics.FromImage(backImage)
        With myGraphics
            .InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic
            .SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            .CompositingQuality = Drawing2D.CompositingQuality.HighQuality
            .DrawImage(topImage, myPoint)
            .Dispose()
        End With
        returnImage = backImage
        Return returnImage
    End Function

    'Draw part of Image on Same Image in specified Rectangle.
    Public Sub DrawPieceOfImageOnTheImage(ByRef myImage As Bitmap, ByRef destRec As Rectangle, ByRef srcRect As Rectangle, ByRef srcUnit As GraphicsUnit, ByVal e As PaintEventArgs)
        e.Graphics.DrawImage(myImage, destRec, srcRect, srcUnit)
        e.Graphics.Dispose()
    End Sub

    Public Function drawImageOnImage(ByRef fromPic As Bitmap, ByRef fromRec As Rectangle, ByRef toPic As Bitmap, ByRef toRec As Rectangle, ByRef srcUnit As GraphicsUnit)
        '***TEST OUT***
        'Get the Bitmaps and a Graphics object for the destination Bitmap.
        Dim fr_bm As New Bitmap(fromPic)
        Dim to_bm As New Bitmap(toPic)
        Dim gr As Graphics = Graphics.FromImage(to_bm)

        ' Draw from the source to the destination.
        gr.DrawImage(fr_bm, toRec, fromRec, GraphicsUnit.Pixel)
        gr.Dispose()
        Return to_bm
    End Function

    'Copy BitMap from one area to another; BltBit rebuilt.
    Public Sub CopyBitMap(ByRef uSource As Bitmap, ByRef uTarget As Bitmap, ByVal uDestX As Integer, ByVal uDestY As Integer, ByVal uSrcWidth As Integer, ByVal uSrcHeight As Integer, ByVal uSrcX As Integer, ByVal uSrcY As Integer)

        Dim nSrc As New Rectangle
        nSrc = Rectangle.FromLTRB(uSrcX, uSrcY, uSrcX + uSrcWidth, uSrcY + uSrcHeight)

        Dim nDst As New Rectangle
        nDst = Rectangle.FromLTRB(uDestX, uDestY, uDestX + uSrcWidth, uDestY + uSrcHeight)

        Using g As Graphics = Graphics.FromImage(uTarget)
            ' Draw the specified section of the source bitmap to the new one
            g.DrawImage(uSource, nSrc, nDst, GraphicsUnit.Pixel)
        End Using
    End Sub

    'Create QA In GroupBox. //LAYOUT and MATERIAL PASSED TESTING
    '*****NEEDS Images And Details QA*****
    Public Sub createQA(ByRef myGroupBox As GroupBox, ByRef groupName As String)
        Dim myLayouts As String() = {"Top 5 Ordered", "Top 5 Square", "Top 5 Fanned", "Desktop", "Record Center Piece", "Collage", "Spliced"}
        Dim myMaterials As String() = {"Hardwood", "Lit Chestnut Paneling", "Lit Colonial Maple Paneling", "Provincial", "Canadel", "Acacia Variation", "Bamboo Shoots", "Bamboo Weave", "Wood Weave", "Colored Wood"}
        Dim mySizes As String() = {}
        Dim myRB As CustomRadioButton = Nothing
        Dim myLB As ListBox = Nothing
        Select Case groupName
            Case "LAYOUT"
                Dim y As Integer = 90
                For x As Integer = 0 To 6
                    myRB = New CustomRadioButton
                    myRB.Location = New Point(57, y)
                    myRB.Text = myLayouts(x)
                    myRB.Name = "rb" & myLayouts(x)
                    myGroupBox.Controls.Add(myRB)
                    AddHandler myRB.Click, AddressOf HomePage.HandleRadioButtonChange
                    y += 40
                Next
            Case "MATERIAL"
                Dim y As Integer = 60
                For x As Integer = 0 To 9
                    myRB = New CustomRadioButton
                    myRB.Location = New Point(10, y)
                    myRB.Text = myMaterials(x)
                    myRB.Name = "rb" & myMaterials(x)
                    myGroupBox.Controls.Add(myRB)
                    y += 40
                Next
                'Create Back Button
                Dim btnBack As New CustomButton
                myGroupBox.Controls("btnNext").Location = New Point(195, 460)
                btnBack.Location = New Point(45, 460)
                btnBack.Text = "Back"
                btnBack.Name = "btnBack"
                myGroupBox.Controls.Add(btnBack)
            Case "SIZE"
                Dim arrSizes As String() = getSizing(currentLayout)
                myLB = New ListBox
                myLB.Location = New Point(57, 90)
                myLB.DataSource = arrSizes
                myLB.AutoSize = True
                myLB.Name = "lbSizes"
                myGroupBox.Controls.Add(myLB)
            Case "IMAGES"

            Case "DETAILS"
                '*** CREATE FINISH BUTTON ***

            Case Else
                MsgBox("ERROR: createQA()", Title:="ERROR")
        End Select
    End Sub

    Private Function CreateImageList(ByVal FilePath As String)
        'Add Images to ImageList from Directory
        Dim ImageFolderPath As String = FilePath   '"E:\EJ\MyPrograms\PrintCreator\Images\Stock Photos\Panel Layouts\"
        Dim ImagePaths As String() = Directory.GetFiles(ImageFolderPath, "*.png")
        Dim LayoutImages As New ImageList
        If ImagePaths IsNot Nothing Then
            For Each img In ImagePaths
                If img IsNot Nothing Then
                    LayoutImages.Images.Add(Path.GetFileName(img), Image.FromFile(img))
                End If
            Next
        End If
        Return LayoutImages
    End Function

    Public Sub CreateLayoutQA(ByVal DetailsPanel As Panel)
        Dim ImageFolderPath As String = "E:\EJ\MyPrograms\PrintCreator\Images\Stock Photos\Panel Layouts\"
        Dim LayoutImages As ImageList = CreateImageList(ImageFolderPath)
        Dim ImagePaths As String() = Directory.GetFiles(ImageFolderPath, "*.png")

        'Add Panel Label Headers
        For i As Integer = 1 To 5
            Dim MyLabel As New Label
            If i = 1 Then
                MyLabel.Text = i.ToString & " Panel"
            Else
                MyLabel.Text = i.ToString & " Panels"
            End If
            MyLabel.Font = New Font(FontFamily.GenericSansSerif, 16.0F, FontStyle.Bold Or FontStyle.Underline)
            MyLabel.Location = New Point((20 * i), 10)
            DetailsPanel.Controls.Add(MyLabel)
        Next

        'Add Layout Buttons to Details Panel
        'https://stackoverflow.com/questions/33398371/create-list-of-controls


        'NEED TO CHANGE LOOP SO BUTTON LOCATION INCREASES BY INTERVAL (running count?)
        Dim imgName As String = Nothing
        If ImagePaths IsNot Nothing Then
            For Each img In ImagePaths
                If img IsNot Nothing Then
                    imgName = Path.GetFileName(img)
                    Dim MyButton As New Button
                    Select Case imgName.First
                        Case "1"
                            MyButton.Name = imgName
                            MyButton.Size = New Size(20, 20)
                            MyButton.Image = fitImageToControl(LayoutImages.Images(imgName), MyButton)
                    End Select
                End If
            Next
        End If
    End Sub


    'Returns Checked Radio Button in GroupBox. //PASSED TESTING
    Public Function getCheckedButton(ByRef myGroup As GroupBox)
        Dim rButton As RadioButton = myGroup.Controls.OfType(Of RadioButton).FirstOrDefault(Function(r) r.Checked = True)
        If rButton Is Nothing Then Return Nothing
        Return rButton
    End Function

    'Returns Label of Selected QA - Sets currentLayout when called.
    Public Function getSelectedQASection(ByRef myGroup As GroupBox)
        Dim selectedSection As Label = Nothing
        Dim myControl As Label = Nothing
        For Each myControl In myGroup.Controls
            If myControl.ForeColor = Color.ForestGreen AndAlso myControl.Text <> ">>" AndAlso myControl.Font.Underline = True Then
                selectedSection = myControl
            End If
        Next
        currentLayout = selectedSection.Text
        Return selectedSection
    End Function

    'Delete ALL RadioButtons in GroupBox. //PASSED TESTING
    Public Sub DeleteRadioButtons(ByRef myGroup As GroupBox)
        Dim myRB As CustomRadioButton = Nothing
        For Each myRB In myGroup.Controls.OfType(Of CustomRadioButton).Reverse
            myGroup.Controls.Remove(myRB)
        Next
    End Sub

    'Returns Bitmap of Material.
    Public Function getMaterialImage(ByVal strMaterial As String)
        Dim objMaterial As Bitmap = Nothing
        Dim strFilePath As String = Nothing

        Select Case strMaterial
            Case "Hardwood"
                strFilePath = "E:\EJ\MyPrograms\PrintCreator\Images\Desktop Materials\Hardwood.jpg"
                objMaterial = New Bitmap(strFilePath)
            Case "Lit Chestnut Paneling"
                strFilePath = "E:\EJ\MyPrograms\PrintCreator\Images\Desktop Materials\LitChestnutPaneling.jpg"
                objMaterial = New Bitmap(strFilePath)
            Case "Lit Colonial Maple Paneling"
                strFilePath = "E:\EJ\MyPrograms\PrintCreator\Images\Desktop Materials\LitColonialMaplePaneling.jpg"
                objMaterial = New Bitmap(strFilePath)
            Case "Provincial"
                strFilePath = "E:\EJ\MyPrograms\PrintCreator\Images\Desktop Materials\Provincial.jpg"
                objMaterial = New Bitmap(strFilePath)
            Case "Canadel"
                strFilePath = "E:\EJ\MyPrograms\PrintCreator\Images\Desktop Materials\Canadel.jpg"
                objMaterial = New Bitmap(strFilePath)
            Case "Acacia Variation"
                strFilePath = "E:\EJ\MyPrograms\PrintCreator\Images\Desktop Materials\AcaciaVariation.jpg"
                objMaterial = New Bitmap(strFilePath)
            Case "Bamboo Shoots"
                strFilePath = "E:\EJ\MyPrograms\PrintCreator\Images\Desktop Materials\BambooShoots.jpg"
                objMaterial = New Bitmap(strFilePath)
            Case "Bamboo Weave"
                strFilePath = "E:\EJ\MyPrograms\PrintCreator\Images\Desktop Materials\BambooWeave.jpg"
                objMaterial = New Bitmap(strFilePath)
            Case "Wood Weave"
                strFilePath = "E:\EJ\MyPrograms\PrintCreator\Images\Desktop Materials\WoodWeave.jpg"
                objMaterial = New Bitmap(strFilePath)
            Case "Colored Wood"
                strFilePath = "E:\EJ\MyPrograms\PrintCreator\Images\Desktop Materials\ColoredWood.jpg"
                objMaterial = New Bitmap(strFilePath)
            Case Else
                MsgBox("ERROR: getMaterialImage()", Title:="ERROR")
                Return objMaterial
        End Select
        Return objMaterial
    End Function

    'Returns String() of available Sizes based on Layout.
    '****NEEDS Desktop, Fanned AND Collage SIZING******
    Public Function getSizing(ByVal myLayout As String)
        Dim mySizing As String() = Nothing
        Select Case myLayout
            Case "Top 5 Ordered"
                ReDim mySizing(2)
                mySizing = {"12 x 60 (in)", "18 x 90 (in)", "24 x 120 (in)"}    'Same as Square
                Return mySizing
            Case "Top 5 Square"
                ReDim mySizing(3)
                mySizing = {"24 x 24 (in)", "36 x 36 (in)", "40 x 40 (in)"}
                Return mySizing
            Case "Top 5 Fanned"
                ReDim mySizing(5)
                mySizing = {"12 x 24 (in)", "18 x 36 (in)", "24 x 48 (in)", "30 x 60 (in)", "36 x 72 (in)", "40 x 80 (in)"}  '???
                Return mySizing
            Case "Desktop"
                ReDim mySizing(5)
                mySizing = {}  '???

            Case "Record Center Piece"
                ReDim mySizing(3)
                mySizing = {"24 x 24 (in)", "36 x 36 (in)", "40 x 40 (in)"}
                Return mySizing
            Case "Collage"
                ReDim mySizing(5)
                mySizing = {}  '???

            Case "Spliced"
                ReDim mySizing(3)
                mySizing = {"24 x 24 (in)", "36 x 36 (in)", "40 x 40 (in)"}
                Return mySizing
            Case Else
                MsgBox("ERROR: getSizing()", Title:="ERROR")
                Return mySizing
        End Select
        Return mySizing
    End Function

    'Next - Sets the next QA Section Label Green, previous Label Grey.
    Public Sub NextQASection(ByRef section As GroupBox)
        Dim lblSection As Label = getSelectedQASection(section)
        lblSection.Font = New Font(lblSection.Font, FontStyle.Bold)
        Select Case lblSection.Text
            Case "LAYOUT"
                section.Controls("lblArrows1").ForeColor = Color.ForestGreen
                section.Controls("lblMaterial").ForeColor = Color.ForestGreen
                section.Controls("lblMaterial").Font = New Font(section.Controls("lblMaterial").Font, FontStyle.Underline)
            Case "MATERIAL"
                section.Controls("lblArrows2").ForeColor = Color.ForestGreen
                section.Controls("lblSize").ForeColor = Color.ForestGreen
                section.Controls("lblSize").Font = New Font(section.Controls("lblMaterial").Font, FontStyle.Underline)
            Case "SIZE"
                section.Controls("lblArrows3").ForeColor = Color.ForestGreen
                section.Controls("lblImages").ForeColor = Color.ForestGreen
                section.Controls("lblImages").Font = New Font(section.Controls("lblMaterial").Font, FontStyle.Underline)
            Case "IMAGES"
                section.Controls("lblArrows4").ForeColor = Color.ForestGreen
                section.Controls("lblDetails").ForeColor = Color.ForestGreen
                section.Controls("lblDetails").Font = New Font(section.Controls("lblMaterial").Font, FontStyle.Underline)
            Case "DETAILS"
                '*** NEEDS CODE FOR FINISH ***
            Case Else
                MsgBox("ERROR: nextQASection()", Title:="ERROR")
        End Select
    End Sub

    'Back - Sets the previous QA Section Label Green, latter Label Green.
    Public Sub BackQASection(ByRef section As GroupBox)
        Dim lblSection As Label = getSelectedQASection(section)
        lblSection.Font = New Font(lblSection.Font, FontStyle.Bold)
        Select Case lblSection.Text
            Case "MATERIAL"
                section.Controls("lblArrows1").ForeColor = Color.DarkGray
                section.Controls("lblMaterial").ForeColor = Color.DarkGray
                section.Controls("lblMaterial").Font = New Font(section.Controls("lblMaterial").Font, FontStyle.Underline)
            Case "SIZE"
                section.Controls("lblArrows2").ForeColor = Color.DarkGray
                section.Controls("lblSize").ForeColor = Color.DarkGray
                section.Controls("lblSize").Font = New Font(section.Controls("lblMaterial").Font, FontStyle.Underline)
            Case "IMAGES"
                section.Controls("lblArrows3").ForeColor = Color.DarkGray
                section.Controls("lblImages").ForeColor = Color.DarkGray
                section.Controls("lblImages").Font = New Font(section.Controls("lblMaterial").Font, FontStyle.Underline)
            Case "DETAILS"
                section.Controls("lblArrows4").ForeColor = Color.DarkGray
                section.Controls("lblDetails").ForeColor = Color.DarkGray
                section.Controls("lblDetails").Font = New Font(section.Controls("lblMaterial").Font, FontStyle.Underline)
            Case Else
                MsgBox("ERROR: backQASection()", Title:="ERROR")
        End Select
    End Sub


    'Get Image File Path
    Public Function getImagePath(ByRef strImageName As String)
        Dim strImagePath As String = "E:\EJ\MyPrograms\PrintCreator\Images"
        Select Case strImageName
            Case "Top 5 Ordered"
                strImagePath += "\Stock Photos\Stock-Top5Ordered.png"
            Case "Top 5 Square"
                strImagePath += "\Stock Photos\Stock-Top5Square.png"
            Case "Top 5 Fanned"
                strImagePath += "\Stock Photos\Stock-Top3Fanned.jpg"
            Case "Desktop"
                strImagePath += "\Stock Photos\Stock-Desktop-3.jpg"
            Case "Record Center Piece"
                strImagePath += "\Stock Photos\Stock-RecordCenterPiece.png"
            Case "Collage"
                strImagePath += "\Stock Photos\Stock-Collage-5-5.png"
            Case "Spliced"
                strImagePath += "\Stock Photos\Stock-Splice-TRBL-4.png"
            Case Else
                strImagePath = Nothing
        End Select
        Return strImagePath
    End Function

    'Load selected Image into PictureBox
    Public Sub LoadImageToPictureBox(ByRef myQAGroupBox As GroupBox, ByRef myPictureBox As PictureBox)
        Dim selectedQA As CustomRadioButton = Nothing
        Dim strPath As String = Nothing
        selectedQA = getCheckedButton(myQAGroupBox)
        If selectedQA IsNot Nothing Then
            strPath = getImagePath(selectedQA.Text)
            If strPath <> Nothing Then
                Dim myImage As Image = Image.FromFile(strPath)
                myImage = fitImageToGroupBox(myImage, HomePage.groupImage)
                myPictureBox.Width = myImage.Width
                myPictureBox.Height = myImage.Height
                myPictureBox.Image = myImage
                myPictureBox.Invalidate()
            End If
        End If
    End Sub

    'Load selected Image into Panel
    Public Sub LoadImageToPanel(ByRef myQAGroupBox As GroupBox, ByRef myPanel As Panel)
        Dim selectedQA As CustomRadioButton = Nothing
        Dim strPath As String = Nothing
        selectedQA = getCheckedButton(myQAGroupBox)
        If selectedQA IsNot Nothing Then
            strPath = getImagePath(selectedQA.Text)
            If strPath <> Nothing Then
                Dim myImage As Image = Image.FromFile(strPath)
                myImage = fitImageToGroupBox(myImage, HomePage.groupImage)
                myPanel.Width = myImage.Width
                myPanel.Height = myImage.Height
                myPanel.BackgroundImage = myImage
                myPanel.Invalidate()
            End If
        End If
    End Sub

    'Center Control on GroupBox
    Public Sub CenterControl(ByRef myControl As Control, ByRef myPicGroupBox As GroupBox)
        Dim centerPoint As Point = Nothing
        centerPoint.X = (myPicGroupBox.Width / 2) - (myControl.Width / 2)
        centerPoint.Y = (myPicGroupBox.Height / 2) - (myControl.Height / 2)
        myControl.Location = centerPoint
        myPicGroupBox.Invalidate()
    End Sub

    'Center Image on another Image
    Public Function resizeAndCenterImageOnBitmap(ByRef foreImage As Bitmap, ByRef backgroundImage As Bitmap, ByVal foreImageSize As Size)
        Dim returnImage As Image = Nothing
        Dim centerPoint As Point = Nothing
        foreImage = resizeBitmap(foreImage, foreImageSize.Width, foreImageSize.Height)
        centerPoint.X = (backgroundImage.Width / 2) - (foreImage.Width / 2)
        centerPoint.Y = (backgroundImage.Height / 2) - (foreImage.Height / 2)
        returnImage = funDrawImageOnImage(foreImage, backgroundImage, centerPoint)
        Return returnImage
    End Function

    'Return largest side; Width or Height
    Public Function getLargestSide(ByRef myImage As Bitmap)
        If myImage.Width > myImage.Height Then
            Return myImage.Width
        Else
            Return myImage.Height
        End If
    End Function

    'Place Background behind Image
    Public Function addImageBackground(ByRef myImage As Image, ByRef myColor As Color, ByVal thickness As Single)
        Dim returnImage As New Bitmap(myImage.Width, myImage.Height)
        Dim myGraphic As Drawing.Graphics = Drawing.Graphics.FromImage(returnImage)
        Dim mySize As Size = Nothing
        mySize.Width = myImage.Width - (thickness * 5)
        mySize.Height = myImage.Height - (thickness * 5)
        myGraphic.Clear(myColor)
        returnImage = resizeAndCenterImageOnBitmap(myImage, returnImage, mySize)
        Return returnImage
    End Function



End Module

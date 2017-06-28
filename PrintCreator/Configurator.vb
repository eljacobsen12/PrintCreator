Public Class HomePage

    '*************************************************
    '******  USE THREAD POOLING TO LOAD IMAGES  ******
    '*************************************************

    Dim selectedSection As Label = Nothing

    Private Sub HomePage_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Create Graphics object.
        'Dim g As Graphics = PictureBox.CreateGraphics()
        Dim g As Graphics = myPanel.CreateGraphics()

        selectedSection = getSelectedQASection(groupQAsections)
        lblSelectedSectionLabel.Text = selectedSection.Text
        createQA(groupQA, selectedSection.Text)
        'LoadImageToPictureBox(groupQA, PictureBox)
        LoadImageToPanel(groupQA, myPanel)
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        Dim selectedButton As RadioButton = getCheckedButton(groupQA)
        NextQASection(groupQAsections)
        selectedSection = getSelectedQASection(groupQAsections)
        DeleteRadioButtons(groupQA)
        createQA(groupQA, selectedSection.Text)
    End Sub

    Private Sub ImageSplicing1ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImageSplicing1Menu.Click
        ImageSplicingExample1.Show()
    End Sub

    Private Sub ImageSplicing2ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImageSplicing3Menu.Click
        ImageSplicingExample3.Show()
    End Sub

    Private Sub TestFormMenu_Click(sender As Object, e As EventArgs) Handles TestFormMenu.Click
        TestForm.Show()
    End Sub

    Private Sub ImageSplicing2Menu_Click(sender As Object, e As EventArgs) Handles ImageSplicing2Menu.Click
        ImageSplicingExample2.Show()
    End Sub

    'Loads Image into PictureBox whenever a RadioButton is clicked in groupQA GroupBox
    'Public Sub HandleRadioButtonChange(ByVal sender As Object, ByVal e As EventArgs)
    '    If PictureBox IsNot Nothing Then
    '        LoadImageToPictureBox(groupQA, PictureBox)
    '        CenterControl(PictureBox, groupImage)
    '    Else
    '        'Create Panels???
    '    End If
    'End Sub

    'Loads Image into Panel whenever a RadioButton is clicked in groupQA GroupBox
    Public Sub HandleRadioButtonChange(ByVal sender As Object, ByVal e As EventArgs)
        If myPanel IsNot Nothing Then
            LoadImageToPanel(groupQA, myPanel)
            CenterControl(myPanel, groupImage)
        Else
            'Create Panels???
        End If
    End Sub

End Class
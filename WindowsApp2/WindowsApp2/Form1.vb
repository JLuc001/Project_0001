Public Class Form1
    ' Image ratio
    Public dRatioImageWH As Double
    ' Updated Image
    Public resizedImage As Image
    Private ReadOnly Glass As New cGlassWindow

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        dRatioImageWH = 1.501
        Dim img As Image = My.Resources.Alps01
        PictureBox1.BackgroundImage = img
        PictureBox1.Size = New Size(400, CInt((img.Height / img.Width) * 400))
        PictureBox1.BackgroundImageLayout = ImageLayout.Stretch
        PictureBox1.Controls.Add(Glass)

        ' Resize and Save the image for use in GlassWindow class
        resizedImage = fResizeImage(img, PictureBox1.Width, PictureBox1.Height)
        Glass.Size = New Size(CInt(100 * dRatioImageWH), 100)
        Glass.Location = New Point(0, 0)
        Glass.ForeColor = Color.Red
        PictureBox1.Controls.Add(Glass)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        dRatioImageWH = 0.666
        Dim img As Image = My.Resources.Alps02
        PictureBox1.BackgroundImage = img
        PictureBox1.Size = New Size(CInt((img.Width / img.Height) * 400), 400)
        PictureBox1.BackgroundImageLayout = ImageLayout.Stretch
        PictureBox1.Controls.Add(Glass)

        ' Resize and Save the image for use in GlassWindow class
        resizedImage = fResizeImage(img, PictureBox1.Width, PictureBox1.Height)
        Glass.Size = New Size(100, CInt(100 / dRatioImageWH))
        Glass.Location = New Point(0, 0)
        Glass.ForeColor = Color.Red
        PictureBox1.Controls.Add(Glass)

    End Sub

    Public Function fResizeImage(ByVal imgX As Image, ByVal xWidth As Integer, ByVal yHeight As Integer) As Image
        ' ------------------------------------------------------------
        ' Resize the image for the GlassWindow class
        ' ------------------------------------------------------------
        Dim bm As Bitmap = imgX
        Dim width As Integer = CInt(Val(xWidth))
        Dim height As Integer = CInt(Val(yHeight))
        Dim thumb As New Bitmap(width, height)
        Using g As Graphics = Graphics.FromImage(thumb)
            g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
            g.DrawImage(bm, New Rectangle(0, 0, width, height), New Rectangle(0, 0, bm.Width, bm.Height), GraphicsUnit.Pixel)
        End Using
        Return thumb
        ' ------------------------------------------------------------
    End Function

End Class

Public Class cGlassWindow

    ' ------------------------------------------------------------
    Inherits PictureBox ' Control Glass
    ' ------------------------------------------------------------
    ' DefWndProc constant used
    Private Const WM_NCLBUTTONDOWN As Integer = 161 '&HA1
    Private Const HTCAPTION As Integer = 2
    'Private Const HTLEFT As Integer = 10
    'Private Const HTRIGHT As Integer = 11
    'Private Const HTTOP As Integer = 12
    'Private Const HTTOPLEFT As Integer = 13
    'Private Const HTTOPRIGHT As Integer = 14
    'Private Const HTBOTTOM As Integer = 15
    'Private Const HTBOTTOMLEFT As Integer = 16
    Private Const HTBOTTOMRIGHT As Integer = 17
    ' Edge initialisation
    Private mEdge As EdgeEnum = EdgeEnum.None
    ' Edge on the Control
    Private Enum EdgeEnum
        TopLeft
        Top
        TopRight
        Right
        BottomRight
        Bottom
        BottomLeft
        Left
        None
    End Enum
    ' ------------------------------------------------------------

    Public Sub New()
        ' ------------------------------------------------------------
        Me.Cursor = Cursors.SizeAll
        ' ------------------------------------------------------------
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        ' ------------------------------------------------------------
        ' Paint the Glass window (picturebox)
        ' ------------------------------------------------------------
        Try
            ' Get the glass background from the resizedImage 
            Dim r As Rectangle = Me.ClientRectangle
            r.Location = Me.Location
            Using bmp As New Bitmap(r.Width, r.Height), g As Graphics = Graphics.FromImage(bmp)
                g.DrawImage(Form1.resizedImage, Me.ClientRectangle, r, GraphicsUnit.Pixel)
                Me.BackgroundImage = CType(bmp.Clone, Drawing.Image)
            End Using
            ' Draw the border
            e.Graphics.DrawRectangle(New Pen(Me.ForeColor, 4), Me.ClientRectangle)
            e.Graphics.FillRectangle(Brushes.Red, Me.ClientRectangle.Size.Width - 10, Me.ClientRectangle.Size.Height - 10, 10, 10)
            MyBase.OnPaint(e)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        ' ------------------------------------------------------------
    End Sub

    Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
        ' ------------------------------------------------------------
        ' To move and resize the Glass window
        ' ------------------------------------------------------------
        Try
            ' Send the message to the default Control
            Select Case mEdge
                Case EdgeEnum.None
                    Me.Capture = False
                    Me.DefWndProc(Message.Create(Me.Handle, WM_NCLBUTTONDOWN, New IntPtr(HTCAPTION), IntPtr.Zero))
                Case EdgeEnum.BottomRight
                    Me.Capture = False
                    Me.DefWndProc(Message.Create(Me.Handle, WM_NCLBUTTONDOWN, New IntPtr(HTBOTTOMRIGHT), IntPtr.Zero))
                Case Else
                    ' Nothing
            End Select
            ' Raise the MouseDown event
            MyBase.OnMouseDown(e)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        ' ------------------------------------------------------------
    End Sub

    Protected Overrides Sub OnMouseMove(ByVal e As System.Windows.Forms.MouseEventArgs)
        ' ------------------------------------------------------------
        ' Choice of Cursor and Edge
        ' ------------------------------------------------------------
        Try
            If e.Y > Me.ClientSize.Height - 10 And e.X > Me.ClientSize.Width - 10 Then
                ' BOTTOM-RIGHT edge
                Me.Cursor = Cursors.SizeNWSE
                mEdge = EdgeEnum.BottomRight
            Else
                ' NO edge
                Me.Cursor = Cursors.SizeAll
                mEdge = EdgeEnum.None
            End If
            ' Raise the MouseMove event
            MyBase.OnMouseMove(e)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        ' ------------------------------------------------------------
    End Sub

    Protected Overrides Sub OnMove(ByVal e As System.EventArgs)
        ' ------------------------------------------------------------
        ' Control the mouvement of Glass window
        ' ------------------------------------------------------------
        Try
            ' Keep the Glass window inside the main PictureBox
            If Me.Location.X <= 0 Then Me.Location = New Point(0, Me.Location.Y)
            If Me.Location.Y <= 0 Then Me.Location = New Point(Me.Location.X, 0)
            If Me.Location.X + Me.Width >= Form1.PictureBox1.Width Then Me.Location = New Point(Form1.PictureBox1.Width - Me.Width, Me.Location.Y)
            If Me.Location.Y + Me.Height >= Form1.PictureBox1.Height Then Me.Location = New Point(Me.Location.X, Form1.PictureBox1.Height - Me.Height)
            ' Control to be redrawn
            Me.Invalidate()
            ' Raise the Move event
            MyBase.OnMove(e)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        ' ------------------------------------------------------------
    End Sub

    Protected Overrides Sub OnResize(ByVal e As System.EventArgs)
        ' ------------------------------------------------------------
        ' To redraw the whole image based on Glass window
        ' ------------------------------------------------------------
        Try
            ' Minimum limits of Glass window
            If Me.Width < 40 Then Me.Width = CInt(40 * Form1.dRatioImageWH)
            If Me.Height < 40 Then Me.Height = CInt(40 / Form1.dRatioImageWH)
            ' Effect on Resize event

            If Me.Width > Form1.PictureBox1.Width - Me.Location.X Then Me.Width = Form1.PictureBox1.Width - Me.Location.X

            If Me.Height > Form1.PictureBox1.Height - Me.Location.Y Then Me.Height = Form1.PictureBox1.Height - Me.Location.Y

            If Form1.dRatioImageWH > 1 Then Me.Height = CInt(Me.Width / Form1.dRatioImageWH)
            If Form1.dRatioImageWH < 1 Then Me.Width = CInt(Me.Height * Form1.dRatioImageWH)
            'Me.Width = CInt(Me.Height * frmECalendar.dRatioImageWH)


            Form1.Label1.Text = Me.Width.ToString
            Form1.Label2.Text = Me.Height.ToString
            Form1.Label3.Text = Form1.dRatioImageWH.ToString
            Form1.Label4.Text = CBool(Me.Width > Form1.PictureBox1.Width - Me.Location.X).ToString
            Form1.Label5.Text = CBool(Me.Height > Form1.PictureBox1.Height - Me.Location.Y).ToString
            Form1.Label6.Text = (Form1.PictureBox1.Height - Me.Location.Y).ToString

            ' Control to be redrawn
            Me.Invalidate()
            ' Raise the Resize event
            MyBase.OnResize(e)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        ' ------------------------------------------------------------
    End Sub

End Class

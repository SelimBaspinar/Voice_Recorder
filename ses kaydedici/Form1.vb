Imports System.Runtime.InteropServices
Imports System.Text

    Public Class Form1
        Public KullAdi As String = SystemInformation.UserName
    <DllImport("winmm.dll")> _
    Private Shared Function mciSendString(ByVal command As String, ByVal buffer As StringBuilder, ByVal bufferSize As Integer, ByVal hwndCallback As IntPtr) As Integer
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs)
        ' Çalıştır


    End Sub
    Dim IsDraggingForm As Boolean = False
    Private MousePos As New System.Drawing.Point(0, 0)

    Private Sub Form1_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles MyBase.MouseDown
        If e.Button = MouseButtons.Left Then
            IsDraggingForm = True
            MousePos = e.Location
        End If
    End Sub

    Private Sub Form1_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs) Handles MyBase.MouseUp
        If e.Button = MouseButtons.Left Then IsDraggingForm = False
    End Sub

    Private Sub Form1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles MyBase.MouseMove
        If IsDraggingForm Then
            Dim temp As Point = New Point(Me.Location + (e.Location - MousePos))
            Me.Location = temp
            temp = Nothing
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Hide()
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Label4.Text = Label4.Text + 1
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        If Label4.Text = "60" Then
            Label3.Text = Label3.Text + 1
            Label4.Text = "0"
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If Label3.Text = "60" Then
            Label2.Text = Label2.Text + 1
            Label4.Text = "0"
        End If
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Try
            Timer1.Start()
            Timer2.Start()
            Timer3.Start()
            Dim i As Integer
            i = mciSendString("open new type waveaudio alias capture", Nothing, 0, 0)
            i = mciSendString("record capture", Nothing, 0, 0)
            Label1.Text = "Kayıt Başladı..."
            Label1.BackColor = Color.Green
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        Try
            Timer1.Stop()
            Timer2.Stop()
            Timer3.Stop()
            Label4.Text = "0"
            Label3.Text = "0"
            Label2.Text = "0"
            Dim i As Integer
            Dim isim As String
            Try
                isim = InputBox("Lütfen İsim Yazın")
                FolderBrowserDialog1.ShowDialog()
                TextBox1.Text = FolderBrowserDialog1.SelectedPath + ("\" & isim & ".wav")
            Catch ex As Exception
                MsgBox("Lütfen Bilgileri Doğru Doldurun")
            End Try
            
            i = mciSendString("save capture " & TextBox1.Text, Nothing, 0, 0)
            i = mciSendString("close capture", Nothing, 0, 0)
            Label1.Text = "Kaydedildi"
            Label1.BackColor = Color.Yellow
        Catch ex As Exception
            MsgBox("Lütfen Bilgileri Doğru Doldurun")
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        Try
            My.Computer.Audio.Play(TextBox1.Text, AudioPlayMode.Background)
        Catch ex As Exception
            MsgBox("Öncelikle Ses Kaydını Kaydedin.", MsgBoxStyle.Critical, "             Hata                ")
        End Try

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Me.Close()
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Me.WindowState = FormWindowState.Minimized

    End Sub
End Class




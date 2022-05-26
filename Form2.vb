Imports System.Net.Mail
Public Class Form2
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim e_mail As New MailMessage()
        Dim i As Integer
        Try
            e_mail.From = New MailAddress(TextBox5.Text + ComboBox1.SelectedText)
            e_mail.To.Add(TextBox6.Text)
            e_mail.Subject = TextBox7.Text
            e_mail.Body = RichTextBox2.Text

            For i = 0 To CheckedListBox1.Items.Count - 1
                Try
                    e_mail.Attachments.Add(New Attachment(CheckedListBox1.CheckedItems(i).ToString))
                Catch ex As Exception

                End Try
            Next

            If ComboBox1.SelectedItem = "@gmail.com" Then
                Dim smtp As New SmtpClient("smtp.gmail.com")
                smtp.EnableSsl = False
                smtp.Port = 587
                smtp.Credentials = New Net.NetworkCredential(TextBox6.Text, TextBox8.Text)
                smtp.Send(e_mail)
            Else
                Dim smtp As New SmtpClient("smtp.live.com")
                smtp.Port = 587
                smtp.EnableSsl = True
                smtp.Credentials = New Net.NetworkCredential(TextBox6.Text, TextBox8.Text)
            End If

        Catch ex As Exception
            MsgBox("gönderildi")
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        OpenFileDialog1.ShowDialog()
        CheckedListBox1.Items.Add(OpenFileDialog1.FileName)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form1.Show()
        Me.Hide()
    End Sub

    Private Sub Form2_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Form1.Show()
        Me.Hide()
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("@hotmail.com")
        ComboBox1.Items.Add("@gmail.com")
        TextBox8.PasswordChar = "*"
    End Sub
End Class
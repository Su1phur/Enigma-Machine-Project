Imports System.Net.Mail
Public Class Form2
	Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
		Dim SenderUsername As String = TextBox6.Text
		Dim SenderEmailaddress As String = TextBox1.Text
		Dim Password As String = TextBox2.Text
		Dim ReceiverEmailaddress As String = TextBox3.Text
		Dim Subject As String = TextBox5.Text
		Dim ContentText As String = TextBox4.Text  ' content is always encrypted
		Dim Attachment As New Attachment("Import.txt")

		Dim NewEmail As New MailMessage(SenderEmailaddress, ReceiverEmailaddress, Subject, ContentText)

		NewEmail.Attachments.Add(Attachment)

		Dim Smtpclient As New SmtpClient()

		Smtpclient.UseDefaultCredentials = False
		Smtpclient.Credentials = New Net.NetworkCredential(SenderEmailaddress, Password)
		Smtpclient.Port = 587
		Smtpclient.EnableSsl = True
		Smtpclient.Host = "webmail.strathallan.co.uk"

		Try

			Smtpclient.Send(NewEmail)
			MsgBox("Message sent")
		Catch ex As Exception
			MsgBox(ex.Message)
		End Try

	End Sub

End Class
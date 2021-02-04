Imports System.Net.Mail
Imports System.IO.File

Public Class EnvioCorreo

    Public Function SendEmail(ByVal varCorreoDestinatario As String,
                              ByVal varAsunto As String,
                              ByVal varDescripcionCorreo As String,
                              ByRef varMensajeError As String,
                              Optional ByVal varEsHTML As Boolean = False,
                              Optional ByVal varCopiaCorreoDestinatario As String = "",
                              Optional ByVal varArchivosAdjuntos() As String = Nothing) As Boolean

        varMensajeError = String.Empty

        Try

            Dim Smtp_Server As New SmtpClient
            Dim e_mail As New MailMessage()

            Smtp_Server.UseDefaultCredentials = False
            Smtp_Server.Credentials = New Net.NetworkCredential(MailAccount, MailPassword)
            Smtp_Server.Port = Port
            Smtp_Server.EnableSsl = EnableSsl
            Smtp_Server.Host = Host

            Using email As MailMessage = New MailMessage()

                e_mail.From = New MailAddress(MailAccount)
                e_mail.To.Add(varCorreoDestinatario)

                If Not String.IsNullOrEmpty(varCopiaCorreoDestinatario) Then
                    e_mail.CC.Add(varCopiaCorreoDestinatario)
                End If

                e_mail.Subject = varAsunto
                e_mail.IsBodyHtml = varEsHTML
                e_mail.Body = varDescripcionCorreo

                If Not IsNothing(varArchivosAdjuntos) Then
                    For Each adjunto As String In varArchivosAdjuntos
                        If Exists(adjunto) Then
                            Dim attachment As Attachment = New Attachment(adjunto)
                            e_mail.Attachments.Add(attachment)
                        End If
                    Next
                End If

                Smtp_Server.Send(e_mail)

            End Using

            Return True

        Catch ex As Exception
            varMensajeError = ex.Message
            Return False
        End Try
    End Function

End Class

Public Class Acciones

    Public Sub EnvioNotificaciones()

        Dim datos As DatosNotificaciones = New DatosNotificaciones()


        'Salvar las notificaiones en la tabla para su envío.
        Dim nombreSP As String = "sp_MPL_SalvaNotificaciones"
        If Not datos.ExisteSP(nombreSP) Then 'Si no existe el procedimiento almacenado se sale del método.
            Exit Sub
        End If

        datos.SalvarRegistrosNotificaciones(nombreSP)



        'Enviar notificaciones mientras se devuelvan registros.
        nombreSP = "sp_MPL_NotificacionesParaEnvio"
        If Not datos.ExisteSP(nombreSP) Then 'Si no existe el procedimiento almacenado se sale del método.
            Exit Sub
        End If

        Dim fecha As Date = Date.Now

        'Se crea tabla para controlar notificaciones que se vayan repitiendo de forma continua.
        Dim datosRepetidos As DataTable = New DataTable()
        datosRepetidos.Columns.Add("CodigoNotificacion", GetType(Integer))
        datosRepetidos.Columns.Add("CantidadIteraciones", GetType(Integer))

        Do
            Try

                'Se obtienen las notificaiones de la tabla para su envío.
                nombreSP = "sp_MPL_NotificacionesParaEnvio"
                Dim notificaciones As DataTable = datos.ObtenerNotificaciones(nombreSP)

                If notificaciones.Rows.Count = 0 Then
                    Exit Do
                End If

                'Se recorre las notificaciones para enviar una por una.
                For Each notificacion As DataRow In notificaciones.Rows

                    Try

                        'Si se recorre la misma notificacion mas de 10 veces se sale del ciclo.
                        For Each datoRepetido As DataRow In datosRepetidos.Rows
                            If datoRepetido("CantidadIteraciones") > 10 Then
                                Exit Do
                            End If

                            If datoRepetido("CodigoNotificacion") = notificacion("CodigoNotificacion") Then
                                datoRepetido("CantidadIteraciones") += 1
                            End If
                        Next

                        'Si no existe la notificacion en la tabla de repetidas se inserta.
                        Dim existe As Boolean = (From datoRepetido In datosRepetidos.AsEnumerable()
                                                 Where datoRepetido.Field(Of Integer)("CodigoNotificacion") = notificacion("CodigoNotificacion")).Any()

                        If Not existe Then
                            datosRepetidos.Rows.Add(notificacion("CodigoNotificacion"), 1)
                        End If


                        EnviaCorreo(notificacion)

                        'Cambiar estado a Enviado.
                        nombreSP = "sp_MPL_CambiaEstadoNotificaciones"
                        If Not datos.ExisteSP(nombreSP) Then 'Si no existe el procedimiento almacenado se sale del método.
                            Exit Sub
                        End If

                        datos.CambiarEstadoNotificaciones(nombreSP, notificacion("CodigoNotificacion"))


                    Catch ex As Exception

                        GuardaLogBitacora(ex.Message, fecha)

                    End Try

                Next

            Catch ex As Exception
                GuardaLogBitacora(ex.Message, fecha)
            End Try

        Loop



        'Limpiar la tabla de notificaiones.
        nombreSP = "sp_MPL_LimpiarNotificaciones"
        If Not datos.ExisteSP(nombreSP) Then 'Si no existe el procedimiento almacenado se sale del método.
            Exit Sub
        End If

        datos.LimpiarNotificaciones(nombreSP)

    End Sub

    Private Sub EnviaCorreo(ByVal notificacion As DataRow)

        Dim adjuntos As String() = Nothing
        If Not String.IsNullOrEmpty(notificacion("CorreoAdjuntos")) Then
            adjuntos = notificacion("CorreoAdjuntos").ToString.Trim.Split(New Char() {";"c}, StringSplitOptions.RemoveEmptyEntries)
        End If

        Dim mensajError As String = String.Empty
        Dim envioCorreo As EnvioCorreo = New EnvioCorreo()
        If Not envioCorreo.SendEmail(notificacion("CorreoDestinatario"),
                                     notificacion("CorreoAsunto"),
                                     notificacion("CorreoDetalle"),
                                     mensajError,
                                     notificacion("CorreoEsHTML"),
                                     notificacion("CorreoConCopia").Trim(),
                                     adjuntos) Then

            Throw New Exception(mensajError)

        End If

    End Sub

End Class

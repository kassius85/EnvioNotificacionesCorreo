Imports System.Configuration
Imports System.Threading

Public Class Notificaciones
    'Get the Scheduled Time from AppSettings.
    Private ScheduledTime As Date = Date.Parse(ScheduledTimeExec)

    Private Schedular As Timer

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Agregue el código aquí para iniciar el servicio. Este método debería poner
        ' en movimiento los elementos para que el servicio pueda funcionar.
        ScheduleService()
    End Sub

    Protected Overrides Sub OnStop()
        ' Agregue el código aquí para realizar cualquier anulación necesaria para detener el servicio.
        Schedular.Dispose()
    End Sub

    Public Sub ScheduleService()
        Try
            'Initialize the Schedular
            Schedular = New Timer(New TimerCallback(AddressOf SchedularCallback))

            Dim schedule As String = String.Empty
            Dim executeTask As Boolean = True

            'Get the current minute
            Dim tempDate As Date = Date.Now
            Dim currentDate As Date = New Date(tempDate.Year, tempDate.Month, tempDate.Day, tempDate.Hour, tempDate.Minute, 0, 0)

            Select Case Mode
                Case "DAILY"
                    'If Scheduled Time is passed set Schedule for the next day.
                    If currentDate <> ScheduledTime Then
                        If currentDate > ScheduledTime Then ScheduledTime = ScheduledTime.AddDays(1)

                        executeTask = False
                    Else
                        ScheduledTime = ScheduledTime.AddDays(1)
                    End If

                Case "INTERVAL"
                    'Get the Interval in Minutes from AppSettings.
                    If IntervalMinutes > 0 Then
                        If currentDate <> ScheduledTime Then
                            If currentDate > ScheduledTime Then ScheduledTime = ScheduledTime.AddDays(1)

                            executeTask = False
                        Else
                            'Set the Scheduled Time by adding the Interval to Current Time.
                            ScheduledTime = currentDate.AddMinutes(IntervalMinutes)

                            'If Scheduled Time is passed set Schedule for the next Interval.
                            If Date.Now > ScheduledTime Then ScheduledTime = ScheduledTime.AddMinutes(IntervalMinutes)
                        End If
                    Else
                        Throw New Exception("El intervalo en minutos debe ser mayor a cero.")
                    End If

                Case Else
                    Throw New Exception("El modo definido no es válido.")

            End Select

            schedule = ScheduledTime.ToString("dd/MM/yyyy hh:mm:ss tt")

            'Update Schedular
            ChangeSchedular()

            'Execute task
            Dim acciones As Acciones = New Acciones()
            If executeTask Then acciones.EnvioNotificaciones()

            'Save Log
            If MuestraLog Then GuardaLogBitacora(IIf(executeTask, "Proceso ejecutado! ", String.Empty) + "Próxima ejecución aproximada: " + schedule, Date.Now)

        Catch ex As Exception
            ChangeSchedular(Date.Now.AddHours(1))
            GuardaLogBitacora(ex.Message, Date.Now)
        End Try
    End Sub

    Private Sub SchedularCallback(e As Object)
        ScheduleService()
    End Sub

    Private Sub ChangeSchedular(Optional ByVal scheduledTimeTemp As Date = Nothing)
        scheduledTimeTemp = IIf(scheduledTimeTemp = Date.MinValue, ScheduledTime, scheduledTimeTemp)

        'Get the difference in Minutes between the Scheduled and Current Time.
        Dim timeSpan As TimeSpan = scheduledTimeTemp.Subtract(Date.Now)
        Dim dueTime As Integer = Convert.ToInt32(timeSpan.TotalMilliseconds)

        'Change the Timer's Due Time.
        Schedular.Change(dueTime, Timeout.Infinite)
    End Sub
    'FIN DEL TIMER

End Class

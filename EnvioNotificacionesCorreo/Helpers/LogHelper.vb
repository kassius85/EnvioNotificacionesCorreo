Imports System.Configuration
Imports System.IO

Public Module LogHelper

    Public Sub GuardaLogBitacora(ByVal text As String, Optional ByVal fechaHora As Date = Nothing)

        Dim fecha As String = String.Empty
        If Not IsNothing(fechaHora) Then fecha = "_" + fechaHora.ToString("yyyyMMdd_HH-mm-ss")

        Dim path As String = PathLog
        If Not Directory.Exists(path) Then path = "C:\"

        Dim rutalog As String = IO.Path.Combine(path, LogName + fecha)

        File.AppendAllText(rutalog, Date.Now().ToString("dd/MM/yyyy hh:mm:ss tt") + " : " + text + vbCrLf + vbCrLf)
    End Sub

End Module

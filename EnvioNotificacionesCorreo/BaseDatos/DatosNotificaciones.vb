Imports System.Data
Imports System.IO
Imports System.Data.OleDb

Public Class DatosNotificaciones

    Private Property Connection As AccesoDatos.IAcceso

    Public Sub New()
        Connection = CrearConexion()
    End Sub

    Public Function ExisteSP(ByVal nombreSP As String) As Boolean

        Dim ds As DataSet

        Try

            If Environment.GetEnvironmentVariable("SYBASE_OCS").ToString() = "OCS-12_5" Then
                Dim query As String = String.Format("EXEC sp_MPL_ExisteSP '{0}'", nombreSP)
                ds = Connection.ExecSqlResults(query)
            Else
                Dim losParametros() As OleDb.OleDbParameter =
                {
                    New OleDb.OleDbParameter() With {.DbType = DbType.String, .ParameterName = "@NombreSP", .Value = nombreSP}
                }

                ds = Connection.ExecStoredResults("sp_MPL_ExisteSP", losParametros)
            End If

            If Not IsNothing(ds) Then
                If ds.Tables.Count > 0 Then
                    If Not IsNothing(ds.Tables(0)) Then
                        If ds.Tables(0).Rows.Count > 0 Then
                            If TypeOf ds.Tables(0).Rows(0)(0) Is String AndAlso ds.Tables(0).Rows(0)(0) = "ERROR" Then
                                Throw New Exception(ds.Tables(0).Rows(0)(1))
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try

        Dim result1 As String = ds.Tables(0).Rows(0).Item("Existe")
        Dim result2 As Boolean = False

        If Not String.IsNullOrEmpty(result1) Then
            result2 = Boolean.Parse(result1)
        End If

        Return result2
    End Function

    Public Sub SalvarRegistrosNotificaciones(ByVal nombreSP As String)
        Dim ds As DataSet

        Try
            Dim cantReg As Integer = 1000

            If Environment.GetEnvironmentVariable("SYBASE_OCS").ToString() = "OCS-12_5" Then
                Dim query As String = String.Format("EXEC {0} {1}", nombreSP, cantReg)
                ds = Connection.ExecSqlResults(query)
            Else
                Dim losParametros() As OleDb.OleDbParameter =
                {
                    New OleDb.OleDbParameter() With {.DbType = DbType.Int32, .ParameterName = "@CantReg", .Value = cantReg}
                }

                ds = Connection.ExecStoredResults(nombreSP, losParametros)
            End If

            If Not IsNothing(ds) Then
                If ds.Tables.Count > 0 Then
                    If Not IsNothing(ds.Tables(0)) Then
                        If ds.Tables(0).Rows.Count > 0 Then
                            If TypeOf ds.Tables(0).Rows(0)(0) Is String AndAlso ds.Tables(0).Rows(0)(0) = "ERROR" Then
                                Throw New Exception(ds.Tables(0).Rows(0)(1))
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Function ObtenerNotificaciones(ByVal nombreSP As String) As DataTable
        Dim ds As DataSet

        Try
            Dim cantReg As Integer = 1000

            If Environment.GetEnvironmentVariable("SYBASE_OCS").ToString() = "OCS-12_5" Then
                Dim query As String = String.Format("EXEC {0} {1}", nombreSP, cantReg)
                ds = Connection.ExecSqlResults(query)
            Else
                Dim losParametros() As OleDb.OleDbParameter =
                {
                    New OleDb.OleDbParameter() With {.DbType = DbType.Int32, .ParameterName = "@CantReg", .Value = cantReg}
                }

                ds = Connection.ExecStoredResults(nombreSP, losParametros)
            End If

            If Not IsNothing(ds) Then
                If ds.Tables.Count > 0 Then
                    If Not IsNothing(ds.Tables(0)) Then
                        If ds.Tables(0).Rows.Count > 0 Then
                            If TypeOf ds.Tables(0).Rows(0)(0) Is String AndAlso ds.Tables(0).Rows(0)(0) = "ERROR" Then
                                Throw New Exception(ds.Tables(0).Rows(0)(1))
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try

        Return ds.Tables(0)

    End Function

    Public Sub CambiarEstadoNotificaciones(ByVal nombreSP As String,
                                           ByVal codigoNotificacion As Integer)
        Dim ds As DataSet

        Try

            If Environment.GetEnvironmentVariable("SYBASE_OCS").ToString() = "OCS-12_5" Then
                Dim query As String = String.Format("EXEC {0} {1}", nombreSP, codigoNotificacion)
                ds = Connection.ExecSqlResults(query)
            Else
                Dim losParametros() As OleDb.OleDbParameter =
                {
                    New OleDb.OleDbParameter() With {.DbType = DbType.Int32, .ParameterName = "@CodigoNotificacion", .Value = codigoNotificacion}
                }

                ds = Connection.ExecStoredResults(nombreSP, losParametros)
            End If

            If Not IsNothing(ds) Then
                If ds.Tables.Count > 0 Then
                    If Not IsNothing(ds.Tables(0)) Then
                        If ds.Tables(0).Rows.Count > 0 Then
                            If TypeOf ds.Tables(0).Rows(0)(0) Is String AndAlso ds.Tables(0).Rows(0)(0) = "ERROR" Then
                                Throw New Exception(ds.Tables(0).Rows(0)(1))
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Sub LimpiarNotificaciones(ByVal nombreSP As String)
        Dim ds As DataSet

        Try
            Dim cantReg As Integer = 1000

            If Environment.GetEnvironmentVariable("SYBASE_OCS").ToString() = "OCS-12_5" Then
                Dim query As String = String.Format("EXEC {0} {1}", nombreSP, cantReg)
                ds = Connection.ExecSqlResults(query)
            Else
                Dim losParametros() As OleDb.OleDbParameter =
                {
                    New OleDb.OleDbParameter() With {.DbType = DbType.Int32, .ParameterName = "@CantReg", .Value = cantReg}
                }

                ds = Connection.ExecStoredResults(nombreSP, losParametros)
            End If

            If Not IsNothing(ds) Then
                If ds.Tables.Count > 0 Then
                    If Not IsNothing(ds.Tables(0)) Then
                        If ds.Tables(0).Rows.Count > 0 Then
                            If TypeOf ds.Tables(0).Rows(0)(0) Is String AndAlso ds.Tables(0).Rows(0)(0) = "ERROR" Then
                                Throw New Exception(ds.Tables(0).Rows(0)(1))
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

End Class

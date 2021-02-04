Public Module ConexionBD

    Friend Const MICROSOFT_SQL_SERVER As Integer = 1
    Friend Const SYBASE_ADAPTIVE_SERVER As Integer = 2
    Friend MotorBD As String
    Friend So_Info As OperatingSystem

    Public Function CrearConexion() As AccesoDatos.IAcceso

        Dim Conn As AccesoDatos.IAcceso

        GrabarConexion(OLEDBConexion, BaseDatosConexion, UsuarioConexion, PasswordConexion, PasswordConexion, "2")

        Dim objEncrip As New Encripcion.Encripcion()
        MotorBD = objEncrip.DecryptStr(GetSetting("Sistemas", "Conexion", "Motor", CStr(MICROSOFT_SQL_SERVER)))

        Select Case MotorBD

            Case CStr(SYBASE_ADAPTIVE_SERVER)
                Conn = New AccesoDatos.AccesoSybase()

            Case CStr(MICROSOFT_SQL_SERVER)
                Conn = New AccesoDatos.AccesoSQLServer()

            Case Else
                Conn = New AccesoDatos.AccesoSQLServer()

        End Select

        Return Conn

    End Function

    Public Sub GrabarConexion(ByVal Servidor As String,
                              ByVal BD As String,
                              ByVal Usr As String,
                              ByVal Pwd As String,
                              ByVal PwdUsr As String,
                              Optional ByVal MotorBD As String = "")

        Dim strConn As String
        Dim MyEncrip As New Encripcion.Encripcion()

        SaveSetting("Sistemas", "Conexion", "Motor", MyEncrip.EncryptStr(MotorBD))

        'Sybase
        If MotorBD = CStr(SYBASE_ADAPTIVE_SERVER) Then
            strConn = "Provider=ASEOLEDB;Data Source=" & Servidor & ";Initial Catalog=" & BD
        Else
            strConn = "Server=" & Servidor & ";Database=" & BD
        End If

        'Agrega USUARIO y PASSWORD
        strConn = strConn & ";User ID=" & Usr & ";Password=" & Pwd

        'Grabar caracteristicas de la conexion en el Registry, para el acceso por medio de los componentes
        SaveSetting("Sistemas", "Conexion", "Valor", MyEncrip.EncryptStr(strConn))

        'Esto permitirá entre otras cosas que las capas de logica de negocio
        'identifiquen al usuario conectado a la base de datos
        SaveSetting("Sistemas", "Conexion", "Usuario", MyEncrip.EncryptStr(Usr))
        SaveSetting("Sistemas", "Conexion", "Password", MyEncrip.EncryptStr(Pwd))
        SaveSetting("Sistemas", "Conexion", "Servidor", MyEncrip.EncryptStr(Servidor))
        SaveSetting("Sistemas", "Conexion", "BaseDatos", MyEncrip.EncryptStr(BD))
        SaveSetting("Sistemas", "Conexion", "PasswordUsr", MyEncrip.EncryptStr(PwdUsr))

    End Sub

End Module

Imports System.Configuration
Imports System.Reflection

Module Configuracion
    Private _oLEDBConexion As String
    Private _baseDatosConexion As String
    Private _usuarioConexion As String
    Private _passwordConexion As String

    Private _host As String
    Private _port As Integer = 0
    Private _enableSsl As Boolean = Nothing
    Private _mailAccount As String
    Private _mailpassword As String

    Private _mode As String
    Private _scheduledTime As String
    Private _intervalMinutes As Integer = 0
    Private _muestraLog As Boolean = Nothing
    Private _pathLog As String
    Private _logName As String


    Public ReadOnly Property OLEDBConexion() As String
        Get
            If String.IsNullOrEmpty(_oLEDBConexion) Then
                _oLEDBConexion = ConfigurationManager.AppSettings("OLEDBConexion")
            End If

            Return _oLEDBConexion
        End Get
    End Property

    Public ReadOnly Property BaseDatosConexion() As String
        Get
            If String.IsNullOrEmpty(_baseDatosConexion) Then
                _baseDatosConexion = ConfigurationManager.AppSettings("BaseDatosConexion")
            End If

            Return _baseDatosConexion
        End Get
    End Property

    Public ReadOnly Property UsuarioConexion() As String
        Get
            If String.IsNullOrEmpty(_usuarioConexion) Then
                _usuarioConexion = ConfigurationManager.AppSettings("UsuarioConexion")
            End If

            Return _usuarioConexion
        End Get
    End Property

    Public ReadOnly Property PasswordConexion() As String
        Get
            If String.IsNullOrEmpty(_passwordConexion) Then
                _passwordConexion = ConfigurationManager.AppSettings("PasswordConexion")
            End If

            Return _passwordConexion
        End Get
    End Property



    Public ReadOnly Property Host() As String
        Get
            If String.IsNullOrEmpty(_host) Then
                _host = ConfigurationManager.AppSettings("Host")
            End If

            Return _host
        End Get
    End Property

    Public ReadOnly Property Port() As Integer
        Get
            If _port = 0 Then
                _port = Integer.Parse(ConfigurationManager.AppSettings("Port"))
            End If

            Return _port
        End Get
    End Property

    Public ReadOnly Property EnableSsl() As Boolean
        Get
            'If IsNothing(_enableSsl) Then
            _enableSsl = Boolean.Parse(ConfigurationManager.AppSettings("EnableSsl"))
            'End If

            Return _enableSsl
        End Get
    End Property

    Public ReadOnly Property MailAccount() As String
        Get
            If String.IsNullOrEmpty(_mailAccount) Then
                _mailAccount = ConfigurationManager.AppSettings("MailAccount")
            End If

            Return _mailAccount
        End Get
    End Property

    Public ReadOnly Property MailPassword() As String
        Get
            If String.IsNullOrEmpty(_mailpassword) Then
                _mailpassword = ConfigurationManager.AppSettings("MailPassword")
            End If

            Return _mailpassword
        End Get
    End Property



    Public ReadOnly Property Mode() As String
        Get
            If String.IsNullOrEmpty(_mode) Then
                _mode = ConfigurationManager.AppSettings("Mode").ToString.ToUpper()
            End If

            Return _mode
        End Get
    End Property

    Public ReadOnly Property ScheduledTimeExec() As String
        Get
            If String.IsNullOrEmpty(_scheduledTime) Then
                _scheduledTime = ConfigurationManager.AppSettings("ScheduledTime")
            End If

            Return _scheduledTime
        End Get
    End Property

    Public ReadOnly Property IntervalMinutes() As Integer
        Get
            If _intervalMinutes = 0 Then
                _intervalMinutes = Integer.Parse(ConfigurationManager.AppSettings("IntervalMinutes"))
            End If

            Return _intervalMinutes
        End Get
    End Property

    Public ReadOnly Property MuestraLog() As Boolean
        Get
            'If IsNothing(_muestraLog) Then
            _muestraLog = Boolean.Parse(ConfigurationManager.AppSettings("MuestraLog"))
            'End If

            Return _muestraLog
        End Get
    End Property

    Public ReadOnly Property PathLog() As String
        Get
            If String.IsNullOrEmpty(_pathLog) Then
                _pathLog = ConfigurationManager.AppSettings("PathLog")
            End If

            Return _pathLog
        End Get
    End Property

    Public ReadOnly Property LogName() As String
        Get
            If String.IsNullOrEmpty(_logName) Then
                _logName = ConfigurationManager.AppSettings("LogName")
            End If

            Return _logName
        End Get
    End Property



    Public Function GetDescription() As String
        Dim description As String = "Descripción Temporal"

        Try
            Dim executingAssembly As Assembly = Assembly.GetAssembly(GetType(ProjectInstaller))
            Dim targetDir As String = executingAssembly.Location
            Dim config As Configuration = ConfigurationManager.OpenExeConfiguration(targetDir)
            description = config.AppSettings.Settings("ServiceDescription").Value.ToString()

            Return description
        Catch ex As Exception
            Return description
        End Try
    End Function

    Public Function GetDisplayName() As String
        Dim displayName As String = "Nombre Temporal"

        Try
            Dim executingAssembly As Assembly = Assembly.GetAssembly(GetType(ProjectInstaller))
            Dim targetDir As String = executingAssembly.Location
            Dim config As Configuration = ConfigurationManager.OpenExeConfiguration(targetDir)
            displayName = config.AppSettings.Settings("ServiceDisplayName").Value.ToString()

            Return displayName
        Catch ex As Exception
            Return displayName
        End Try
    End Function

    Public Function GetServiceName() As String
        Dim serviceName As String = "NombreTemporal"

        Try
            Dim executingAssembly As Assembly = Assembly.GetAssembly(GetType(ProjectInstaller))
            Dim targetDir As String = executingAssembly.Location
            Dim config As Configuration = ConfigurationManager.OpenExeConfiguration(targetDir)
            serviceName = config.AppSettings.Settings("ServiceName").Value.ToString()

            Return serviceName
        Catch ex As Exception
            Return serviceName
        End Try
    End Function
End Module
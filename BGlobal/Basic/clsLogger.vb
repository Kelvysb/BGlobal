Imports Newtonsoft.Json

Public Class clsLoggerGroup

#Region "Declarations"
    Private strDate As String
    Private strFirstLog As String
    Private strLastLog As String
    Private strProcess As String
    Private objLogs As List(Of clsLogger)
#End Region

#Region "Constructor"
    Public Sub New()
        Try
            strDate = ""
            strFirstLog = ""
            strLastLog = ""
            strProcess = ""
            objLogs = New List(Of clsLogger)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "FFunctions"
    Public Function Clone() As clsLoggerGroup
        Try
            Return Deserialize(Serialize)
        Catch Ex As Exception
            Throw Ex
        End Try
    End Function

    Public Function Serialize() As String
        Dim objReturn As String
        Try
            objReturn = JsonConvert.SerializeObject(Me, Formatting.Indented)
            Return objReturn
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function Serialize(p_objclsLogginGroup As clsLoggerGroup) As String
        Dim objReturn As String
        Try
            objReturn = JsonConvert.SerializeObject(p_objclsLogginGroup, Formatting.Indented)
            Return objReturn
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function Deserialize(p_strJson As String) As clsLoggerGroup
        Dim objReturn As clsLoggerGroup
        Try
            objReturn = JsonConvert.DeserializeObject(Of clsLoggerGroup)(p_strJson)
            Return objReturn
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region "Properties"
    ''' <summary>
    ''' Process name
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonProperty("PROCESS")>
    Public Property Process() As String
        Get
            Return strProcess
        End Get
        Set(ByVal p_strProcess As String)
            strProcess = p_strProcess
        End Set
    End Property

    ''' <summary>
    ''' Day of the log file
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonProperty("LOGDATE")>
    Public Property LogDate() As String
        Get
            Return strDate
        End Get
        Set(ByVal p_strDate As String)
            strDate = p_strDate
        End Set
    End Property

    ''' <summary>
    ''' Date time of the first log file
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonProperty("FIRSTLOG")>
    Public Property FirstLog() As String
        Get
            Return strFirstLog
        End Get
        Set(ByVal p_strFirstLog As String)
            strFirstLog = p_strFirstLog
        End Set
    End Property

    ''' <summary>
    ''' Date time of the last log file
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonProperty("LASTLOG")>
    Public Property LastLog() As String
        Get
            Return strLastLog
        End Get
        Set(ByVal p_strLastLog As String)
            strLastLog = p_strLastLog
        End Set
    End Property

    ''' <summary>
    ''' Logs
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonProperty("LOGS")>
    Public Property Logs() As List(Of clsLogger)
        Get
            Return objLogs
        End Get
        Set(ByVal p_objLogs As List(Of clsLogger))
            objLogs = p_objLogs
        End Set
    End Property

#End Region

End Class

Public Class clsLogger

#Region "Declarations"
    Private strCode As String
    Private strProcess As String
    Private strDescription As String
    Private strSubroutine As String
    Private strCurrentClass As String
    Private intLine As Integer
    Private strUser As String
    Private strComputer As String
    Private blnIsError As Boolean
#End Region

#Region "Constructors"
    Public Sub New()
        Try
            strCode = ""
            strProcess = ""
            strDescription = ""
            strSubroutine = ""
            strCurrentClass = ""
            intLine = 0
            strUser = ""
            strComputer = ""
            blnIsError = False
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "functions"

    Public Function Clone() As clsLogger
        Try
            Return Deserialize(Serialize)
        Catch Ex As Exception
            Throw Ex
        End Try
    End Function


    Public Function Serialize() As String
        Dim objReturn As String
        Try
            objReturn = JsonConvert.SerializeObject(Me, Formatting.Indented)
            Return objReturn
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function Serialize(p_objclsLoggin As clsLogger) As String
        Dim objReturn As String
        Try
            objReturn = JsonConvert.SerializeObject(p_objclsLoggin, Formatting.Indented)
            Return objReturn
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function Deserialize(p_strJson As String) As clsLogger
        Dim objReturn As clsLogger
        Try
            objReturn = JsonConvert.DeserializeObject(Of clsLogger)(p_strJson)
            Return objReturn
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region "Properties"
    ''' <summary>
    ''' Log Code
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonProperty("CODE")>
    Public Property Code() As String
        Get
            Return strCode
        End Get
        Set(ByVal p_strCode As String)
            strCode = p_strCode
        End Set
    End Property

    ''' <summary>
    ''' Log Process
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonProperty("PROCESS")>
    Public Property Process() As String
        Get
            Return strProcess
        End Get
        Set(ByVal p_strProcess As String)
            strProcess = p_strProcess
        End Set
    End Property

    ''' <summary>
    ''' Log Description
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonProperty("DESCRIPTION")>
    Public Property Description() As String
        Get
            Return strDescription
        End Get
        Set(ByVal p_strDescription As String)
            strDescription = p_strDescription
        End Set
    End Property

    ''' <summary>
    ''' Log Subroutine
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonProperty("SUBROUTINE")>
    Public Property Subroutine() As String
        Get
            Return strSubroutine
        End Get
        Set(ByVal p_strSubroutine As String)
            strSubroutine = p_strSubroutine
        End Set
    End Property

    ''' <summary>
    ''' Log Class
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonProperty("CURRENTCLASS")>
    Public Property CurrentClass() As String
        Get
            Return strCurrentClass
        End Get
        Set(ByVal p_strCurrentClass As String)
            strCurrentClass = p_strCurrentClass
        End Set
    End Property

    ''' <summary>
    ''' Log Line
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonProperty("LINE")>
    Public Property Line() As Integer
        Get
            Return intLine
        End Get
        Set(ByVal p_intLine As Integer)
            intLine = p_intLine
        End Set
    End Property

    ''' <summary>
    ''' Log User
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonProperty("USER")>
    Public Property User() As String
        Get
            Return strUser
        End Get
        Set(ByVal p_strUser As String)
            strUser = p_strUser
        End Set
    End Property

    ''' <summary>
    ''' Log Computer
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonProperty("COMPUTER")>
    Public Property Computer() As String
        Get
            Return strComputer
        End Get
        Set(ByVal p_strComputer As String)
            strComputer = p_strComputer
        End Set
    End Property

    ''' <summary>
    ''' Log Is Error
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonProperty("ISERROR")>
    Public Property IsError() As Boolean
        Get
            Return blnIsError
        End Get
        Set(ByVal p_blnIsError As Boolean)
            blnIsError = p_blnIsError
        End Set
    End Property

#End Region

End Class

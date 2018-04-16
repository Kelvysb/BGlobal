Imports System.IO

Public Class BLogger
#Region "Declarations"
    Private Shared objInstance As BLogger
    Private strPath As String
#End Region

#Region "Constructor"
    Private Sub New(p_strPath As String)
        Try
            If p_strPath.EndsWith("\") = False Then
                p_strPath = p_strPath & "\"
            End If
            strPath = p_strPath
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Functions"
    Public Shared Sub sbInitialize(p_strPath As String)
        Try
            objInstance = New BLogger(p_strPath)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Sub sbLog(p_Exception As Exception)
        Try
            Call sbLog("", p_Exception)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Sub sbLog(p_strProcess As String, p_Exception As Exception)

        Dim strCode As String
        Dim strDescription As String
        Dim strCurrentClass As String
        Dim strSubRoutine As String
        Dim intLine As Integer
        Dim objAuxInner As Exception

        Try

            strCode = p_Exception.GetHashCode

            strDescription = ""

            objAuxInner = p_Exception

            If p_strProcess.Trim = "" AndAlso p_Exception.Source IsNot Nothing Then
                p_strProcess = p_Exception.Source
            End If

            strCurrentClass = ""
            strSubRoutine = ""

            Do While objAuxInner IsNot Nothing
                strDescription = strDescription & objAuxInner.Message & vbNewLine & "-------------" & objAuxInner.StackTrace & vbNewLine & vbNewLine

                Try
                    intLine = New StackTrace(objAuxInner).GetFrame(0).GetFileLineNumber
                Catch ex As Exception
                    intLine = 0
                End Try
                Try
                    strCurrentClass = objAuxInner.TargetSite.ReflectedType.Name
                Catch ex As Exception
                    strCurrentClass = ""
                End Try
                Try
                    strSubRoutine = objAuxInner.TargetSite.Name
                Catch ex As Exception
                    strSubRoutine = ""
                End Try

                objAuxInner = objAuxInner.InnerException
            Loop

            Call sbLog(p_strProcess, strCode, strDescription, strCurrentClass, strSubRoutine, intLine, "", "", True)

        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Sub sbLog(p_strProcess As String, p_strCode As String, p_strDescription As String, p_blnIsError As Boolean)
        Try
            Call sbLog(p_strProcess, p_strCode, p_strDescription, "", "", 0, "", "", p_blnIsError)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Sub sbLog(p_strProcess As String, p_strCode As String, p_strDescription As String, p_strSubRoutine As String, p_intLine As Integer, p_blnIsError As Boolean)
        Try
            Call sbLog(p_strProcess, p_strCode, p_strDescription, "", p_strSubRoutine, p_intLine, "", "", p_blnIsError)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Sub sbLog(p_strProcess As String, p_strCode As String, p_strDescription As String, p_strCurrentClass As String, p_strSubRoutine As String, p_intLine As Integer, p_blnIsError As Boolean)
        Try
            Call sbLog(p_strProcess, p_strCode, p_strDescription, p_strCurrentClass, p_strSubRoutine, p_intLine, "", "", p_blnIsError)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Sub sbLog(p_strProcess As String, p_strCode As String, p_strDescription As String, p_strCurrentClass As String, p_strSubRoutine As String, p_intLine As Integer, p_strUser As String, p_blnIsError As Boolean)
        Try
            Call sbLog(p_strProcess, p_strCode, p_strDescription, p_strCurrentClass, p_strSubRoutine, p_intLine, p_strUser, "", p_blnIsError)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Sub sbLog(p_strProcess As String, p_strCode As String, p_strDescription As String, p_strCurrentClass As String, p_strSubRoutine As String, p_intLine As Integer, p_strUser As String, p_strComputer As String, p_blnIsError As Boolean)

        Dim objFileRead As StreamReader
        Dim objFileWrite As StreamWriter
        Dim objLog As clsLoggerGroup
        Dim strAuxFileContent As String
        Dim strAuxFile As String

        Try

            If Directory.Exists(strPath) = False Then
                Directory.CreateDirectory(strPath)
            End If

            Try
                If p_strProcess.Trim = "" Then
                    p_strProcess = Reflection.Assembly.GetEntryAssembly.GetName.Name
                End If
            Catch ex As Exception
                p_strProcess = "Log"
            End Try

            If p_strUser.Trim = "" Then
                p_strUser = GlobalFunctions.fnUserName()
            End If

            If p_strComputer.Trim = "" Then
                p_strComputer = GlobalFunctions.fnComputerName()
            End If

            strAuxFile = strPath & p_strProcess & "_" & Now.ToString("yyyyMMdd") & ".log"

            If File.Exists(strAuxFile) Then
                objFileRead = New StreamReader(strAuxFile)
                strAuxFileContent = objFileRead.ReadToEnd
                objFileRead.Close()
                objFileRead.Dispose()
                objFileRead = Nothing
                Try
                    objLog = clsLoggerGroup.Deserialize(strAuxFileContent)
                Catch ex As Exception
                    objFileWrite = New StreamWriter(strAuxFile & ".corrupt", False)
                    objFileWrite.WriteLine(strAuxFileContent)
                    objFileWrite.Close()
                    objFileWrite.Dispose()
                    objFileWrite = Nothing
                    objLog = New clsLoggerGroup With {.Process = p_strProcess, .LogDate = Now.ToString("yyyyMMdd"), .FirstLog = Now.ToString("yyyy-MM-dd HH:mm:ss"), .LastLog = Now.ToString("yyyy-MM-dd HH:mm:ss")}
                End Try
            Else
                objLog = New clsLoggerGroup With {.Process = p_strProcess, .LogDate = Now.ToString("yyyyMMdd"), .FirstLog = Now.ToString("yyyy-MM-dd HH:mm:ss"), .LastLog = Now.ToString("yyyy-MM-dd HH:mm:ss")}
            End If


            objLog.Logs.Add(New clsLogger() With {.Code = p_strCode,
                                                 .Process = p_strProcess,
                                                 .Computer = p_strComputer,
                                                 .CurrentClass = p_strCurrentClass,
                                                 .Description = p_strDescription,
                                                 .Line = p_intLine,
                                                 .Subroutine = p_strSubRoutine,
                                                 .User = p_strUser,
                                                 .IsError = p_blnIsError})

            objLog.LastLog = Now.ToString("yyyy-MM-dd HH:mm:ss")

            strAuxFileContent = objLog.Serialize

            objFileWrite = New StreamWriter(strAuxFile, False)
            objFileWrite.WriteLine(strAuxFileContent)
            objFileWrite.Close()
            objFileWrite.Dispose()
            objFileWrite = Nothing

        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Properties"
    Public Shared ReadOnly Property Instance As BLogger
        Get
            If objInstance IsNot Nothing Then
                Return objInstance
            Else
                Throw New Exception("Must be initialized.")
            End If
        End Get
    End Property
#End Region
End Class

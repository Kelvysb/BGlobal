
'Copyright 2018 Kelvys B. Pantaleão

'This file is part of BGlobal

'BGlobal Is free software: you can redistribute it And/Or modify
'it under the terms Of the GNU General Public License As published by
'the Free Software Foundation, either version 3 Of the License, Or
'(at your option) any later version.

'This program Is distributed In the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty Of
'MERCHANTABILITY Or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License For more details.

'You should have received a copy Of the GNU General Public License
'along with this program.  If Not, see <http://www.gnu.org/licenses/>.


'Este arquivo é parte Do programa BGlobal

'BGlobal é um software livre; você pode redistribuí-lo e/ou 
'modificá-lo dentro dos termos da Licença Pública Geral GNU como 
'publicada pela Fundação Do Software Livre (FSF); na versão 3 da 
'Licença, ou(a seu critério) qualquer versão posterior.

'Este programa é distribuído na esperança de que possa ser  útil, 
'mas SEM NENHUMA GARANTIA; sem uma garantia implícita de ADEQUAÇÃO
'a qualquer MERCADO ou APLICAÇÃO EM PARTICULAR. Veja a
'Licença Pública Geral GNU para maiores detalhes.

'Você deve ter recebido uma cópia da Licença Pública Geral GNU junto
'com este programa, Se não, veja <http://www.gnu.org/licenses/>.

'GitHub: https://github.com/Kelvysb/BGlobal

Imports Newtonsoft.Json
Imports System.Net
Imports Newtonsoft.Json.Linq
Imports System.IO

Public Class RestWebService

#Region "Declarations"
    Private strURL As String
    Private CONS_TIMEOUT_UPDATE As Integer
#End Region

#Region "Constructor"
    Public Sub New(p_strUrl As String)
        Try
            strURL = p_strUrl
            CONS_TIMEOUT_UPDATE = 30
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Sub New(p_strUrl As String, p_inttimeout As Integer)
        Try
            strURL = p_strUrl
            CONS_TIMEOUT_UPDATE = p_inttimeout
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Functions"

    Public Function fnHttpWebRequestPost(Of T)(p_objInput As T, p_objMethod As String) As String

        Dim strReturn As String
        Dim objHttpClient As WebClient
        Dim strAuxRequest As String
        Dim strAuxResult As String
        Dim strAuxReturn As String
        Dim objReturn As JObject
        Dim objResponse As HttpWebResponse
        Dim objContent As HttpWebRequest
        Dim straAuxContent As String
        Dim objencoding As New Text.UTF8Encoding()
        Dim objBytes As Byte()


        Try

            strReturn = ""

            objHttpClient = New WebClient

            strAuxResult = p_objMethod & "Result"

            strAuxRequest = strURL & "/" & p_objMethod

            straAuxContent = JsonConvert.SerializeObject(p_objInput)
            objBytes = objencoding.GetBytes(straAuxContent)

            objContent = DirectCast(HttpWebRequest.Create(New Uri(strAuxRequest)), HttpWebRequest)
            objContent.Method = "POST"
            objContent.ContentType = "text/plain;charset=utf-8"
            objContent.ContentLength = objBytes.Length

            Using requestStream As Stream = objContent.GetRequestStream()
                requestStream.Write(objBytes, 0, objBytes.Length)
            End Using

            objResponse = objContent.GetResponse()

            Using reader As StreamReader = New StreamReader(objResponse.GetResponseStream)
                strAuxReturn = reader.ReadToEnd
            End Using

            objReturn = JObject.Parse(strAuxReturn)

            strReturn = objReturn(strAuxResult)

            Return strReturn

        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function fnHttpWebRequestGet(Of T)(p_objInput As T, p_objMethod As String) As String

        Dim strReturn As String
        Dim objHttpClient As WebClient
        Dim strAuxRequest As String
        Dim strAuxReturn As String
        Dim strAuxResult As String
        Dim objReturn As JObject
        Dim straAuxContent As String
        Dim objStreamResult As Stream

        Try

            strReturn = ""

            objHttpClient = New WebClient()

            strAuxResult = p_objMethod & "Result"

            straAuxContent = JsonConvert.SerializeObject(p_objInput)

            strAuxRequest = strURL & "/" & p_objMethod & "/" & straAuxContent

            objStreamResult = objHttpClient.OpenRead(New Uri(strAuxRequest))

            Using reader As StreamReader = New StreamReader(objStreamResult)
                strAuxReturn = reader.ReadToEnd
            End Using

            objReturn = JObject.Parse(strAuxReturn)

            strReturn = objReturn(strAuxResult)

            Return strReturn

        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function fnHttpWebRequestGet(p_objMethod As String) As String

        Dim strReturn As String
        Dim objHttpClient As WebClient
        Dim strAuxRequest As String
        Dim strAuxReturn As String
        Dim strAuxResult As String
        Dim objReturn As JObject
        Dim objStreamResult As Stream

        Try

            strReturn = ""

            objHttpClient = New WebClient()

            strAuxResult = p_objMethod & "Result"

            strAuxRequest = strURL & "/" & p_objMethod

            objStreamResult = objHttpClient.OpenRead(New Uri(strAuxRequest))

            Using reader As StreamReader = New StreamReader(objStreamResult)
                strAuxReturn = reader.ReadToEnd
            End Using

            objReturn = JObject.Parse(strAuxReturn)

            strReturn = objReturn(strAuxResult)

            Return strReturn

        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

End Class



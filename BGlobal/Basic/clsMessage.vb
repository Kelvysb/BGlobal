
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
''' <summary>
''' General communication message class
''' </summary>
Public Class clsMessage

#Region "Declarations"
    Private strType As String
    Private objParameters As List(Of clsParameter)
    Private strData As List(Of String)
    Private strStatus As String
    Private strMessage As String
#End Region

#Region "Constructors"
    Public Sub New()
        Try
            strType = ""
            objParameters = New List(Of clsParameter)
            strData = New List(Of String)
            strStatus = ""
            strMessage = ""
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub New(p_xmlEntrada As XDocument)
        Try
            Call sbFromXml(p_xmlEntrada)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Functions and Subroutines"
    ''' <summary>
    ''' Convert to XML
    ''' </summary>
    ''' <remarks></remarks>
    Public Function fnToXml() As XDocument
        Dim Retorno As XDocument
        Try

            Retorno = New XDocument(<Message></Message>)

            Retorno.<Message>.First.Add(<Type><%= strType.Trim %></Type>)

            Retorno.<Message>.First.Add(<Parameters></Parameters>)
            For i = 0 To objParameters.Count - 1
                Retorno.<Message>.<Parameters>.First.Add(objParameters(i).fnToXml.Elements)
            Next


            Retorno.<Message>.First.Add(<Data></Data>)
            For i = 0 To strData.Count - 1
                Retorno.<Message>.<Data>.First.Add(<Dado><%= strData(i).Trim %></Dado>)
            Next

            Retorno.<Message>.First.Add(<Status><%= strStatus.Trim %></Status>)
            Retorno.<Message>.First.Add(<Message><%= strMessage.Trim %></Message>)
            Return Retorno

        Catch Ex As Exception
            Throw Ex
        End Try
    End Function

    Private Sub sbFromXml(p_xmlEntrada As XDocument)
        Try

            strType = p_xmlEntrada.<Message>.<Type>.Value
            objParameters = New List(Of clsParameter)
            For i = 0 To p_xmlEntrada.<Message>.<Parameters>.<Parameter>.Count - 1
                objParameters.Add(New clsParameter(New XDocument(p_xmlEntrada.<Message>.<Parameters>.<Parameter>(i))))
            Next
            strData = New List(Of String)
            For i = 0 To p_xmlEntrada.<Message>.<Data>.<Dado>.Count - 1
                strData.Add(p_xmlEntrada.<Message>.<Data>.<Dado>(i).Value)
            Next
            strStatus = p_xmlEntrada.<Message>.<Status>.Value
            strMessage = p_xmlEntrada.<Message>.<Message>.Value

        Catch Ex As Exception
            Throw Ex
        End Try
    End Sub

    ''' <summary>
    ''' Clone the object
    ''' </summary>
    ''' <remarks></remarks>
    Public Function Clone() As clsMessage
        Try

            Return New clsMessage(fnToXml)

        Catch Ex As Exception
            Throw Ex
        End Try
    End Function

    ''' <summary>
    ''' Get single parameter
    ''' </summary>
    ''' <param name="p_strKey">Parameter key</param>
    ''' <returns></returns>
    Private Function fnGetParameter(p_strKey As String) As clsParameter

        Dim Retorno As clsParameter

        Try

            Retorno = objParameters.Find(Function(Param As clsParameter) Param.Key.Trim.ToUpper = p_strKey.Trim.ToUpper)

            If Retorno Is Nothing Then
                Throw New Exception("Parameter not found: " & p_strKey)
            End If

            Return Retorno

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' Set single parameter
    ''' </summary>
    ''' <param name="p_strKey">Parameter key</param>
    ''' <param name="p_strType">Parameter Type</param>
    ''' <param name="p_strValue">Parameter value</param>
    Private Sub sbSetParameter(p_strKey As String, p_strValue As String, p_strType As String)

        Dim objAuxParameter As clsParameter

        Try

            objAuxParameter = objParameters.Find(Function(Param As clsParameter) Param.Key.Trim.ToUpper = p_strKey.Trim.ToUpper)

            If objAuxParameter IsNot Nothing Then
                objAuxParameter.Value = p_strValue
                objAuxParameter.Type = p_strType
            Else
                objParameters.Add(New clsParameter() With {.Key = p_strKey, .Value = p_strValue, .Type = p_strType})
            End If

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Set single parameter
    ''' </summary>
    ''' <param name="p_objParameter">Parameter</param>
    Private Sub sbSetParameter(p_objParameter As clsParameter)
        Try
            sbSetParameter(p_objParameter.Key, p_objParameter.Value, p_objParameter.Type)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Function fnSerialize() As String

        Dim Retorno As String

        Try

            Retorno = JsonConvert.SerializeObject(Me, Formatting.Indented)

            Return Retorno

        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function fnSerialize(p_objMessage As clsMessage) As String

        Dim Retorno As String

        Try

            Retorno = JsonConvert.SerializeObject(p_objMessage, Formatting.Indented)

            Return Retorno

        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function fnDeserialize(p_strJson As String) As clsMessage

        Dim Retorno As clsMessage

        Try

            Retorno = JsonConvert.DeserializeObject(Of clsMessage)(p_strJson)

            Return Retorno

        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Sub sbAddParameter(p_strKey As String)
        Try
            Call sbAddParameter(p_strKey, "", "")
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Sub sbAddParameter(p_strKey As String, p_strValue As String)
        Try
            Call sbAddParameter(p_strKey, p_strValue, "")
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Sub sbAddParameter(p_strKey As String, p_strValue As String, p_strType As String)
        Try
            Call sbAddParameter(New clsParameter() With {.Key = p_strKey, .Value = p_strValue, .Type = p_strType})
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Sub sbAddParameter(p_objParameter As clsParameter)
        Try

            If p_objParameter.Key.Trim <> "" Then
                If objParameters.Find(Function(Param As clsParameter) Param.Key.Trim.ToUpper = p_objParameter.Key.Trim.ToUpper) Is Nothing Then
                    objParameters.Add(p_objParameter)
                Else
                    Throw New Exception("Parameter already exists.")
                End If
            Else
                Throw New Exception("The Key can't be found.")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Overrides Function ToString() As String

        Dim strReturn As String

        Try

            strReturn = strType.Trim
            For Each parameter As clsParameter In objParameters
                strReturn = strReturn & ";" & parameter.ToString.Trim
            Next
            For Each Data As String In strData
                strReturn = strReturn & ";" & Data.Trim
            Next
            strReturn = strReturn & ";" & strStatus.Trim
            strReturn = strReturn & ";" & strMessage.Trim

            Return strReturn

        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Sub sbSignMessage(p_strSignature As String)

        Dim strKey As String
        Dim strAuxMessage As String
        Dim strHash As String

        Try

            strKey = p_strSignature.Trim.ToUpper
            sbAddParameter("HASH", strKey)
            strAuxMessage = ToString()
            strHash = GlobalFunctions.fnHashMD5(strAuxMessage).Trim.ToUpper
            Parameters("HASH").Value = strHash

        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Function fnCheckSignature(p_strSignature As String) As Boolean


        Dim blnReturn As Boolean
        Dim strHashMensagem As String
        Dim strChave As String
        Dim strAuxMensagem As String
        Dim strHash As String

        Try

            blnReturn = False

            strChave = p_strSignature.Trim.ToUpper

            strHashMensagem = Parameters("HASH").Value.Trim.ToUpper
            Parameters("HASH").Value = strChave.Trim.ToUpper

            strAuxMensagem = ToString()

            strHash = GlobalFunctions.fnHashMD5(strAuxMensagem).Trim.ToUpper

            If strHash.Trim.ToUpper = strHashMensagem.Trim.ToUpper Then
                blnReturn = True
            End If

            Parameters("HASH").Value = strHashMensagem

        Catch ex As Exception
            blnReturn = False
        End Try

        Return blnReturn

    End Function

#End Region

#Region "Properties"
    ''' <summary>
    ''' Type da solicitação
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonProperty("TYPE")>
    Public Property Type() As String
        Get
            Return strType
        End Get
        Set(ByVal p_strType As String)
            strType = p_strType
        End Set
    End Property

    ''' <summary>
    ''' Type da solicitação
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonProperty("PARAMETERS")>
    Public ReadOnly Property Parameters() As List(Of clsParameter)
        Get
            Return objParameters
        End Get
    End Property

    ''' <summary>
    ''' Type da solicitação
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Parameters(p_strKey As String) As clsParameter
        Get
            Return fnGetParameter(p_strKey)
        End Get
        Set(ByVal p_objParameters As clsParameter)
            Call sbSetParameter(p_strKey, p_objParameters.Value, p_objParameters.Type)
        End Set
    End Property

    ''' <summary>
    ''' Data da Soicitacao
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonProperty("DATA")>
    Public Property Data() As List(Of String)
        Get
            Return strData
        End Get
        Set(ByVal p_strData As List(Of String))
            strData = p_strData
        End Set
    End Property

    ''' <summary>
    ''' Status
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonProperty("STATUS")>
    Public Property Status() As String
        Get
            Return strStatus
        End Get
        Set(ByVal p_strStatus As String)
            strStatus = p_strStatus
        End Set
    End Property

    ''' <summary>
    ''' Message
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonProperty("MESSAGE")>
    Public Property Message() As String
        Get
            Return strMessage
        End Get
        Set(ByVal p_strMessage As String)
            strMessage = p_strMessage
        End Set
    End Property
#End Region

End Class

''' <summary>
''' Parameter Class
''' </summary>
Public Class clsParameter

#Region "Declarations"
    Private strKey As String
    Private strType As String
    Private strValue As String
#End Region

#Region "Constructors"
    Public Sub New()
        Try
            strKey = ""
            strType = ""
            strValue = ""
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub New(p_xmlInput As XDocument)
        Try
            Call sbFromXml(p_xmlInput)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Functions and subroutines"
    ''' <summary>
    ''' Converte o objeto em Xml
    ''' </summary>
    ''' <remarks></remarks>
    Public Function fnToXml() As XDocument
        Dim xmlReturn As XDocument
        Try

            xmlReturn = New XDocument(<Parameter></Parameter>)

            xmlReturn.<Parameter>.First.Add(<Key><%= strKey.Trim %></Key>)
            xmlReturn.<Parameter>.First.Add(<Type><%= strType.Trim %></Type>)
            xmlReturn.<Parameter>.First.Add(<Value><%= strValue.Trim %></Value>)
            Return xmlReturn

        Catch Ex As Exception
            Throw Ex
        End Try
    End Function

    ''' <summary>
    ''' Cria o objeto apartir de um xml
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub sbFromXml(p_xmlInput As XDocument)
        Try

            strKey = p_xmlInput.<Parameter>.<Key>.Value
            strType = p_xmlInput.<Parameter>.<Type>.Value
            strValue = p_xmlInput.<Parameter>.<Value>.Value

        Catch Ex As Exception
            Throw Ex
        End Try
    End Sub

    ''' <summary>
    ''' Cria um clone do objeto
    ''' </summary>
    ''' <remarks></remarks>
    Public Function Clone() As clsParameter
        Try

            Return New clsParameter(fnToXml)

        Catch Ex As Exception
            Throw Ex
        End Try
    End Function

    Public Overrides Function ToString() As String

        Dim strRetorno As String

        Try

            strRetorno = strKey.Trim
            strRetorno = strRetorno & ";" & strType.Trim
            strRetorno = strRetorno & ";" & strValue.Trim

            Return strRetorno

        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

#Region "Properties"
    ''' <summary>
    ''' Parameter Key
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonProperty("KEY")>
    Public Property Key() As String
        Get
            Return strKey
        End Get
        Set(ByVal p_strKey As String)
            strKey = p_strKey
        End Set
    End Property

    ''' <summary>
    ''' Parameter Type
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonProperty("TYPE")>
    Public Property Type() As String
        Get
            Return strType
        End Get
        Set(ByVal p_strType As String)
            strType = p_strType
        End Set
    End Property

    ''' <summary>
    ''' Parameter Value
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonProperty("VALUE")>
    Public Property Value() As String
        Get
            Return strValue
        End Get
        Set(ByVal p_strValue As String)
            strValue = p_strValue
        End Set
    End Property

#End Region

End Class


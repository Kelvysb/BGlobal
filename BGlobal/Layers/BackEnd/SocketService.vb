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

Imports System.Net
Imports System.Net.Sockets
Imports System.Threading
Imports System.Text

Public Class SocketServer

#Region "Declarations"
    Private objIp As Net.IPAddress
    Private intPort As Integer
    Private objThread As Threading.Thread
    Private objThreadResponse As Threading.Thread
    Private objAllDone As ManualResetEvent
    Private objCancel As CancellationTokenSource
#End Region

#Region "Constructor"
    Public Sub New()
        Try
            objCancel = New CancellationTokenSource()
            objAllDone = New ManualResetEvent(False)
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Events"
    Public Event evReceive(sender As Object, message As String, objReturn As Socket)
    Public Event evError(sender As Object, ex As Exception)
#End Region

#Region "Functions"
    Public Sub sbStart(p_strIp As String, p_intPort As Integer)
        Try

            objIp = IPAddress.Parse(p_strIp)
            intPort = p_intPort

            objThread = New Thread(Sub()
                                       Call sbStartListening()
                                   End Sub)

            objThread.Start()

        Catch ex As Exception
            RaiseEvent evError(Me, ex)
        End Try
    End Sub

    Public Sub sbResponse(p_objHandler As Socket, p_strDados As String)
        Try


            objThreadResponse = New Thread(Sub()
                                               Call sbSend(p_objHandler, p_strDados)
                                           End Sub)

            objThreadResponse.Start()

        Catch ex As Exception
            RaiseEvent evError(Me, ex)
        End Try
    End Sub

    Private Sub sbStartListening()

        Dim objBytes(1024) As Byte
        Dim objLocalEndPoint As IPEndPoint
        Dim objListener As Socket

        Try

            ' Data buffer for incoming data.
            objLocalEndPoint = New IPEndPoint(objIp, intPort)

            'Create a TCP/IP socket.
            objListener = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)

            'Bind the socket to the local endpoint And listen for incoming connections.
            objListener.Bind(objLocalEndPoint)
            objListener.Listen(100)

            Do While Not objCancel.IsCancellationRequested
                'Set the event to nonsignaled state.
                Try
                    objAllDone.Reset()
                Catch ex As Exception
                    objAllDone = New ManualResetEvent(False)
                    objAllDone.Reset()
                End Try

                'Start an asynchronous socket to listen for connections.
                objListener.BeginAccept(New AsyncCallback(AddressOf sbAcceptCallback), objListener)

                'Wait until a connection Is made before continuing.
                objAllDone.WaitOne(New TimeSpan(0, 0, 10)) 'objCancel.Token)

            Loop

        Catch CancEx As OperationCanceledException

        Catch ex As Exception
            RaiseEvent evError(Me, ex)
        End Try

    End Sub

    Private Sub sbAcceptCallback(ar As IAsyncResult)

        Dim objListener As Socket
        Dim objHandler As Socket
        Dim objState As clsStateObject

        Try
            'Signal the main thread to continue.
            objAllDone.Set()

            'Get the socket that handles the client request.
            objListener = DirectCast(ar.AsyncState, Socket)
            objHandler = objListener.EndAccept(ar)

            'Create the state object.
            objState = New clsStateObject()

            objState.WorkSocket = objHandler
            objHandler.BeginReceive(objState.Buffer, 0, objState.BufferSize, 0, New AsyncCallback(AddressOf sbReadCallback), objState)

        Catch ex As Exception
            RaiseEvent evError(Me, ex)
        End Try
    End Sub

    Private Sub sbReadCallback(ar As IAsyncResult)

        Dim objContent As String
        Dim objState As clsStateObject
        Dim objHandler As Socket
        Dim intBytesRead As Integer
        Dim strResponse As String
        Dim strAuxString As String

        Try


            objContent = String.Empty
            'Retrieve the state object And the handler socket
            'from the asynchronous state object.
            objState = DirectCast(ar.AsyncState, clsStateObject)
            objHandler = objState.WorkSocket

            'Read data from the client socket. 
            intBytesRead = objHandler.EndReceive(ar)


            If intBytesRead > 0 Then

                'There might be more data, so store the data received so far.
                objState.Sb.Append(Encoding.UTF8.GetString(objState.Buffer, 0, intBytesRead))

                strAuxString = objState.Sb.ToString()
                If strAuxString.EndsWith("/n") Then
                    strResponse = strAuxString.Substring(0, strAuxString.Length - 2)
                    RaiseEvent evReceive(Me, strResponse, objHandler)
                Else
                    'Get the rest of the data.
                    objHandler.BeginReceive(objState.Buffer, 0, objState.BufferSize, 0, New AsyncCallback(AddressOf sbReadCallback), objState)
                End If


            Else

                strResponse = ""
                'All the data has arrived put it in strResponse.
                If objState.Sb.Length > 1 Then
                    strResponse = objState.Sb.ToString()
                End If

                'Signal that all bytes have been received.
                RaiseEvent evReceive(Me, strResponse, objHandler)

            End If

        Catch ex As Exception
            RaiseEvent evError(Me, ex)
        End Try
    End Sub

    Private Sub sbSend(handler As Socket, data As String)

        Dim objByteData As Byte()

        Try

            'Convert the string data to byte data using ASCII encoding.
            objByteData = Encoding.UTF8.GetBytes(data)

            'Begin sending the data to the remote device.
            handler.BeginSend(objByteData, 0, objByteData.Length, 0, New AsyncCallback(AddressOf sbSendCallback), handler)

        Catch ex As Exception
            RaiseEvent evError(Me, ex)
        End Try
    End Sub

    Private Sub sbSendCallback(ar As IAsyncResult)

        Dim objHandler As Socket
        Dim intBytesSent As Integer

        Try
            'Retrieve the socket from the state object.
            objHandler = DirectCast(ar.AsyncState, Socket)

            'Complete sending the data to the remote device.
            intBytesSent = objHandler.EndSend(ar)

            objHandler.Shutdown(SocketShutdown.Both)

            objHandler.Close()

        Catch ex As Exception
            RaiseEvent evError(Me, ex)
        End Try
    End Sub

    Public Sub Shutdown()
        Try
            objCancel.Cancel()
            objAllDone.Dispose()
        Catch ex As Exception

        End Try
    End Sub
#End Region

End Class

Public Class SocketClient

#Region "Declarations"
    Private intPort As Integer
    Private objIp As IPAddress
    Private objConnectDone As ManualResetEvent
    Private objSendDone As ManualResetEvent
    Private objReceiveDone As ManualResetEvent
    Private strResponse As String
    Private objThread As Thread
#End Region

#Region "Cosntructor"
    Public Sub New(p_strIP As String, p_intPort As Integer)
        Try

            intPort = p_intPort
            objIp = IPAddress.Parse(p_strIP)

            objConnectDone = New ManualResetEvent(False)
            objSendDone = New ManualResetEvent(False)
            objReceiveDone = New ManualResetEvent(False)
            strResponse = String.Empty

        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Events"
    Public Event evError(sender As Object, Ex As Exception)
    Public Event evReceive(sender As Object, message As String)
#End Region

#Region "Functions"
    Public Sub sbSend(p_strData As String)
        Try
            objThread = New Thread(Sub()
                                       Call sbSendData(p_strData)
                                   End Sub)
            objThread.Start()
        Catch ex As Exception
            RaiseEvent evError(Me, ex)
        End Try
    End Sub

    Private Sub sbSendData(p_strData As String)

        Dim objRemoteEP As IPEndPoint
        Dim objClient As Socket

        Try

            'Establish the remote endpoint for the socket.
            objRemoteEP = New IPEndPoint(objIp, intPort)

            'Create a TCP/IP socket.
            objClient = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)

            'Connect to the remote endpoint.
            objClient.BeginConnect(objRemoteEP, New AsyncCallback(AddressOf sbConnectCallback), objClient)
            objConnectDone.WaitOne()

            'Send data to the remote device.
            sbSend(objClient, p_strData)
            objSendDone.WaitOne()

            'Receive the strResponse from the remote device.
            sbReceive(objClient)
            objReceiveDone.WaitOne()

            'Write the strResponse to the console.
            RaiseEvent evReceive(Me, strResponse)

            'Release the socket.
            objClient.Shutdown(SocketShutdown.Both)

            objClient.Close()

        Catch ex As Exception
            RaiseEvent evError(Me, ex)
        End Try
    End Sub

    Private Sub sbConnectCallback(ar As IAsyncResult)

        Dim objClient As Socket

        Try

            'Retrieve the socket from the state object.
            objClient = DirectCast(ar.AsyncState, Socket)

            'Complete the connection.
            objClient.EndConnect(ar)

            'Signal that the connection has been made.
            objConnectDone.Set()

        Catch e As Exception
            RaiseEvent evError(Me, e)
        End Try

    End Sub

    Private Sub sbReceive(client As Socket)

        Dim objState As clsStateObject = New clsStateObject()

        Try

            'Create the state object.
            objState = New clsStateObject()
            objState.WorkSocket = client

            'Begin receiving the data from the remote device.
            client.BeginReceive(objState.Buffer, 0, objState.BufferSize, 0, New AsyncCallback(AddressOf sbReceiveCallback), objState)

        Catch e As Exception
            RaiseEvent evError(Me, e)
        End Try

    End Sub

    Private Sub sbReceiveCallback(ar As IAsyncResult)

        Dim objState As clsStateObject
        Dim objClient As Socket
        Dim intBytesRead As Integer

        Try

            'from the asynchronous state object.
            objState = DirectCast(ar.AsyncState, clsStateObject)
            objClient = objState.WorkSocket

            'Read data from the remote device.
            intBytesRead = objClient.EndReceive(ar)

            If intBytesRead > 0 Then

                'There might be more data, so store the data received so far.
                objState.Sb.Append(Encoding.UTF8.GetString(objState.Buffer, 0, intBytesRead))


                'Get the rest of the data.
                objClient.BeginReceive(objState.Buffer, 0, objState.BufferSize, 0, New AsyncCallback(AddressOf sbReceiveCallback), objState)

            Else

                'All the data has arrived put it in strResponse.


                strResponse = ""

                If objState.Sb.Length > 1 Then
                    strResponse = objState.Sb.ToString()
                End If

                'Signal that all bytes have been received.
                objReceiveDone.Set()

            End If

        Catch ex As Exception
            RaiseEvent evError(Me, ex)
        End Try

    End Sub

    Private Sub sbSend(client As Socket, data As String)

        Dim byteData As Byte()

        Try
            byteData = Encoding.ASCII.GetBytes(data)

            client.BeginSend(byteData, 0, byteData.Length, 0, New AsyncCallback(AddressOf sbSendCallback), client)

        Catch ex As Exception
            RaiseEvent evError(Me, ex)
        End Try
    End Sub

    Private Sub sbSendCallback(ar As IAsyncResult)

        Dim objClient As Socket
        Dim intBytesSent As Integer
        Try

            'Retrieve the socket from the state object.
            objClient = DirectCast(ar.AsyncState, Socket)

            'Complete sending the data to the remote device.
            intBytesSent = objClient.EndSend(ar)
            objClient.Shutdown(SocketShutdown.Send)

            'Signal that all bytes have been sent.
            objSendDone.Set()

        Catch ex As Exception
            RaiseEvent evError(Me, ex)
        End Try

    End Sub
#End Region

End Class

Friend Class clsStateObject

#Region "Declarations"
    Private objWorkSocket As Socket
    Private intBufferSize As Integer
    Private objBuffer(1024) As Byte
    Private objSb As StringBuilder

#End Region

#Region "Constructor"
    Public Sub New()
        Try
            objWorkSocket = Nothing
            intBufferSize = 1024
            objSb = New StringBuilder()
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Properties"
    Public Property WorkSocket() As Socket
        Get
            Return objWorkSocket
        End Get
        Set(ByVal value As Socket)
            objWorkSocket = value
        End Set
    End Property
    Public Property BufferSize() As Integer
        Get
            Return intBufferSize
        End Get
        Set(ByVal value As Integer)
            intBufferSize = value
        End Set
    End Property
    Public Property Buffer() As Byte()
        Get
            Return objBuffer
        End Get
        Set(ByVal value As Byte())
            objBuffer = value
        End Set
    End Property
    Public Property Sb() As StringBuilder
        Get
            Return objSb
        End Get
        Set(ByVal value As StringBuilder)
            objSb = value
        End Set
    End Property
#End Region

End Class

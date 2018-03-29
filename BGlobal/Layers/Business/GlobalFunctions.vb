
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

Imports System.IO
Imports System.Security.Cryptography
Imports System.Reflection
Imports System.Text
Imports System.Drawing.Imaging
Imports System.Drawing
Imports System.Net

Public Class GlobalFunctions

#Region "Constructor"
    Private Sub New()

    End Sub
#End Region

#Region "Functions"
    Public Shared Function fnBreakLine(ByVal p_strText As String, ByVal p_intLimitPerLine As Integer) As List(Of String)

        Dim objReturn As List(Of String)
        Dim strAuxString As List(Of String)
        Dim strAuxLine As String
        Dim i As Integer

        Try

            objReturn = New List(Of String)

            If p_strText.Trim <> "" Then

                strAuxString = p_strText.Split(" ").ToList

                strAuxLine = ""

                For i = 0 To strAuxString.Count - 1

                    If strAuxString(i).Trim.Length > p_intLimitPerLine And strAuxLine.Trim <> "" Then

                        objReturn.Add(strAuxLine.Trim)
                        strAuxLine = ""

                        objReturn.Add(strAuxString(i).Trim.Substring(0, p_intLimitPerLine - 1) & "-")
                        strAuxLine = strAuxString(i).Trim.Substring(p_intLimitPerLine - 1)

                    ElseIf strAuxString(i).Trim.Length > p_intLimitPerLine And strAuxLine.Trim = "" Then

                        objReturn.Add(strAuxString(i).Trim.Substring(0, p_intLimitPerLine - 1) & "-")
                        strAuxLine = strAuxString(i).Trim.Substring(p_intLimitPerLine - 1)

                    ElseIf ((strAuxLine & strAuxString(i)).Trim.Length > p_intLimitPerLine) And strAuxLine.Trim <> "" Then



                        objReturn.Add(strAuxLine.Trim)

                        strAuxLine = ""

                        strAuxLine = strAuxLine & strAuxString(i) & " "

                    Else

                        strAuxLine = strAuxLine & strAuxString(i) & " "

                    End If

                Next

                If strAuxLine.Trim <> "" Then
                    objReturn.Add(strAuxLine.Trim)
                End If

            End If

            Return objReturn

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function fnHashMD5(ByVal p_strInput As String) As String

        Dim Retorno As String
        Dim objMD5Provider As MD5CryptoServiceProvider
        Dim bytHash() As Byte

        Try

            Retorno = String.Empty

            objMD5Provider = New MD5CryptoServiceProvider
            bytHash = objMD5Provider.ComputeHash(System.Text.Encoding.UTF8.GetBytes(p_strInput))
            Retorno = BitConverter.ToString(bytHash).Replace("-", String.Empty)

            Return Retorno

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function fnHashMD5File(ByVal p_strFilePath As String, Optional ByVal blnRaw_output As Boolean = False) As String

        Dim fs As FileStream
        Dim md5 As New MD5CryptoServiceProvider
        Dim strHashFinal As String = ""

        Try

            If Not System.IO.File.Exists(p_strFilePath) Then
                Throw New Exception("File not found.")
            End If

            fs = New System.IO.FileStream(p_strFilePath, IO.FileMode.Open, FileAccess.Read)

            md5.ComputeHash(fs)
            fs.Close()

            For i As Integer = 0 To md5.Hash.Length - 1
                If blnRaw_output Then
                    strHashFinal += Convert.ToChar(md5.Hash(i))
                Else
                    strHashFinal += Convert.ToString(md5.Hash(i).ToString("x2"))
                End If
            Next

            Return strHashFinal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function fnHashSHA1(ByVal p_strInput As String) As String

        Dim Retorno As String
        Dim objSHA1Provider As SHA1CryptoServiceProvider
        Dim bytHash() As Byte

        Try

            Retorno = String.Empty

            objSHA1Provider = New SHA1CryptoServiceProvider
            bytHash = objSHA1Provider.ComputeHash(System.Text.Encoding.ASCII.GetBytes(p_strInput))

            For Each b As Byte In bytHash
                Retorno += b.ToString("x2")
            Next

            Return Retorno

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function fnHashSHA1base64(ByVal p_strInput As String) As String

        Dim Retorno As String
        Dim objSHA1Provider As SHA1CryptoServiceProvider
        Dim bytHash() As Byte

        Try

            Retorno = String.Empty

            objSHA1Provider = New SHA1CryptoServiceProvider
            bytHash = objSHA1Provider.ComputeHash(System.Text.Encoding.ASCII.GetBytes(p_strInput))

            Return Convert.ToBase64String(bytHash)

            Return Retorno

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function fnImageToString(p_imgImage As Image) As String

        Dim Retorno As String
        Dim buffer() As Byte

        Try

            Retorno = ""

            If p_imgImage Is Nothing = False Then

                Using ms As New MemoryStream

                    p_imgImage.Save(ms, ImageFormat.Jpeg)
                    ms.Flush()
                    ms.Position = 0
                    buffer = ms.ToArray
                    Retorno = fnByteToStringHex(buffer)

                End Using

            End If

            Return Retorno

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function fnStringToImage(p_strImage As String) As Image

        Dim Retorno As Image
        Dim buffer() As Byte
        Dim Conversor As ImageConverter

        Try
            Retorno = Nothing

            If String.IsNullOrEmpty(p_strImage) = False Then

                buffer = fnStringToByteHex(p_strImage)

                Conversor = New ImageConverter

                Retorno = Conversor.ConvertFrom(buffer)

            End If

            Return Retorno

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function fnByteToStringHex(ByVal Value As Byte()) As String

        Dim strTemp As StringBuilder

        Try

            strTemp = New StringBuilder(Value.Length * 2)

            For Each b As Byte In Value
                strTemp.Append(Conversion.Hex(b).PadLeft(2, "0"))
            Next

            Return strTemp.ToString()

        Catch ex As Exception
            Throw ex
        End Try

    End Function

    Public Shared Function fnStringToByteHex(ByVal Value As String) As Byte()

        Dim bytes() As Byte

        Dim i As Integer
        Dim counter As Integer

        Try

            ReDim bytes(0 To (Len(Value) / 2) - 1)

            If Value <> vbNullString Then

                For i = 1 To Len(Value) Step 2

                    bytes(counter) = CDbl(Val("&H" & Mid(Value, i, 2)))
                    counter = counter + 1

                Next i


            End If

            Return bytes

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function fnCallDll(p_strBasePath As String, p_strDll As String, p_strClass As String, p_strRotine As String) As Object

        Dim Retorno As Object
        Dim objAssembly As Assembly
        Dim objType As Type

        Try
            Retorno = Nothing

            If p_strBasePath.Trim <> "" Then
                If p_strBasePath.EndsWith("\") = False Then
                    p_strBasePath = p_strBasePath & "\"
                End If
            End If

            If p_strDll.Trim <> "" Then
                If p_strDll.Trim.ToUpper.EndsWith(".DLL") = True Then
                    p_strDll = p_strDll.Trim.Substring(0, p_strDll.Length - 4)
                End If
            End If

            If File.Exists(p_strBasePath & p_strDll & ".dll") = False Then
                Throw New Exception("Dll not found: " & p_strBasePath & p_strDll)
            End If

            objAssembly = Assembly.LoadFrom(p_strBasePath & p_strDll.Trim & ".dll")
            objType = objAssembly.GetType(p_strDll.Trim & "." & p_strClass.Trim)

            If p_strRotine.Trim <> "" Then
                If objType.GetMember(p_strRotine).Length = 0 Then
                    Throw New Exception("Error loading assembly: " & p_strBasePath & p_strDll & vbNewLine & "Method not found: " & p_strRotine)
                End If
            End If

            Retorno = Activator.CreateInstance(objType)

            Return Retorno

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function fnCallDll(p_strBasePath As String, p_strDll As String, p_strClass As String, p_strRotine As String, p_objParametros() As Object) As Object

        Dim Retorno As Object
        Dim objAssembly As Assembly
        Dim objType As Type

        Try
            Retorno = Nothing

            If p_strBasePath.Trim <> "" Then
                If p_strBasePath.EndsWith("\") = False Then
                    p_strBasePath = p_strBasePath & "\"
                End If
            End If

            If p_strDll.Trim <> "" Then
                If p_strDll.Trim.ToUpper.EndsWith(".DLL") = True Then
                    p_strDll = p_strDll.Trim.Substring(0, p_strDll.Trim.Length - 4)
                End If
            End If

            If File.Exists(p_strBasePath & p_strDll & ".dll") = False Then
                Throw New Exception("Dll not found: " & p_strBasePath & p_strDll)
            End If

            objAssembly = Assembly.LoadFrom(p_strBasePath & p_strDll.Trim & ".dll")
            objType = objAssembly.GetType(p_strDll.Trim & "." & p_strClass.Trim)

            If p_strRotine.Trim <> "" Then
                If objType.GetMember(p_strRotine).Length = 0 Then
                    Throw New Exception("Error creating assembly: " & p_strBasePath & p_strDll & vbNewLine & "Method not found: " & p_strRotine)
                End If
            End If


            Retorno = Activator.CreateInstance(objType, p_objParametros)

            Return Retorno

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function fnCallDll(p_strBasePath As String, p_strDll As String, p_strClass As String, p_strRotine As String, p_intEmpresa As Integer, p_strUsuario As String) As Object

        Dim Retorno As Object
        Dim objAssembly As Assembly
        Dim objType As Type
        Dim objParametros() As Object

        Try
            Retorno = Nothing

            ReDim objParametros(1)
            objParametros(0) = p_intEmpresa
            objParametros(1) = p_strUsuario

            If p_strBasePath.Trim <> "" Then
                If p_strBasePath.EndsWith("\") = False Then
                    p_strBasePath = p_strBasePath & "\"
                End If
            End If

            If p_strDll.Trim <> "" Then
                If p_strDll.Trim.ToUpper.EndsWith(".DLL") = True Then
                    p_strDll = p_strDll.Trim.Replace(".DLL", "")
                End If
            End If

            If File.Exists(p_strBasePath & p_strDll & ".dll") = False Then
                Throw New Exception("Dll not found: " & p_strBasePath & p_strDll)
            End If

            objAssembly = Assembly.LoadFrom(p_strBasePath & p_strDll.Trim & ".dll")
            objType = objAssembly.GetType(p_strDll.Trim & "." & p_strClass.Trim)

            If p_strRotine.Trim <> "" Then
                If objType.GetMember(p_strRotine).Length = 0 Then
                    Throw New Exception("Error creating assembly: " & p_strBasePath & p_strDll & vbNewLine & "Method not found: " & p_strRotine)
                End If
            End If

            Retorno = Activator.CreateInstance(objType, objParametros)

            Return Retorno

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function fnCallDll(p_strBasePath As String, p_strDll As String, p_strClass As String) As Object

        Try

            Return fnCallDll(p_strBasePath, p_strDll, p_strClass, "")

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function fnCallDll(p_strBasePath As String, p_strDll As String, p_strClass As String, p_objParametros() As Object) As Object
        Try

            Return fnCallDll(p_strBasePath, p_strDll, p_strClass, "", p_objParametros)

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function fnCallDll(p_strBasePath As String, p_strDll As String, p_strClass As String, p_intEmpresa As Integer, p_strUsuario As String) As Object
        Try
            Return fnCallDll(p_strBasePath, p_strDll, p_strClass, "", p_intEmpresa, p_strUsuario)
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function fnComputerName() As String

        Dim strRetorno As String

        Try
            strRetorno = System.Net.Dns.GetHostName()
            Return strRetorno
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function fnUserName() As String

        Dim strRetorno As String

        Try
            strRetorno = Environment.UserName
            Return strRetorno
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function fnComputerIPv4() As String

        Dim strRetorno As String

        Try
            strRetorno = ""
            For Each Ip As IPAddress In System.Net.Dns.GetHostEntry(fnComputerName).AddressList
                If Ip.AddressFamily = Sockets.AddressFamily.InterNetwork Then
                    strRetorno = Ip.ToString
                    Exit For
                End If
            Next
            If strRetorno = "" Then
                strRetorno = System.Net.Dns.GetHostEntry(fnComputerName).AddressList(0).ToString()
            End If
            Return strRetorno
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function fnComputerIPv6() As String

        Dim strRetorno As String

        Try
            strRetorno = ""
            For Each Ip As IPAddress In System.Net.Dns.GetHostEntry(fnComputerName).AddressList
                If Ip.AddressFamily = Sockets.AddressFamily.InterNetworkV6 Then
                    strRetorno = Ip.ToString
                    Exit For
                End If
            Next
            If strRetorno = "" Then
                strRetorno = System.Net.Dns.GetHostEntry(fnComputerName).AddressList(0).ToString()
            End If
            Return strRetorno
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function fnComputerIPAll() As List(Of String)

        Dim strRetorno As List(Of String)
        Dim auxIp As List(Of IPAddress)

        Try
            strRetorno = New List(Of String)

            auxIp = System.Net.Dns.GetHostEntry(fnComputerName).AddressList.ToList
            For Each Ip As IPAddress In auxIp
                strRetorno.Add(Ip.ToString)
            Next

            Return strRetorno

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function fnSetImgOpacity(ByVal p_imgImage As Image, ByVal p_sngOpacity As Single) As Image

        Dim bmpPic As New Bitmap(p_imgImage.Width, p_imgImage.Height)
        Dim gfxPic As Graphics = Graphics.FromImage(bmpPic)
        Dim cmxPic As New ColorMatrix()
        Dim iaPic As New ImageAttributes()

        Try

            p_sngOpacity = p_sngOpacity / 100

            cmxPic.Matrix33 = p_sngOpacity

            iaPic.SetColorMatrix(cmxPic, ColorMatrixFlag.[Default], ColorAdjustType.Bitmap)
            gfxPic.DrawImage(p_imgImage, New Rectangle(0, 0, bmpPic.Width, bmpPic.Height), 0, 0, p_imgImage.Width, p_imgImage.Height, GraphicsUnit.Pixel, iaPic)

            gfxPic.Dispose()
            iaPic.Dispose()

            Return bmpPic

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function fnSetImgOpacity(ByVal p_imgImage As Bitmap, ByVal p_sngOpacity As Single) As Bitmap

        Dim bmpPic As New Bitmap(p_imgImage.Width, p_imgImage.Height)
        Dim gfxPic As Graphics = Graphics.FromImage(bmpPic)
        Dim cmxPic As New ColorMatrix()
        Dim iaPic As New ImageAttributes()

        Try

            p_sngOpacity = p_sngOpacity / 100

            cmxPic.Matrix33 = p_sngOpacity

            iaPic.SetColorMatrix(cmxPic, ColorMatrixFlag.[Default], ColorAdjustType.Bitmap)
            gfxPic.DrawImage(p_imgImage, New Rectangle(0, 0, p_imgImage.Width, bmpPic.Height), 0, 0, p_imgImage.Width, p_imgImage.Height, GraphicsUnit.Pixel, iaPic)

            gfxPic.Dispose()
            iaPic.Dispose()

            Return bmpPic

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function fnProprotionalHeight(p_dblOriginalWidth As Double, p_dblActualWidth As Double, p_dblOriginalHeight As Double) As Double

        Dim dblRetorno As String

        Try

            dblRetorno = p_dblOriginalHeight * (p_dblActualWidth / p_dblOriginalWidth)

            Return dblRetorno

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function fnMOD10(ByVal p_strvalue As String) As String

        Dim ret As Integer
        Dim Soma As Integer
        Dim Total As Integer
        Dim Contador As Integer
        Dim Peso As Integer
        Dim Digito As Integer

        Try
            If p_strvalue.Length = 0 Then Return -1


            Peso = 2
            Soma = 0
            Contador = p_strvalue.Length

            Do While (Contador >= 1)
                Soma = Integer.Parse(p_strvalue.Substring(Contador - 1, 1)) * Peso
                If (Soma <= 9) Then
                    Total += Soma
                Else
                    Total += Integer.Parse(Soma.ToString()(0)) + Integer.Parse(Soma.ToString()(1))
                End If

                Peso = IIf(Peso = 2, 1, 2)
                Contador -= 1
            Loop

            Digito = 10 - (Total Mod 10)
            If Digito > 9 Then Digito = 0
            ret = Digito

            Return ret.ToString
        Catch ex As Exception
            Throw
        End Try

    End Function

    Public Shared Function fnRandomPassword(p_intDigits As Integer, p_blnNumbers As Boolean, p_blnLowerCase As Boolean, p_blnUpperCase As Boolean, p_blnEspecialChars As Boolean) As String

        Dim Retorno As String
        Dim strCaracter As List(Of String)
        Dim objRandom As Random
        Dim intIndex As Integer
        Try

            Retorno = ""
            objRandom = New Random
            Randomize()

            For i = 0 To p_intDigits - 1

                strCaracter = New List(Of String)

                If p_blnNumbers = True Then
                    strCaracter.Add(objRandom.Next(0, 10))
                    strCaracter.Add(objRandom.Next(0, 10))
                    strCaracter.Add(objRandom.Next(0, 10))
                    strCaracter.Add(objRandom.Next(0, 10))
                    strCaracter.Add(objRandom.Next(0, 10))
                    strCaracter.Add(objRandom.Next(0, 10))
                    strCaracter.Add(objRandom.Next(0, 10))
                End If

                If p_blnUpperCase = True Then
                    strCaracter.Add(ChrW(objRandom.Next(65, 91)))
                    strCaracter.Add(ChrW(objRandom.Next(65, 91)))
                    strCaracter.Add(ChrW(objRandom.Next(65, 91)))
                    strCaracter.Add(ChrW(objRandom.Next(65, 91)))
                End If

                If p_blnLowerCase = True Then
                    strCaracter.Add(ChrW(objRandom.Next(97, 123)))
                    strCaracter.Add(ChrW(objRandom.Next(97, 123)))
                    strCaracter.Add(ChrW(objRandom.Next(97, 123)))
                    strCaracter.Add(ChrW(objRandom.Next(97, 123)))
                End If

                If p_blnEspecialChars = True Then
                    If objRandom.Next(0, 2) = 0 Then
                        strCaracter.Add("@")
                    Else
                        strCaracter.Add("!")
                    End If
                End If

                intIndex = objRandom.Next(0, strCaracter.Count)

                Retorno = Retorno & strCaracter(intIndex)

            Next

            Return Retorno

        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function fnRandomPassword(p_intDigits As Integer, p_blnNumbers As Boolean, p_blnLetras As Boolean, p_blnEspecialChars As Boolean) As String
        Try
            Return fnRandomPassword(p_intDigits, p_blnNumbers, p_blnLetras, p_blnLetras, p_blnEspecialChars)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function fnRandomPassword(p_intDigits As Integer, p_blnNumbers As Boolean, p_blnLetras As Boolean) As String
        Try
            Return fnRandomPassword(p_intDigits, p_blnNumbers, p_blnLetras, p_blnLetras, False)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function fnRandomPassword(p_intDigits As Integer) As String
        Try
            Return fnRandomPassword(p_intDigits, True, True, True, True)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function fnComputeCartesianCoordinate(p_dblangle As Double, p_dblRadius As Double) As Double()

        Dim dblAngleRad As Double
        Dim dblX As Double
        Dim dblY As Double

        Try

            dblAngleRad = (Math.PI / 180.0) * (p_dblangle - 90)
            dblX = p_dblRadius * Math.Cos(dblAngleRad)
            dblY = p_dblRadius * Math.Sin(dblAngleRad)

            Return {dblX, dblY}

        Catch ex As Exception
            Throw
        End Try

    End Function

    Public Shared Function fnComputeCircunference(p_dblRadius As Double) As Double
        Try
            Return 2 * Math.PI * p_dblRadius
        Catch ex As Exception
            Throw
        End Try

    End Function


#End Region

End Class

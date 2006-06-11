Option Explicit On 
Option Strict On

Imports System.IO
Imports System.Math
Imports System.Text

Public Class CCoder
    Implements IDisposable


#Region "Atributes"
    Private Const gc_intPragCuloare As Integer = 245
#End Region


#Region "New/Dispose"

    Public Sub New()

    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Try

        Catch exc As Exception
        End Try
    End Sub

#End Region


#Region "Private Functions"

    Private Overloads Function Code(ByVal bmpPatern As Bitmap, ByVal InputText As String) As Bitmap
        Dim bmpResult As Bitmap

        Dim i As Integer
        Dim j As Integer
        Dim k As Integer

        Dim c As Integer
        Dim cr As Integer
        Dim cg As Integer
        Dim cb As Integer
        Dim s As String
        Dim clr As System.Drawing.Color

        bmpResult = New Bitmap(bmpPatern.Width, bmpPatern.Height)

        For k = 0 To InputText.Length - 1
            c = Asc(InputText.Substring(k, 1))

            s = CStr(c)
            s = s.PadLeft(3, "0"c)

            cr = CInt(s.Substring(0, 1)) + 1
            cg = CInt(s.Substring(1, 1)) + 1
            cb = CInt(s.Substring(2, 1)) + 1
            SetR(i, j, bmpPatern, bmpResult, cr)
            SetG(i, j, bmpPatern, bmpResult, cg)
            SetB(i, j, bmpPatern, bmpResult, cb)
        Next

        Dim ii As Integer
        Dim jj As Integer
        For ii = i To bmpPatern.Height - 1
            For jj = 0 To bmpPatern.Width - 1
                If (ii = i AndAlso jj >= j) OrElse ii > i Then
                    clr = bmpPatern.GetPixel(jj, ii)
                    bmpResult.SetPixel(jj, ii, clr)
                End If
            Next
        Next

        Return bmpResult
    End Function

    Private Sub SetR(ByRef i As Integer, ByRef j As Integer, ByRef bmpPatern As Bitmap, ByRef bmpNew As Bitmap, ByRef cr As Integer)
        Dim bOk As Boolean = False
        Dim clr As System.Drawing.Color

        Try
            Do
                clr = bmpPatern.GetPixel(j, i)

                If bmpPatern.GetPixel(j, i).R < gc_intPragCuloare Then
                    bmpNew.SetPixel(j, i, Color.FromArgb(clr.A, (clr.R + cr) Mod 256, clr.G, clr.B))
                    bOk = True
                Else
                    bmpNew.SetPixel(j, i, clr)
                End If

                j += 1
                If j = bmpPatern.Width Then
                    j = 0
                    i += 1
                End If
            Loop While Not bOk
        Catch exc As Exception
            Throw exc
        End Try
    End Sub

    Private Sub SetG(ByRef i As Integer, ByRef j As Integer, ByRef bmpPatern As Bitmap, ByRef bmpNew As Bitmap, ByRef cg As Integer)
        Dim bOk As Boolean = False
        Dim clr As System.Drawing.Color

        Try
            Do
                clr = bmpPatern.GetPixel(j, i)

                If clr.G < gc_intPragCuloare Then
                    bmpNew.SetPixel(j, i, Color.FromArgb(clr.A, clr.R, (clr.G + cg) Mod 256, clr.B))
                    bOk = True
                Else
                    bmpNew.SetPixel(j, i, clr)
                End If

                j += 1
                If j = bmpPatern.Width Then
                    j = 0
                    i += 1
                End If
            Loop While Not bOk
        Catch exc As Exception
            Throw exc
        End Try
    End Sub

    Private Sub SetB(ByRef i As Integer, ByRef j As Integer, ByRef bmpPatern As Bitmap, ByRef bmpNew As Bitmap, ByRef cb As Integer)
        Dim bOk As Boolean = False
        Dim clr As System.Drawing.Color

        Try
            Do
                clr = bmpPatern.GetPixel(j, i)

                If clr.B < gc_intPragCuloare Then
                    bmpNew.SetPixel(j, i, Color.FromArgb(clr.A, clr.R, clr.G, (clr.B + cb) Mod 256))
                    bOk = True
                Else
                    bmpNew.SetPixel(j, i, clr)
                End If

                j += 1
                If j = bmpPatern.Width Then
                    j = 0
                    i += 1
                End If
            Loop While Not bOk
        Catch exc As Exception
            Throw exc
        End Try
    End Sub

    Private Overloads Function Decode(ByVal strPaternFile As String, ByVal strMessageFile As String) As String
        Dim bmpPatern As Bitmap
        Dim bmpMessage As Bitmap

        Dim c As Char
        Dim i As Integer
        Dim j As Integer

        Try
            Dim strResult As New StringBuilder()

            bmpPatern = New Bitmap(strPaternFile)
            bmpMessage = New Bitmap(strMessageFile)

            c = GetChar(bmpPatern, bmpMessage, i, j)
            While c <> Char.MinValue
                strResult.Append(c)
                c = GetChar(bmpPatern, bmpMessage, i, j)
            End While

            bmpPatern.Dispose()
            bmpMessage.Dispose()

            Return strResult.ToString
        Catch exc As Exception
            Throw exc
        End Try
    End Function

    Private Function GetChar(ByVal bmpPatern As Bitmap, ByVal bmpMessage As Bitmap, ByRef i As Integer, ByRef j As Integer) As Char
        Dim charResult As Char
        Dim strChar As String
        Dim intCrt As Integer

        Try
            If j >= bmpPatern.Width OrElse i >= bmpPatern.Height Then
                Return Char.MinValue
            End If

            intCrt = GetR(bmpPatern, bmpMessage, i, j)
            If intCrt < 0 Then
                Return Char.MinValue
            End If
            strChar = CStr(intCrt)

            intCrt = GetG(bmpPatern, bmpMessage, i, j)
            If intCrt < 0 Then
                Return Char.MinValue
            End If
            strChar &= CStr(intCrt)

            intCrt = GetB(bmpPatern, bmpMessage, i, j)
            If intCrt < 0 Then
                Return Char.MinValue
            End If
            strChar &= CStr(intCrt)

            charResult = Chr(CInt(strChar))

            Return charResult
        Catch exc As Exception
            Throw exc
        End Try
    End Function

    Private Function GetR(ByVal bmpPatern As Bitmap, ByVal bmpMessage As Bitmap, ByRef i As Integer, ByRef j As Integer) As Integer
        Dim intResult As Integer
        Dim intPatern As Integer
        Dim intMessage As Integer

        Try
            If j >= bmpPatern.Width OrElse i >= bmpPatern.Height Then
                Return -1
            End If
            intPatern = bmpPatern.GetPixel(j, i).R
            intMessage = bmpMessage.GetPixel(j, i).R

            While intPatern = intMessage
                j += 1
                If j = bmpPatern.Width Then
                    j = 0
                    i += 1
                End If
                If j >= bmpPatern.Width OrElse i >= bmpPatern.Height Then
                    Return -1
                End If

                intPatern = bmpPatern.GetPixel(j, i).R

                If intPatern < gc_intPragCuloare Then
                    intMessage = bmpMessage.GetPixel(j, i).R
                Else
                    intMessage = intPatern
                End If
            End While

            intResult = (((intMessage + 256) - intPatern) Mod 256) - 1
            If intResult < 0 Then
                intResult = 0
            End If

            j += 1
            If j = bmpPatern.Width Then
                j = 0
                i += 1
            End If

            Return intResult
        Catch exc As Exception
            Throw exc
        End Try
    End Function

    Private Function GetG(ByVal bmpPatern As Bitmap, ByVal bmpMessage As Bitmap, ByRef i As Integer, ByRef j As Integer) As Integer
        Dim intResult As Integer
        Dim intPatern As Integer
        Dim intMessage As Integer

        Try
            If j >= bmpPatern.Width OrElse i >= bmpPatern.Height Then
                Return -1
            End If
            intPatern = bmpPatern.GetPixel(j, i).G
            intMessage = bmpMessage.GetPixel(j, i).G

            While intPatern = intMessage
                j += 1
                If j = bmpPatern.Width Then
                    j = 0
                    i += 1
                End If
                If j >= bmpPatern.Width OrElse i >= bmpPatern.Height Then
                    Return -1
                End If

                intPatern = bmpPatern.GetPixel(j, i).G
                If intPatern < gc_intPragCuloare Then
                    intMessage = bmpMessage.GetPixel(j, i).G
                Else
                    intMessage = intPatern
                End If
            End While

            intResult = (((intMessage + 256) - intPatern) Mod 256) - 1
            If intResult < 0 Then
                intResult = 0
            End If

            j += 1
            If j = bmpPatern.Width Then
                j = 0
                i += 1
            End If

            Return intResult
        Catch exc As Exception
            Throw exc
        End Try
    End Function

    Private Function GetB(ByVal bmpPatern As Bitmap, ByVal bmpMessage As Bitmap, ByRef i As Integer, ByRef j As Integer) As Integer
        Dim intResult As Integer
        Dim intPatern As Integer
        Dim intMessage As Integer

        Try
            If j >= bmpPatern.Width OrElse i >= bmpPatern.Height Then
                Return -1
            End If
            intPatern = bmpPatern.GetPixel(j, i).B
            intMessage = bmpMessage.GetPixel(j, i).B

            While intPatern = intMessage
                j += 1
                If j = bmpPatern.Width Then
                    j = 0
                    i += 1
                End If
                If j >= bmpPatern.Width OrElse i >= bmpPatern.Height Then
                    Return -1
                End If

                intPatern = bmpPatern.GetPixel(j, i).B
                If intPatern < gc_intPragCuloare Then
                    intMessage = bmpMessage.GetPixel(j, i).B
                Else
                    intMessage = intPatern
                End If
            End While

            intResult = (((intMessage + 256) - intPatern) Mod 256) - 1
            If intResult < 0 Then
                intResult = 0
            End If

            j += 1
            If j = bmpPatern.Width Then
                j = 0
                i += 1
            End If

            Return intResult
        Catch exc As Exception
            Throw exc
        End Try
    End Function

    Private Function ReadFile(ByVal strFileText As String) As String
        Dim strResult As String
        Dim srReadText As StreamReader

        Try
            srReadText = New StreamReader(strFileText)
            strResult = srReadText.ReadToEnd

            Return strResult
        Catch exc As Exception
            Throw exc
        Finally

            Try
                srReadText.Close()
            Catch
            End Try

        End Try
    End Function

#End Region


#Region "Public Functions"

    Public Overloads Sub Code(ByVal strPaternFile As String, ByVal strMessage As String, ByVal strMessageFile As String)
        Dim bmpPatern As Bitmap
        Dim bmpMessage As Bitmap
        Try

            bmpPatern = New Bitmap(strPaternFile)
            bmpMessage = Code(bmpPatern, strMessage)   'ReadFile(strTextFile))
            bmpMessage.Save(strMessageFile)

        Catch exc As Exception
            Throw exc
        Finally

            Try
                bmpPatern.Dispose()
            Catch
            End Try

            Try
                bmpMessage.Dispose()
            Catch
            End Try

        End Try
    End Sub

    Public Overloads Sub Decode(ByVal strPaternFile As String, ByVal strMessageFile As String, ByRef strTextFile As String)
        Dim swWrite As StreamWriter
        Try

            'swWrite = New StreamWriter(strTextFile)
            'swWrite.Write(Decode(strPaternFile, strMessageFile))
            strTextFile = Decode(strPaternFile, strMessageFile)

        Catch exc As Exception
            Throw exc
        Finally

            Try
                swWrite.Close()
            Catch
            End Try

        End Try
    End Sub

#End Region


End Class

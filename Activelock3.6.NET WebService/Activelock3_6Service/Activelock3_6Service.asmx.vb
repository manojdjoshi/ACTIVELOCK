Option Strict On
Option Explicit On

Imports System.Text
Imports System.IO
Imports ActiveLock3_6NET
Imports System.Web.Services
Imports System.Data
Imports System.Data.OleDb
Imports System.Web.Services.Protocols
Imports System.ComponentModel

' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class Activelock3_6Service

    Inherits System.Web.Services.WebService

    Public ActiveLock3AlugenGlobals_definst As New AlugenGlobals
    Public ActiveLock3Globals_definst As New Globals
    Public GeneratorInstance As _IALUGenerator
    Public mKeyStoreType As IActiveLock.LicStoreType
    Public mProductsStoreType As IActiveLock.ProductsStoreType
    Public mProductsStoragePath As String
    Public ActiveLock As _IActiveLock

    Friend objConnection As OleDbConnection
    Friend objCommand As OleDbCommand
    Friend objDataReader As OleDbDataReader
    Friend fileMDB As String = AppPath() & "\" & "licenses.mdb"

    '<WebMethod()> _
    'Public Function HelloWorld() As String
    '    Return "Hello World"
    'End Function
    <WebMethod()> _
    Public Function GetTrial(ByVal pInstallCode As String, ByVal pSoftwareName As String, _
        ByVal pSoftwareVersion As String) As String
        Dim mLicenseKey As String = String.Empty
        Dim pRegisteredLevel As String = "Demo Only"
        'validate PackageSerial
        mLicenseKey = GetTrialForSerial(pInstallCode)
        If mLicenseKey.Length = 0 Then
            Dim mLicType As ProductLicense.ALLicType = ProductLicense.ALLicType.allicPeriodic
            Dim mDays As String = "15"
            Dim mExpireDate As String = GetExpirationDate(mLicType, mDays)
            Dim mRegDate As String = Date.UtcNow.ToString("yyyy/MM/dd")
            Dim maximumUsers As Short = 5
            Dim networkLicense As ProductLicense.LicFlags = ProductLicense.LicFlags.alfSingle
            mLicenseKey = GenerateKey(pInstallCode, pSoftwareName, pSoftwareVersion, mLicType _
              , pRegisteredLevel, mExpireDate, mRegDate, networkLicense, maximumUsers)
        Else
            'already have a trial
            'return the mLicenseKey
            mLicenseKey = "You have previously obtained a 15 day Trial License for this application"
        End If
        Return mLicenseKey
    End Function

    <WebMethod()> _
    Public Function GetLicense(ByVal pInstallCode As String, ByVal pSoftwareName As String, _
        ByVal pSoftwareVersion As String, ByVal pRegisteredLevel As String, ByVal networkLicense As ProductLicense.LicFlags, _
        ByVal maximumUsers As Short, ByVal mLicType As ProductLicense.ALLicType, ByVal mDays As String) As String
        Dim mLicenseKey As String = String.Empty
        'validate PackageSerial
        If GetTrialForSerial(pInstallCode).Length = 0 Then
            'Dim mLicType As ProductLicense.ALLicType = ProductLicense.ALLicType.allicPeriodic
            'Dim mDays As String = "30"
            Dim mExpireDate As String = GetExpirationDate(mLicType, mDays)
            Dim mRegDate As String = Date.UtcNow.ToString("yyyy/MM/dd")
            'Dim maximumUsers As Short = 5
            mLicenseKey = GenerateKey(pInstallCode, pSoftwareName, pSoftwareVersion, mLicType _
              , pRegisteredLevel, mExpireDate, mRegDate, networkLicense, maximumUsers)
        Else
            mLicenseKey = "You have already obtained a License for this installation code"
        End If
        Return mLicenseKey
    End Function

    Public Function GetTrialForSerial(ByVal pInstallCode As String) As String
        'get serial 
        Dim mResult As String = String.Empty
        Dim mPackageSerial As String = GetUserFromInstallCode(pInstallCode)
        Dim strSQL As String = "SELECT * FROM license WHERE InstCode='" & pInstallCode & "'"

        Try
            objConnection = New OleDbConnection
            objConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                        "Data Source= " & fileMDB & ";"
            objConnection.Open()

            objCommand = New OleDbCommand
            objCommand.Connection = objConnection
            objCommand.CommandTimeout = 60
            objCommand.CommandText = strSQL
            objCommand.CommandType = CommandType.Text
            objDataReader = objCommand.ExecuteReader()
            If objDataReader.HasRows Then
                objDataReader.Read()
                mResult = CType(objDataReader("LibCode"), String)
            End If
        Finally
            If Not objDataReader.IsClosed Then
                objDataReader.Close()
            End If
            If Not objConnection.State = ConnectionState.Closed Then
                objConnection.Close()
            End If
        End Try
        Return mResult
    End Function

    Public Function GenerateKey(ByVal pInstallCode As String, ByVal pSoftwareName As String, _
        ByVal pSoftwareVersion As String, ByVal pLicType As ProductLicense.ALLicType, _
        ByVal pRegisteredLevel As String, ByVal pExpireDate As String, ByVal pRegDate As String, _
        ByVal networkLicense As ProductLicense.LicFlags, ByVal maximumUsers As Short) As String

        Dim mLicenseKey As String = String.Empty
        Dim strSQL As String = String.Empty

        'generate license key
        Try
            InitActiveLock()
            'ActiveLock = ActiveLock3Globals_definst.NewInstance()
            ActiveLock.SoftwareName = pSoftwareName
            ActiveLock.SoftwareVersion = pSoftwareVersion

            'Dim Generator As New AlugenGlobals
            'GeneratorInstance = Generator.GeneratorInstance(IActiveLock.ProductsStoreType.alsMDBFile)
            'GeneratorInstance.StoragePath = AppPath() & "\licenses.mdb"

            'generate license object
            Dim Lic As ProductLicense
            Lic = ActiveLock3Globals_definst.CreateProductLicense(pSoftwareName, pSoftwareVersion, "", _
                      ProductLicense.LicFlags.alfSingle, pLicType, "", _
                      pRegisteredLevel, _
                      pExpireDate, , pRegDate, , maximumUsers)
            ' Pass it to IALUGenerator to generate the key
            mLicenseKey = GeneratorInstance.GenKey(Lic, pInstallCode, pRegisteredLevel)
            'split license key into 64byte chunks
            mLicenseKey = Make64ByteChunks(mLicenseKey & "aLck" & pInstallCode)

            'TODO - get locktype and username from InstallCode
            Dim pLockType As String = GetLockTypeFromInstallCode(pInstallCode)
            Dim pUserName As String = GetUserFromInstallCode(pInstallCode)
            Dim strLicType As String = String.Empty
            Select Case pLicType
                Case ProductLicense.ALLicType.allicPeriodic
                    strLicType = "Periodic"
                Case ProductLicense.ALLicType.allicPermanent
                    strLicType = "Permanent"
                Case ProductLicense.ALLicType.allicTimeLocked
                    strLicType = "Time Locked"
            End Select
            'save the license to licenses database
            Try
                strSQL = "INSERT INTO license ( Progname, progver, RegDate, ExpDate, LicType, LockType, RegLevel, InstCode, UserName, LibCode )" & _
                            " VALUES(@productName,@productversion,@registrationDate,@expiresAfter,@licenseType,@lockType,@registeredLevel,@installationCode,@userName,@licenseCode)"

                objConnection = New OleDbConnection
                objConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                            "Data Source= " & fileMDB & ";"

                objCommand = New OleDbCommand
                objCommand.CommandTimeout = 60
                objCommand.CommandText = strSQL
                objCommand.CommandType = CommandType.Text

                objCommand.Parameters.Clear()
                objCommand.Parameters.AddWithValue("@productName", pSoftwareName)
                objCommand.Parameters.AddWithValue("@productversion", pSoftwareVersion)
                objCommand.Parameters.AddWithValue("@registrationDate", pRegDate)
                objCommand.Parameters.AddWithValue("@expiresAfter", pExpireDate)
                objCommand.Parameters.AddWithValue("@licenseType", strLicType)
                objCommand.Parameters.AddWithValue("@lockType", pLockType)
                objCommand.Parameters.AddWithValue("@registeredLevel", pRegisteredLevel)
                objCommand.Parameters.AddWithValue("@installationCode", pInstallCode)
                objCommand.Parameters.AddWithValue("@userName", pUserName)
                objCommand.Parameters.AddWithValue("@licenseCode", mLicenseKey)

                objConnection.Open()
                objCommand.Connection = objConnection
                objCommand.ExecuteNonQuery()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(ex.Message)
            Finally
                If Not objConnection.State = ConnectionState.Closed Then
                    objConnection.Close()
                End If
            End Try
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message)
        Finally
            'final
        End Try
        Return mLicenseKey
    End Function

    Public Function AppPath() As String
        AppPath = System.IO.Path.GetDirectoryName(Server.MapPath("Default.aspx"))
    End Function

    Private Function GetExpirationDate(ByVal pLicType As ProductLicense.ALLicType, ByVal pDays As String) As String
        Dim mResult As String = String.Empty
        If pLicType = ProductLicense.ALLicType.allicTimeLocked Then
            If pDays.Length = 0 Then mResult = Date.UtcNow.AddDays(30).ToString("yyyy/MM/dd")
            mResult = CType(mResult, DateTime).ToString("yyyy/MM/dd")
        Else
            If pDays.Length = 0 Then pDays = "30"
            mResult = Date.UtcNow.AddDays(CType(pDays, Double)).ToString("yyyy/MM/dd")
        End If
        Return mResult
    End Function

    Private Shared Function Make64ByteChunks(ByRef strdata As String) As String
        ' Breaks a long string into chunks of 64-byte lines.
        Dim i As Integer
        Dim Count As Integer
        Dim strNew64Chunk As String
        Dim sResult As String = ""

        Count = strdata.Length
        For i = 0 To Count Step 64
            If i + 64 > Count Then
                strNew64Chunk = strdata.Substring(i)

            Else
                strNew64Chunk = strdata.Substring(i, 64)
            End If
            If sResult.Length > 0 Then
                sResult = sResult & strNew64Chunk
                'sResult = sResult & vbCrLf & strNew64Chunk
            Else
                sResult = sResult & strNew64Chunk
            End If
        Next

        Make64ByteChunks = sResult
    End Function

    Private Function GetUserFromInstallCode(ByVal strInstCode As String) As String
        ' Retrieves lock string and user info from the request string
        '
        Dim a() As String
        GetUserFromInstallCode = ""

        If strInstCode Is Nothing Or strInstCode = "" Then Exit Function
        strInstCode = ActiveLock3Globals_definst.Base64Decode(strInstCode)

        If Not strInstCode Is Nothing _
          AndAlso strInstCode.Substring(0, 1) = "+" Then
            strInstCode = strInstCode.Substring(1)
        End If

        Dim arrProdVer() As String
        arrProdVer = Split(strInstCode, "&&&") ' Extract the software name and version
        strInstCode = arrProdVer(0)

        a = Split(strInstCode, vbLf)
        GetUserFromInstallCode = a(a.Length - 1)

    End Function

    Private Function GetLockTypeFromInstallCode(ByVal strInstCode As String) As String
        ' Retrieves lock string and user info from the request string
        '
        Dim lockTypesString As String = String.Empty
        Dim a() As String
        Dim i As Integer
        Dim aString As String
        Dim usedLockNone As Boolean
        Dim noKey As String
        noKey = Chr(110) & Chr(111) & Chr(107) & Chr(101) & Chr(121)
        GetLockTypeFromInstallCode = ""

        If strInstCode Is Nothing Or strInstCode = "" Then Exit Function
        strInstCode = ActiveLock3Globals_definst.Base64Decode(strInstCode)

        If Not strInstCode Is Nothing _
          AndAlso strInstCode.Substring(0, 1) = "+" Then
            strInstCode = strInstCode.Substring(1)
            usedLockNone = True
        End If

        a = Split(strInstCode, vbLf)
        If usedLockNone = True Then
            For i = LBound(a) To UBound(a) - 1
                aString = a(i)
                If i = LBound(a) Then
                    If aString = "00 00 00 00 00 00" Or aString = "00-00-00-00-00-00" Or aString = "" Or aString = "Not Available" Then
                    Else
                        lockTypesString = lockTypesString & "MAC Address"
                    End If
                ElseIf i = LBound(a) + 1 Then
                    If lockTypesString.Length > 0 Then lockTypesString = lockTypesString & "+"
                    lockTypesString = lockTypesString & "Computer Name"
                ElseIf i = LBound(a) + 2 Then
                    If aString = "Not Available" Or aString = "0000-0000" Then
                    Else
                        If lockTypesString.Length > 0 Then lockTypesString = lockTypesString & "+"
                        lockTypesString = lockTypesString & "HDD Volume Serial"
                    End If
                ElseIf i = LBound(a) + 3 Then
                    If aString = "Not Available" Then
                    Else
                        If lockTypesString.Length > 0 Then lockTypesString = lockTypesString & "+"
                        lockTypesString = lockTypesString & "HDD Firmware Serial"
                    End If
                ElseIf i = LBound(a) + 4 Then
                    If lockTypesString.Length > 0 Then lockTypesString = lockTypesString & "+"
                    lockTypesString = lockTypesString & "Windows Serial"
                ElseIf i = LBound(a) + 5 Then
                    If aString = "Not Available" Then
                    Else
                        If lockTypesString.Length > 0 Then lockTypesString = lockTypesString & "+"
                        lockTypesString = lockTypesString & "BIOS Version"
                    End If
                ElseIf i = LBound(a) + 6 Then
                    If lockTypesString.Length > 0 Then lockTypesString = lockTypesString & "+"
                    lockTypesString = lockTypesString & "Motherboard Serial"
                ElseIf i = LBound(a) + 7 Then
                    If aString = "Not Available" Then
                    Else
                        If lockTypesString.Length > 0 Then lockTypesString = lockTypesString & "+"
                        lockTypesString = lockTypesString & "Local IP Address"
                    End If

                ElseIf i = LBound(a) + 8 Then
                    If aString = "Not Available" Then
                    Else
                        If lockTypesString.Length > 0 Then lockTypesString = lockTypesString & "+"
                        lockTypesString = lockTypesString & "External IP Address"
                    End If
                ElseIf i = LBound(a) + 9 Then
                    If aString = "Not Available" Then
                    Else
                        If lockTypesString.Length > 0 Then lockTypesString = lockTypesString & "+"
                        lockTypesString = lockTypesString & "Computer Fingerprint"
                    End If
                ElseIf i = LBound(a) + 10 Then
                    If aString = "Not Available" Then
                    Else
                        If lockTypesString.Length > 0 Then lockTypesString = lockTypesString & "+"
                        lockTypesString = lockTypesString & "Memory ID"
                    End If
                ElseIf i = LBound(a) + 11 Then
                    If aString = "Not Available" Then
                    Else
                        If lockTypesString.Length > 0 Then lockTypesString = lockTypesString & "+"
                        lockTypesString = lockTypesString & "CPU ID"
                    End If
                ElseIf i = LBound(a) + 12 Then
                    If aString = "Not Available" Then
                    Else
                        If lockTypesString.Length > 0 Then lockTypesString = lockTypesString & "+"
                        lockTypesString = lockTypesString & "Baseboard ID"
                    End If
                ElseIf i = LBound(a) + 13 Then
                    If aString = "Not Available" Then
                    Else
                        If lockTypesString.Length > 0 Then lockTypesString = lockTypesString & "+"
                        lockTypesString = lockTypesString & "Video ID"
                    End If
                End If
            Next i
        Else '"+" was not used, therefore one or more lockTypes were specified in the application
            For i = LBound(a) To UBound(a) - 1
                aString = a(i)
                If i = LBound(a) And aString <> noKey Then
                    If aString = "00 00 00 00 00 00" Or aString = "00-00-00-00-00-00" Or aString = "" Or aString = "Not Available" Then
                    Else
                        lockTypesString = lockTypesString & "MAC Address"
                    End If
                ElseIf i = (LBound(a) + 1) And aString <> noKey Then
                    If lockTypesString.Length > 0 Then lockTypesString = lockTypesString & "+"
                    lockTypesString = lockTypesString & "Computer Name"
                ElseIf i = (LBound(a) + 2) And aString <> noKey Then
                    If aString = "Not Available" Or aString = "0000-0000" Then
                    Else
                        If lockTypesString.Length > 0 Then lockTypesString = lockTypesString & "+"
                        lockTypesString = lockTypesString & "HDD Volume Serial"
                    End If
                ElseIf i = (LBound(a) + 3) And aString <> noKey Then
                    If aString = "Not Available" Then
                    Else
                        If lockTypesString.Length > 0 Then lockTypesString = lockTypesString & "+"
                        lockTypesString = lockTypesString & "HDD Firmware Serial"
                    End If
                ElseIf i = (LBound(a) + 4) And aString <> noKey Then
                    If lockTypesString.Length > 0 Then lockTypesString = lockTypesString & "+"
                    lockTypesString = lockTypesString & "Windows Serial"
                ElseIf i = (LBound(a) + 5) And aString <> noKey Then
                    If aString = "Not Available" Then
                    Else
                        If lockTypesString.Length > 0 Then lockTypesString = lockTypesString & "+"
                        lockTypesString = lockTypesString & "BIOS Version"
                    End If
                ElseIf i = (LBound(a) + 6) And aString <> noKey Then
                    If aString = "Not Available" Then
                    Else
                        If lockTypesString.Length > 0 Then lockTypesString = lockTypesString & "+"
                        lockTypesString = lockTypesString & "Motherboard Serial"
                    End If
                ElseIf i = (LBound(a) + 7) And aString <> noKey Then
                    If aString = "Not Available" Then
                    Else
                        If lockTypesString.Length > 0 Then lockTypesString = lockTypesString & "+"
                        lockTypesString = lockTypesString & "Local IP Address"
                    End If

                ElseIf i = (LBound(a) + 8) And aString <> noKey Then
                    If aString = "Not Available" Then
                    Else
                        If lockTypesString.Length > 0 Then lockTypesString = lockTypesString & "+"
                        lockTypesString = lockTypesString & "External IP Address"
                    End If
                ElseIf i = (LBound(a) + 9) And aString <> noKey Then
                    If aString = "Not Available" Then
                    Else
                        If lockTypesString.Length > 0 Then lockTypesString = lockTypesString & "+"
                        lockTypesString = lockTypesString & "Computer Fingerprint"
                    End If
                ElseIf i = (LBound(a) + 10) And aString <> noKey Then
                    If aString = "Not Available" Then
                    Else
                        If lockTypesString.Length > 0 Then lockTypesString = lockTypesString & "+"
                        lockTypesString = lockTypesString & "Memory ID"
                    End If
                ElseIf i = (LBound(a) + 11) And aString <> noKey Then
                    If aString = "Not Available" Then
                    Else
                        If lockTypesString.Length > 0 Then lockTypesString = lockTypesString & "+"
                        lockTypesString = lockTypesString & "CPU ID"
                    End If
                ElseIf i = (LBound(a) + 12) And aString <> noKey Then
                    If aString = "Not Available" Then
                    Else
                        If lockTypesString.Length > 0 Then lockTypesString = lockTypesString & "+"
                        lockTypesString = lockTypesString & "Baseboard ID"
                    End If
                ElseIf i = (LBound(a) + 13) And aString <> noKey Then
                    If aString = "Not Available" Then
                    Else
                        If lockTypesString.Length > 0 Then lockTypesString = lockTypesString & "+"
                        lockTypesString = lockTypesString & "Video ID"
                    End If
                End If
            Next i
        End If

        Return lockTypesString

    End Function
    Private Sub InitActiveLock()
        On Error GoTo InitForm_Error
        ActiveLock = ActiveLock3Globals_definst.NewInstance()
        ActiveLock.KeyStoreType = IActiveLock.LicStoreType.alsFile  'mKeyStoreType

        Dim MyAL As New Globals
        Dim MyGen As New AlugenGlobals

        'Use the following for ASP.NET applications
        'ActiveLock.Init(Application.StartupPath & "\bin")
        'Use the following for the VB.NET applications
        ActiveLock.Init(AppPath() & "\bin")

        ' Initialize Generator
        GeneratorInstance = MyGen.GeneratorInstance(IActiveLock.ProductsStoreType.alsMDBFile)
        'If File.Exists(mProductsStoragePath) = False Then
        '    Select Case mProductsStoreType
        '        Case IActiveLock.ProductsStoreType.alsINIFile
        '            mProductsStoragePath = AppPath() & "\licenses.ini"
        '        Case IActiveLock.ProductsStoreType.alsMDBFile
        '            mProductsStoragePath = AppPath() & "\licenses.mdb"
        '        Case IActiveLock.ProductsStoreType.alsXMLFile
        '            mProductsStoragePath = AppPath() & "\licenses.xml"
        '    End Select
        'End If
        GeneratorInstance.StoragePath = AppPath() & "\licenses.mdb"

        On Error GoTo 0
        Exit Sub

InitForm_Error:

        'MessageBox.Show("Error " & Err.Number & " (" & Err.Description & ") in procedure InitForm of Form frmMain", modALUGEN.ACTIVELOCKSTRING)
    End Sub

End Class
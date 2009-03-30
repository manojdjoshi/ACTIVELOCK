'*   ActiveLock
'*   Copyright 1998-2002 Nelson Ferraz
'*   Copyright 2003-2008 The ActiveLock Software Group (ASG)
'*   All material is the property of the contributing authors.
'*
'*   Redistribution and use in source and binary forms, with or without
'*   modification, are permitted provided that the following conditions are
'*   met:
'*
'*     [o] Redistributions of source code must retain the above copyright
'*         notice, this list of conditions and the following disclaimer.
'*
'*     [o] Redistributions in binary form must reproduce the above
'*         copyright notice, this list of conditions and the following
'*         disclaimer in the documentation and/or other materials provided
'*         with the distribution.
'*
'*   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
'*   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
'*   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
'*   A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
'*   OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
'*   SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
'*   LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
'*   DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
'*   THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
'*   (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
'*   OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'*
'*
Option Strict On
Option Explicit On 

Imports System.Text
Imports System.IO
Imports ActiveLock3_6NET
Imports MagicAjax
Imports System.Security.Cryptography
Imports System.Text.RegularExpressions
Imports System.Reflection

Namespace ASPNETAlugen3

    Public Class ASPNETAlugen3
        Inherits System.Web.UI.Page

        Public ActiveLock3AlugenGlobals_definst As New AlugenGlobals
        Public ActiveLock3Globals_definst As New Globals
        Friend GeneratorInstance As _IALUGenerator
        Friend ActiveLock As _IActiveLock
        Friend dt As New DataTable
        Private msgPleaseWait As String = "(generating codes - please wait)"

        Public mKeyStoreType As IActiveLock.LicStoreType
        Public mProductsStoreType As IActiveLock.ProductsStoreType
        Public mProductsStoragePath As String
        Public blnIsFirstLaunch As Boolean

        ' Hardware keys from the Installation Code
        Public MACaddress, ComputerName As String
        Public VolumeSerial, FirmwareSerial As String
        Public WindowsSerial, BIOSserial As String
        Public MotherboardSerial, IPaddress As String
        Public ExternalIP, Fingerprint As String
        Public Memory, CPUID As String
        Public BaseboardID, VideoID As String
        Public systemEvent As Boolean

        Public PROJECT_INI_FILENAME As String
        Public cboProducts_Array() As String


#Region " Web Form Designer Generated Code "

        'This call is required by the Web Form Designer.
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub

        'Protected WithEvents ltlAlert As System.Web.UI.WebControls.Literal
        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
            'CODEGEN: This method call is required by the Web Form Designer
            'Do not modify it using the code editor.
            InitializeComponent()
        End Sub

#End Region

        Public Shared ReadOnly ExpBase64 As Regex = New Regex("^[a-zA-Z0-9+/=]{4,}$", RegexOptions.Compiled)

        Public Sub AppendLockString(ByRef strLock As String, ByVal newSubString As String)
            '===============================================================================
            ' Name: Sub AppendLockString
            ' Input:
            '   ByRef strLock As String - The lock string to be appended to, returns as an output
            '   ByVal newSubString As String - The string to be appended to the lock string if strLock is empty string
            ' Output:
            '   Appended lock string and installation code
            ' Purpose: Appends the lock string to the given installation code
            ' Remarks: None
            '===============================================================================

            If strLock = "" Then
                strLock = newSubString
            Else
                strLock = strLock & vbLf & newSubString
            End If
        End Sub
        Public Function ReconstructedInstallationCode() As String
            Dim strLock As String = Nothing
            Dim strReq As String
            Dim noKey As String
            noKey = Chr(110) & Chr(111) & Chr(107) & Chr(101) & Chr(121)

            If Me.chkLockMACaddress.Checked = True Then
                AppendLockString(strLock, MACaddress)
            Else
                AppendLockString(strLock, noKey)
            End If
            If Me.chkLockComputer.Checked = True Then
                AppendLockString(strLock, ComputerName)
            Else
                AppendLockString(strLock, noKey)
            End If
            If Me.chkLockHD.Checked = True Then
                AppendLockString(strLock, VolumeSerial)
            Else
                AppendLockString(strLock, noKey)
            End If
            If Me.chkLockHDfirmware.Checked = True Then
                AppendLockString(strLock, FirmwareSerial)
            Else
                AppendLockString(strLock, noKey)
            End If
            If Me.chkLockWindows.Checked = True Then
                AppendLockString(strLock, WindowsSerial)
            Else
                AppendLockString(strLock, noKey)
            End If
            If Me.chkLockBIOS.Checked = True Then
                AppendLockString(strLock, BIOSserial)
            Else
                AppendLockString(strLock, noKey)
            End If
            If Me.chkLockMotherboard.Checked = True Then
                AppendLockString(strLock, MotherboardSerial)
            Else
                AppendLockString(strLock, noKey)
            End If
            If Me.chkLockIP.Checked = True Then
                AppendLockString(strLock, IPaddress)
            Else
                AppendLockString(strLock, noKey)
            End If
            If Me.chkLockExternalIP.Checked = True Then
                AppendLockString(strLock, ExternalIP)
            Else
                AppendLockString(strLock, noKey)
            End If
            If Me.chkLockFingerprint.Checked = True Then
                AppendLockString(strLock, Fingerprint)
            Else
                AppendLockString(strLock, noKey)
            End If
            If Me.chkLockMemory.Checked = True Then
                AppendLockString(strLock, Memory)
            Else
                AppendLockString(strLock, noKey)
            End If
            If Me.chkLockCPUID.Checked = True Then
                AppendLockString(strLock, CPUID)
            Else
                AppendLockString(strLock, noKey)
            End If
            If Me.chkLockBaseboardID.Checked = True Then
                AppendLockString(strLock, BaseboardID)
            Else
                AppendLockString(strLock, noKey)
            End If
            If Me.chkLockVideoID.Checked = True Then
                AppendLockString(strLock, VideoID)
            Else
                AppendLockString(strLock, noKey)
            End If

            If Not strLock Is Nothing _
              AndAlso strLock.Substring(0, 1) = vbLf Then
                strLock = strLock.Substring(2)
            End If

            Dim Index, i As Integer
            Dim strInstCode As String
            strInstCode = ActiveLock3Globals_definst.Base64Decode(txtInstallCode.Text)

            If strInstCode = "" Then Return Nothing

            If Not strInstCode Is Nothing _
              AndAlso strInstCode.Substring(0, 1) = "+" Then
                strInstCode = strInstCode.Substring(2)
            End If
            Dim arrProdVer() As String
            arrProdVer = Split(strInstCode, "&&&") ' Extract the software name and version
            strInstCode = arrProdVer(0)
            Index = 0
            i = 1
            ' Get to the last vbLf, which denotes the ending of the lock code and beginning of user name.
            Do While i > 0
                i = strInstCode.IndexOf(vbLf, Index) 'InStr(Index + 1, strInstCode, vbLf)
                If i > 0 Then Index = i + 1
            Loop
            ' user name starts from Index+1 to the end
            Dim user As String
            user = strInstCode.Substring(Index)

            ' combine with user name
            strReq = strLock & vbLf & user

            ' combine with app name and version
            strReq = strReq & "&&&" & cboProduct.Items(cboProduct.SelectedIndex).Text

            ' base-64 encode the request
            Dim strReq2 As String
            strReq2 = ActiveLock3Globals_definst.Base64Encode("+" & strReq)
            ReconstructedInstallationCode = strReq2

        End Function

        Public Function GetUserSoftwareNameandVersionfromInstallCode(ByVal strInstCode As String) As String
            On Error GoTo noInfo
            If strInstCode = "" Then Return String.Empty
            strInstCode = ActiveLock3Globals_definst.Base64Decode(txtInstallCode.Text)
            Dim arrProdVer() As String
            arrProdVer = Split(strInstCode, "&&&")
            GetUserSoftwareNameandVersionfromInstallCode = Trim$(arrProdVer(1))
noInfo:
        End Function

        Private Sub UpdateKeyGenButtonStatus()
            If txtInstallCode.Text = "" Then
                cmdGenerateCode.Enabled = False
            Else
                If cboProduct.SelectedIndex >= 0 Then
                    cmdGenerateCode.Enabled = True
                End If
            End If
        End Sub

        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
            'Put user code to initialize the page here

            ' Populate Product List on Product Code Generator tab
            ' and Key Gen tab with product info from licenses.ini
            Dim i As Integer
            Dim arrProdInfos() As ProductInfo
            Dim MyGen As New AlugenGlobals
            GeneratorInstance = MyGen.GeneratorInstance(IActiveLock.ProductsStoreType.alsINIFile)
            GeneratorInstance.StoragePath = AppPath() & "\licenses.ini"

            arrProdInfos = GeneratorInstance.RetrieveProducts()
            If arrProdInfos.Length = 0 Then Exit Sub

            If Not IsPostBack Then
                blnIsFirstLaunch = True
                'load form settings
                LoadFormSetting()
            End If

            ' Initialize AL
            'InitActiveLock()

            'Populate the DataGrid control
            Dim dc As New DataColumn("Name", System.Type.GetType("System.String"))
            If Not dt.Columns.Contains("Name") Then dt.Columns.Add(dc)
            dc = New DataColumn("Version", System.Type.GetType("System.String"))
            If Not dt.Columns.Contains("Version") Then dt.Columns.Add(dc)
            dc = New DataColumn("VCode", System.Type.GetType("System.String"))
            If Not dt.Columns.Contains("VCode") Then dt.Columns.Add(dc)
            dc = New DataColumn("GCode", System.Type.GetType("System.String"))
            If Not dt.Columns.Contains("GCode") Then dt.Columns.Add(dc)

            If Not MagicAjax.MagicAjaxContext.Current.IsAjaxCall Then
                cboProduct.Items.Clear()
            End If

            For i = 0 To arrProdInfos.Length - 1
                PopulateUI(arrProdInfos(i))
            Next

            grdProducts.AutoGenerateColumns = False

            Dim dgc As New LimitColumn ' BoundColumn
            dgc.HeaderText = "Product"
            dgc.DataField = "Name"
            dgc.ItemStyle.Width = Unit.Percentage(35)
            dgc.SortExpression = "Name"
            grdProducts.Columns.Add(dgc)

            dgc = New LimitColumn 'BoundColumn
            dgc.HeaderText = "Version"
            dgc.DataField = "Version"
            dgc.ItemStyle.Wrap = True
            dgc.ItemStyle.Width = Unit.Percentage(15)
            dgc.SortExpression = "Version"
            grdProducts.Columns.Add(dgc)

            dgc = New LimitColumn 'BoundColumn
            dgc.HeaderText = "VCode"
            dgc.DataField = "VCode"
            dgc.ItemStyle.Wrap = True
            dgc.ItemStyle.Width = Unit.Percentage(25)
            dgc.SortExpression = "VCode"
            dgc.CharacterLimit = 120
            grdProducts.Columns.Add(dgc)

            dgc = New LimitColumn 'BoundColumn
            dgc.HeaderText = "GCode"
            dgc.DataField = "GCode"
            dgc.ItemStyle.Wrap = True
            dgc.ItemStyle.Width = Unit.Percentage(25)
            dgc.SortExpression = "GCode"
            dgc.CharacterLimit = 120
            grdProducts.Columns.Add(dgc)

            ' Create a DataView from the DataTable.
            Dim dv As DataView = New DataView(dt)
            If sortExpression.Value.Length > 0 Then
                dv.Sort = sortExpression.Value.ToString() & " " & sortOrder.Value.ToString()
            Else
                dv.Sort = String.Empty
            End If

            grdProducts.DataSource = dv
            grdProducts.DataBind()

            If Not IsPostBack Then

                'initialize RegisteredLevels
                strRegisteredLevelDBName = AddBackSlash(AppPath) & "RegisteredLevelDB.dat"
                If Not MagicAjax.MagicAjaxContext.Current.IsAjaxCall Then
                    loadRegisteredLevels()
                End If

                grdProducts.SelectedIndex = dt.Rows.Count - 1
                txtProductName.Text = arrProdInfos(i - 1).Name
                txtProductVersion.Text = arrProdInfos(i - 1).Version
                txtVCode.Text = arrProdInfos(i - 1).VCode
                txtGCode.Text = arrProdInfos(i - 1).GCode

                cmdGenerateCode.Attributes.Add("OnClick", "document.getElementById('" & txtVCode.ClientID & "').value='" & msgPleaseWait & "'; document.getElementById('" & txtGCode.ClientID & "').value='" & msgPleaseWait & "'")
                cmdCopyVCode.Attributes.Add("OnClick", "CopyToClipboard('" & txtVCode.ClientID & "');")
                cmdCopyGCode.Attributes.Add("OnClick", "CopyToClipboard('" & txtGCode.ClientID & "');")
                cmdCopyLicenseKey.Attributes.Add("OnClick", "CopyToClipboard('" & txtLicenseKey.ClientID & "');")
                cmdPasteInstallCode.Attributes.Add("OnClick", "PasteFromClipboard('" & txtInstallCode.ClientID & "');")
                cmdPaste.Attributes.Add("OnClick", "PasteFromClipboard('" & txtInstallCode.ClientID & "');")

                grdProducts.SelectedItem.Attributes.Item("onmouseover") = "this.style.cursor='hand'"
                grdProducts.SelectedItem.Attributes.Remove("onmouseout")

                txtProductName.Attributes.Add("onKeyPress", "return letternumber(event)")
                txtProductVersion.Attributes.Add("onKeyPress", "return letternumber(event)")

                'chkLockBIOS.Attributes.Add("onclick", "fnchkLockBIOS_CheckedChanged();")

                pnlProducts.Visible = True
                pnlLicenses.Visible = False

                'Assume that the application LockType is not LOckNone only
                txtUserName.Enabled = False
                txtUserName.ReadOnly = True
                txtUserName.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
                cboProduct.SelectedIndex = Convert.ToInt32(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "cboProducts", CStr(0)))
                txtDays.Text = ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "txtDays", CStr(365))
            End If

            'If CheckDuplicate(txtProductName.Text, txtProductVersion.Text) Then
            '    cmdGenerateCode.Enabled = False
            '    cmdAddProduct.Enabled = False
            '    cmdRemoveProduct.Enabled = True
            'Else
            '    cmdGenerateCode.Enabled = True
            '    cmdAddProduct.Enabled = True
            '    cmdRemoveProduct.Enabled = False
            'End If
            blnIsFirstLaunch = False

        End Sub
        Public Sub LoadFormSetting()
            'Read the program INI file to retrieve control settings
            On Error GoTo LoadFormSetting_Error

            If Not blnIsFirstLaunch Then Exit Sub

            On Error Resume Next
            'Read the program INI file to retrieve control settings
            Dim CallingAssemblyName As String = Assembly.GetCallingAssembly.GetName().Name
            PROJECT_INI_FILENAME = AppPath() & "\ASPNETAlugen3.6.ini"

            'cboProducts.SelectedIndex = Convert.ToInt32(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "cboProducts", CStr(0)))
            cboLicenseType.SelectedIndex = Convert.ToInt32(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "cboLicenseType", CStr(1)))
            cboRegLevel.SelectedIndex = Convert.ToInt32(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "cboRegLevel", CStr(0)))
            If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkUseItemData", CStr(0)) = "False" Then
                chkUseItemData.Checked = False
            Else
                chkUseItemData.Checked = True
            End If
            If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockBIOS", CStr(0)) = "False" Then
                chkLockBIOS.Checked = False
            Else
                chkLockBIOS.Checked = True
            End If
            If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockComputer", CStr(0)) = "False" Then
                chkLockComputer.Checked = False
            Else
                chkLockComputer.Checked = True
            End If
            If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockHD", CStr(0)) = "False" Then
                chkLockHD.Checked = False
            Else
                chkLockHD.Checked = True
            End If
            If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockHDfirmware", CStr(0)) = "False" Then
                chkLockHDfirmware.Checked = False
            Else
                chkLockHDfirmware.Checked = True
            End If
            If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockIP", CStr(0)) = "False" Then
                chkLockIP.Checked = False
            Else
                chkLockIP.Checked = True
            End If
            If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockMACaddress", CStr(0)) = "False" Then
                chkLockMACaddress.Checked = False
            Else
                chkLockMACaddress.Checked = True
            End If
            If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockMotherboard", CStr(0)) = "False" Then
                chkLockMotherboard.Checked = False
            Else
                chkLockMotherboard.Checked = True
            End If
            If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockWindows", CStr(0)) = "False" Then
                chkLockWindows.Checked = False
            Else
                chkLockWindows.Checked = True
            End If
            If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockExternalIP", CStr(0)) = "False" Then
                chkLockExternalIP.Checked = False
            Else
                chkLockExternalIP.Checked = True
            End If
            If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockFingerprint", CStr(0)) = "False" Then
                chkLockFingerprint.Checked = False
            Else
                chkLockFingerprint.Checked = True
            End If
            If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockMemory", CStr(0)) = "False" Then
                chkLockMemory.Checked = False
            Else
                chkLockMemory.Checked = True
            End If
            If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockCPUID", CStr(0)) = "False" Then
                chkLockCPUID.Checked = False
            Else
                chkLockCPUID.Checked = True
            End If
            If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockBaseboardID", CStr(0)) = "False" Then
                chkLockBaseboardID.Checked = False
            Else
                chkLockBaseboardID.Checked = True
            End If
            If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockVideoID", CStr(0)) = "False" Then
                chkLockVideoID.Checked = False
            Else
                chkLockVideoID.Checked = True
            End If

            If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkNetworkedLicense", CStr(0)) = "False" Then
                chkNetworkedLicense.Checked = False
            Else
                chkNetworkedLicense.Checked = True
            End If

            txtMaxCount.Text = ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "txtMaxCount", CStr(5))
            'txtDays.Text = ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "txtDays", CStr(365))

            optStrength0.Checked = CBool(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "optStrength0", "False"))
            optStrength1.Checked = CBool(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "optStrength1", "True"))
            optStrength2.Checked = CBool(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "optStrength2", "False"))
            optStrength3.Checked = CBool(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "optStrength3", "False"))
            optStrength4.Checked = CBool(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "optStrength4", "False"))
            optStrength5.Checked = CBool(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "optStrength5", "False"))

            mKeyStoreType = CType(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "KeyStoreType", CStr(1)), IActiveLock.LicStoreType)
            mProductsStoreType = CType(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "ProductsStoreType", CStr(0)), IActiveLock.ProductsStoreType)
            mProductsStoragePath = ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "ProductsStoragePath", AppPath() & "\licenses.ini")
            If Not File.Exists(mProductsStoragePath) And mProductsStoreType = IActiveLock.ProductsStoreType.alsMDBFile Then
                mProductsStoreType = IActiveLock.ProductsStoreType.alsINIFile
                mProductsStoragePath = AppPath() & "\licenses.ini"
            End If

            blnIsFirstLaunch = False

            On Error GoTo 0
            Exit Sub

LoadFormSetting_Error:

            SayAjax("Error " & Err.Number & " (" & Err.Description & ") in procedure LoadFormSetting of Form frmMain")

        End Sub
        Public Sub SaveFormSettings()
            'save form settings
            On Error GoTo SaveFormSettings_Error
            Dim mnReturnValue As Long

            PROJECT_INI_FILENAME = AppPath() & "\ASPNETAlugen3.6.ini"

            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "cboProducts", CStr(cboProduct.SelectedIndex))
            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "cboLicType", CStr(cboLicenseType.SelectedIndex))
            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "cboRegisteredLevel", CStr(cboRegLevel.SelectedIndex))
            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkItemData", chkUseItemData.Checked.ToString)
            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkNetworkedLicense", chkNetworkedLicense.Checked.ToString)

            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "KeyStoreType", CStr(mKeyStoreType))
            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "ProductsStoreType", CStr(mProductsStoreType))
            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "ProductsStoragePath", CStr(mProductsStoragePath))

            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockBIOS", chkLockBIOS.Checked.ToString)
            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockComputer", chkLockComputer.Checked.ToString)
            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockHD", chkLockHD.Checked.ToString)
            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockHDfirmware", chkLockHDfirmware.Checked.ToString)
            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockIP", chkLockIP.Checked.ToString)
            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockMACaddress", chkLockMACaddress.Checked.ToString)
            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockMotherboard", chkLockMotherboard.Checked.ToString)
            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockWindows", chkLockWindows.Checked.ToString)
            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockExternalIP", chkLockExternalIP.Checked.ToString)
            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockFingerpoint", chkLockFingerprint.Checked.ToString)
            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockMemory", chkLockMemory.Checked.ToString)
            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockCPUID", chkLockCPUID.Checked.ToString)
            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockBaseboardID", chkLockBaseboardID.Checked.ToString)
            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockVideoID", chkLockVideoID.Checked.ToString)

            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "txtMaxCount", txtMaxCount.Text)
            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "txtDays", txtDays.Text)

            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "optstrength0", CStr(optStrength0.Checked))
            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "optstrength1", CStr(optStrength1.Checked))
            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "optstrength2", CStr(optStrength2.Checked))
            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "optstrength3", CStr(optStrength3.Checked))
            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "optstrength4", CStr(optStrength4.Checked))
            mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "optstrength5", CStr(optStrength5.Checked))

            On Error GoTo 0
            Exit Sub

SaveFormSettings_Error:

            SayAjax("Error " & Err.Number & " (" & Err.Description & ") in procedure SaveFormSettings of Page Unload")
        End Sub

        Private Sub cmdProducts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdProducts.Click
            ResetCalendarControls()
            pnlProducts.Visible = True
            pnlLicenses.Visible = False
        End Sub

        Private Sub cmdLicenses_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLicenses.Click
            ResetCalendarControls()
            pnlProducts.Visible = False
            pnlLicenses.Visible = True
        End Sub

        Public Function AppPath() As String
            AppPath = System.IO.Path.GetDirectoryName(Server.MapPath("Default.aspx"))
        End Function

        Function CreateDataSource(ByVal Name As String, ByVal Ver As String, ByVal Code1 As String, ByVal Code2 As String) As ICollection

            Dim dr As DataRow = dt.NewRow()
            CreateDataSource = Nothing
            'Create new row and add data to it
            dr(0) = Name
            dr(1) = Ver
            dr(2) = Code1
            dr(3) = Code2

            'add the row to the datatable
            dt.Rows.Add(dr)

        End Function

        Private Sub grdProducts_ItemCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles grdProducts.ItemCreated
            If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Or e.Item.ItemType = ListItemType.SelectedItem Then
                If e.Item.ItemIndex <> grdProducts.SelectedIndex Then
                    e.Item.Attributes.Add("onmouseover", "this.style.backgroundColor='#F0F0F0';this.style.cursor='pointer'")
                    e.Item.Attributes.Add("onmouseout", "this.style.backgroundColor='#FFFFFF';this.style.color='black'")
                    e.Item.Attributes.Add("onclick", "javascript:__doPostBack('" & "grdProducts$" & "_ctl" & (e.Item.ItemIndex + 2) & "$_ctl0','')")
                Else
                    e.Item.Attributes.Item("onmouseover") = "this.style.cursor='hand'"
                    e.Item.Attributes.Remove("onmouseout")
                End If
            End If
        End Sub

        Private Sub grdProducts_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdProducts.SelectedIndexChanged

            grdProducts.SelectedItem.Attributes.Item("onmouseover") = "this.style.cursor='hand'"
            grdProducts.SelectedItem.Attributes.Remove("onmouseout")

            FillDataFromSelectedRow()

        End Sub

        Private Sub FillDataFromSelectedRow()
            txtProductName.Text = grdProducts.SelectedItem.Cells(1).Text
            txtProductVersion.Text = grdProducts.SelectedItem.Cells(2).Text

            Dim arrProdInfos() As ProductInfo
            Dim MyGen As New AlugenGlobals
            GeneratorInstance = MyGen.GeneratorInstance(IActiveLock.ProductsStoreType.alsINIFile)
            GeneratorInstance.StoragePath = AppPath() & "\licenses.ini"
            arrProdInfos = GeneratorInstance.RetrieveProducts()

            For i As Integer = 0 To arrProdInfos.Length - 1
                If arrProdInfos(i).Name = grdProducts.SelectedItem.Cells(1).Text _
                AndAlso arrProdInfos(i).Version = grdProducts.SelectedItem.Cells(2).Text Then
                    txtVCode.Text = arrProdInfos(i).VCode
                    txtGCode.Text = arrProdInfos(i).GCode
                    Exit For
                End If
            Next
        End Sub

        Private Sub Say(ByVal Message As String)
            'work only in postback scenario
            ' Format string properly
            Message = Message.Replace("'", "\'")
            Message = Message.Replace(Convert.ToChar(10), "\n")
            Message = Message.Replace(Convert.ToChar(13), "")
            ' Display as JavaScript alert
            ltlAlert.Text = "alert('" & Message & "')"
        End Sub

        Private Sub SayAjax(ByVal Message As String)
            'for ajax purpose
            ' Format string properly
            Message = Message.Replace("'", "\'")
            Message = Message.Replace(Convert.ToChar(10), "\n")
            Message = Message.Replace(Convert.ToChar(13), "")

            plhSay.Controls.Clear()
            Dim myLiteral As New WebControls.Literal
            myLiteral.Text = Message
            Dim myDIV As New HtmlGenericControl("div")
            myDIV.Attributes.Add("style", "background-color: #F0F0F0; border:solid 1px red; padding:3px; position: absolute; top: 10; right: 10")
            myDIV.Controls.Add(myLiteral)
            plhSay.Controls.Add(myDIV)

        End Sub

        Private Sub cmdAddProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddProduct.Click
            'This button add the new product into the list and into the INI file

            If txtVCode.Text.Trim.Length = 0 Then
                Say("New product VCode is missing.")
                Exit Sub
            ElseIf txtGCode.Text.Trim.Length = 0 Then
                Say("New product GCode is missing.")
                Exit Sub
            End If
            If CheckDuplicate(txtProductName.Text, txtProductVersion.Text) Then
                Say("Selected Product Name and Version already exist in the product list.")
                Exit Sub
            End If

            AddRow(txtProductName.Text, txtProductVersion.Text, txtVCode.Text, txtGCode.Text, True)

            cmdRemoveProduct.Enabled = True
            cmdAddProduct.Enabled = False ' disallow repeated clicking of Add button
            cmdValidateCode.Enabled = True

            SayAjax(String.Format("Product {0} ({1}) added successfuly.", txtProductName.Text, txtProductVersion.Text))
        End Sub

        Private Sub cmdRemoveProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemoveProduct.Click
            Dim dr As DataRow = dt.Rows(grdProducts.SelectedItem.ItemIndex)
            dt.Rows.Remove(dr)
            Dim strName, strVer As String
            strName = grdProducts.SelectedItem.Cells(1).Text
            strVer = grdProducts.SelectedItem.Cells(2).Text

            GeneratorInstance.DeleteProduct(strName, strVer)
            grdProducts.DataSource = dt
            grdProducts.DataBind()

            grdProducts.SelectedIndex = dt.Rows.Count - 1
            FillDataFromSelectedRow()

            SayAjax(String.Format("Product {0} ({1}) removed successfuly.", strName, strVer))
        End Sub

        Private Sub grdProducts_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles grdProducts.SortCommand
            ' Retrieve the data source from session state.
            Dim mydt As DataTable = dt

            ' Create a DataView from the DataTable.
            Dim dv As DataView = New DataView(mydt)

            ' The DataView provides an easy way to sort. Simply set the
            ' Sort property with the name of the field to sort by.
            If sortExpression.Value <> e.SortExpression Then
                sortExpression.Value = e.SortExpression
                sortOrder.Value = ""
            End If
            If sortOrder.Value = "ASC" Then
                sortOrder.Value = "DESC"
            Else
                sortOrder.Value = "ASC"
            End If

            dv.Sort = sortExpression.Value & " " & sortOrder.Value

            ' Re-bind the data source and specify that it should be sorted
            ' by the field specified in the SortExpression property.
            grdProducts.DataSource = dv
            grdProducts.DataBind()

        End Sub

        Private Sub cmdGenerateCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGenerateCode.Click

            If txtProductName.Text = "" Then
                SayAjax("Product Name is empty.")
                Exit Sub
            ElseIf txtProductVersion.Text = "" Then
                SayAjax("Product Version is empty.")
                Exit Sub
            ElseIf CheckDuplicate(txtProductName.Text, txtProductVersion.Text) Then
                SayAjax("Selected Product Name and Version already exist in the products list.")
                txtVCode.Text = ""
                txtGCode.Text = ""
                Exit Sub
            End If
            If txtProductName.Text.Contains("-") Then
                MsgBox("Dash '-' is not allowed in Product Name.", vbInformation)
                Exit Sub
            End If
            If txtProductVersion.Text.Contains("-") Then
                MsgBox("Dash '-' is not allowed in Product Version.", vbInformation)
                Exit Sub
            End If
            txtVCode.Text = ""
            txtGCode.Text = ""

            Try

                ' ALCrypto DLL with 1024-bit strength
                If optStrength0.Checked = True Then
                    Dim Key As New RSAKey
                    ReDim Key.data(32)
                    'Adding a delegate for AddressOf CryptoProgressUpdate did not work
                    'for now. Modified alcrypto3NET.dll to create a second generate function
                    'rsa_generate2 that does not deal with progress monitoring - ialkan
                    'retrieve url info

                    ' Get the current date format and save it to regionalSymbol variable
                    Get_locale()
                    ' Use this trick to temporarily set the date format to "yyyy/MM/dd"
                    Set_locale("")

                    Dim siteNameUrl As String
                    siteNameUrl = Request.ServerVariables.Get("SERVER_NAME")
                    Select Case siteNameUrl
                        Case "localhost"
                            If modALUGEN.rsa_generate2_local(Key, 1024) = RETVAL_ON_ERROR Then
                                Set_locale(regionalSymbol)
                                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, modALUGEN.ACTIVELOCKSTRING, STRRSAERROR)
                            End If
                        Case Else
                            If modALUGEN.rsa_generate2(Key, 1024) = RETVAL_ON_ERROR Then
                                Set_locale(regionalSymbol)
                                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, modALUGEN.ACTIVELOCKSTRING, STRRSAERROR)
                            End If
                    End Select

                    ' extract private and public key blobs
                    Dim strBlob As String
                    Dim blobLen As Integer
                    Select Case siteNameUrl
                        Case "localhost"
                            If rsa_public_key_blob_local(Key, vbNullString, blobLen) = RETVAL_ON_ERROR Then
                                Set_locale(regionalSymbol)
                                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
                            End If
                        Case Else
                            If rsa_public_key_blob(Key, vbNullString, blobLen) = RETVAL_ON_ERROR Then
                                Set_locale(regionalSymbol)
                                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
                            End If
                    End Select

                    If blobLen > 0 Then
                        strBlob = New String(Chr(0), blobLen)
                        Select Case siteNameUrl
                            Case "localhost"
                                If rsa_public_key_blob_local(Key, strBlob, blobLen) = RETVAL_ON_ERROR Then
                                    Set_locale(regionalSymbol)
                                    Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
                                End If
                            Case Else
                                If rsa_public_key_blob(Key, strBlob, blobLen) = RETVAL_ON_ERROR Then
                                    Set_locale(regionalSymbol)
                                    Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
                                End If
                        End Select
                        'System.Diagnostics.Debug.WriteLine("Public blob: " & strBlob)
                        txtVCode.Text = strBlob
                    End If

                    'FUTURE RSA IMPLEMENTATION USING NATIVE VB.NET FUNCTIONS
                    'Dim rsa As New RSACryptoServiceProvider
                    'Dim xmlPublicKey As String
                    'Dim xmlPrivateKey As String
                    'xmlPublicKey = rsa.ToXmlString(False)
                    'xmlPrivateKey = rsa.ToXmlString(True)
                    'rsa.FromXmlString(xmlPublicKey)
                    'txtCode1.Text = xmlPublicKey
                    'txtCode2.Text = xmlPrivateKey

                    Select Case siteNameUrl
                        Case "localhost"
                            rsa_private_key_blob_local(Key, vbNullString, blobLen)
                            If modALUGEN.rsa_private_key_blob_local(Key, vbNullString, blobLen) = RETVAL_ON_ERROR Then
                                Set_locale(regionalSymbol)
                                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
                            End If
                        Case Else
                            If modALUGEN.rsa_private_key_blob(Key, vbNullString, blobLen) = RETVAL_ON_ERROR Then
                                Set_locale(regionalSymbol)
                                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
                            End If
                    End Select

                    If blobLen > 0 Then
                        strBlob = New String(Chr(0), blobLen)
                        Select Case siteNameUrl
                            Case "localhost"
                                If modALUGEN.rsa_private_key_blob_local(Key, strBlob, blobLen) = RETVAL_ON_ERROR Then
                                    Set_locale(regionalSymbol)
                                    Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
                                End If
                            Case Else
                                If modALUGEN.rsa_private_key_blob(Key, strBlob, blobLen) = RETVAL_ON_ERROR Then
                                    Set_locale(regionalSymbol)
                                    Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
                                End If
                        End Select
                        'System.Diagnostics.Debug.WriteLine("Private blob: " & strBlob)
                        txtGCode.Text = strBlob
                    End If

                    ' done with the key - throw it away
                    Select Case siteNameUrl
                        Case "localhost"
                            If modALUGEN.rsa_freekey_local(Key) = RETVAL_ON_ERROR Then
                                Set_locale(regionalSymbol)
                                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
                            End If
                        Case Else
                            If modALUGEN.rsa_freekey(Key) = RETVAL_ON_ERROR Then
                                Set_locale(regionalSymbol)
                                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
                            End If
                    End Select

                    ' Test generated key for correctness by recreating it from the blobs
                    ' Note:
                    ' ====
                    ' Due to an outstanding bug in ALCrypto3NET.dll, sometimes this step will crash the app because
                    ' the generated keyset is bad.
                    ' The work-around for the time being is to keep regenerating the keyset until eventually
                    ' you'll get a valid keyset that no longer crashes.
                    Dim strdata As String : strdata = "This is a test string to be encrypted."
                    Select Case siteNameUrl
                        Case "localhost"
                            If modALUGEN.rsa_createkey_local(txtVCode.Text, txtVCode.Text.Length, txtGCode.Text, txtGCode.Text.Length, Key) = RETVAL_ON_ERROR Then
                                Set_locale(regionalSymbol)
                                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
                            End If
                            ' It worked! We're all set to go.
                            If modALUGEN.rsa_freekey_local(Key) = RETVAL_ON_ERROR Then
                                Set_locale(regionalSymbol)
                                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
                            End If
                        Case Else
                            If modALUGEN.rsa_createkey(txtVCode.Text, txtVCode.Text.Length, txtGCode.Text, txtGCode.Text.Length, Key) = RETVAL_ON_ERROR Then
                                Set_locale(regionalSymbol)
                                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
                            End If
                            ' It worked! We're all set to go.
                            If modALUGEN.rsa_freekey(Key) = RETVAL_ON_ERROR Then
                                Set_locale(regionalSymbol)
                                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
                            End If
                    End Select

                Else   ' CryptoAPI - RSA with given strength

                    Dim strPublicBlob As String, strPrivateBlob As String
                    Dim imodulus As Integer

                    If optStrength1.Checked = True Then
                        imodulus = 4096
                    ElseIf optStrength2.Checked = True Then
                        imodulus = 2048
                    ElseIf optStrength3.Checked = True Then
                        imodulus = 1536
                    ElseIf optStrength4.Checked = True Then
                        imodulus = 1024
                    ElseIf optStrength5.Checked = True Then
                        imodulus = 512
                    End If

                    'create new instance of RSACryptoServiceProvider
                    'Dim cspParam As CspParameters = New CspParameters
                    'cspParam.Flags = CspProviderFlags.UseMachineKeyStore
                    'cspParam.KeyContainerName = txtName.Text & txtVer.Text
                    'cspParam.KeyNumber = 2 'signature key pair
                    ''Set the CSP Provider Type PROV_RSA_FULL
                    'cspParam.ProviderType = 1
                    ''Set the CSP Provider Name
                    'cspParam.ProviderName = "Microsoft Base Cryptographic Provider v1.0"

                    'create new instance of RSACryptoServiceProvider
                    Dim rsaCSP As New System.Security.Cryptography.RSACryptoServiceProvider(imodulus)   ', cspParam)

                    'Generate public and private key data and allowing their exporting.
                    'True to include private parameters; otherwise, false

                    Dim rsaPubParams As RSAParameters       'stores public key
                    Dim rsaPrivateParams As RSAParameters   'stores private key
                    rsaPrivateParams = rsaCSP.ExportParameters(True)
                    rsaPubParams = rsaCSP.ExportParameters(False)

                    strPrivateBlob = rsaCSP.ToXmlString(True)
                    strPublicBlob = rsaCSP.ToXmlString(False)

                    'ok = ActiveLock3Globals_definst.ContainerChange(txtName.Text & txtVer.Text)
                    'ok = ActiveLock3Globals_definst.CryptoAPIAction(1, txtName.Text & txtVer.Text, "", "", strPublicBlob, strPrivateBlob, modulus)
                    txtVCode.Text = "RSA" & imodulus & strPublicBlob
                    txtGCode.Text = "RSA" & imodulus & strPrivateBlob

                    rsaPubParams = Nothing
                    rsaPrivateParams = Nothing
                    rsaCSP = Nothing

                End If
                Set_locale(regionalSymbol)

                'cmdGenerateCode.Enabled = False

                UpdateAddButtonStatus()
                Set_locale(regionalSymbol)
                SayAjax("Product codes generated successfully.")

            Catch ex As Exception
                Set_locale(regionalSymbol)
                SayAjax("Error!")

            End Try
        End Sub

        ' Validate and enable Add/Change button as appropriate
        Private Function CheckDuplicate(ByRef Name As String, ByRef Ver As String) As Boolean
            CheckDuplicate = False
            Dim dr As DataRow
            'dr = dt.Rows(grdProducts.SelectedIndex)
            Dim i As Integer
            For i = 0 To dt.Rows.Count - 1
                dr = dt.Rows(i)
                If CType(dr(0), String) = Name Then
                    If CType(dr(1), String) = Ver Then
                        CheckDuplicate = True
                        Exit Function
                    End If
                End If
            Next
        End Function

        Private Sub cmdValidateCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdValidateCode.Click
            Dim KEY As RSAKey = Nothing
            Dim strdata As String
            Dim strSig As String, rc As Integer

            If txtVCode.Text = "" Then
                SayAjax("Product VCode field is empty.")
                Exit Sub
            ElseIf txtGCode.Text = "" Then
                SayAjax("Product GCode field is empty.")
                Exit Sub
            ElseIf txtVCode.Text = "" And txtGCode.Text = "" Then
                SayAjax("Product VCode and GCode fields are blank. Nothing to validate.")
                Exit Sub ' nothing to validate
            End If

            strdata = "I love Activelock"

            ' Get the current date format and save it to regionalSymbol variable
            Get_locale()
            ' Use this trick to temporarily set the date format to "yyyy/MM/dd"
            Set_locale("")

            ' ALCrypto DLL with 1024-bit strength
            If strLeft(txtVCode.Text, 3) <> "RSA" Then
                ' Validate to keyset to make sure it's valid.
                SayAjax("Validating keyset...")
                ReDim KEY.data(32)

                Dim siteNameUrl As String
                siteNameUrl = Request.ServerVariables.Get("SERVER_NAME")
                Select Case siteNameUrl
                    Case "localhost"
                        rc = modALUGEN.rsa_createkey_local(txtVCode.Text, txtVCode.Text.Length, txtGCode.Text, txtGCode.Text.Length, KEY)
                    Case Else
                        rc = modALUGEN.rsa_createkey(txtVCode.Text, txtVCode.Text.Length, txtGCode.Text, txtGCode.Text.Length, KEY)
                End Select
                If rc = RETVAL_ON_ERROR Then
                    Set_locale(regionalSymbol)
                    SayAjax("Code not valid! " & vbCrLf & STRRSAERROR)
                    'UpdateStatus(txtName.Text & " (" & txtVer.Text & ") " & STRRSAERROR)
                    Exit Sub
                End If

                ' sign it
                strSig = RSASign(txtVCode.Text, txtGCode.Text, strdata, siteNameUrl)
                rc = RSAVerify(txtVCode.Text, strdata, strSig, siteNameUrl)
                Dim dr As DataRow
                dr = dt.Rows(grdProducts.SelectedIndex)
                If rc = 0 Then
                    SayAjax(CType(dr(0), String) & " (" & CType(dr(1), String) & ") validated successfully.")
                Else
                    SayAjax(CType(dr(0), String) & " (" & CType(dr(1), String) & ") GCode-VCode mismatch!")
                End If
                ' It worked! We're all set to go.
                Select Case siteNameUrl
                    Case "localhost"
                        If modALUGEN.rsa_freekey_local(KEY) = RETVAL_ON_ERROR Then
                            Set_locale(regionalSymbol)
                            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
                        End If
                    Case Else
                        If modALUGEN.rsa_freekey(KEY) = RETVAL_ON_ERROR Then
                            Set_locale(regionalSymbol)
                            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
                        End If
                End Select

            Else  '.NET RSA

                Try

                    ' ------------------ begin Message from Ismail ------------------
                    ' This code block is used to Encrypt/Sign and then Validate/Decrypt
                    ' a small size text.
                    ' This code uses "I love Activelock" to validate the given Public/Private key pair
                    ' If you try to do the same with a much longer string, these routines will fail
                    ' with a "Bad Length" error
                    ' Increasing the cpher strength (say from 1024 to 2048-bits) will allow you to
                    ' run this code with much longer data strings
                    ' Activelock DLL uses a different algorithm to sign/validate
                    ' This section is functional, but more than that it's provided here
                    ' as the entire solution for a typical RSA signing/validation algorithm
                    ' ------------------ end Message from Ismail ------------------

                    Dim rsaCSP As New System.Security.Cryptography.RSACryptoServiceProvider
                    Dim rsaPubParams As RSAParameters 'stores public key
                    Dim strPublicBlob, strPrivateBlob As String

                    strPublicBlob = txtVCode.Text
                    strPrivateBlob = txtGCode.Text

                    ' ENCRYPT PLAIN TEXT USING THE PUBLIC KEY
                    ' Convert the data string to a byte array.
                    Dim toEncrypt As Byte()     ' Holds message in bytes
                    Dim enc As New UTF8Encoding ' new instance of Unicode8 instance
                    Dim encrypted As Byte() ' holds encrypted data
                    Dim encryptedPlainText As String

                    If strLeft(txtVCode.Text, 6) = "RSA512" Then
                        strPublicBlob = strRight(txtVCode.Text, Len(txtVCode.Text) - 6)
                    Else
                        strPublicBlob = strRight(txtVCode.Text, Len(txtVCode.Text) - 7)
                    End If
                    rsaCSP.FromXmlString(strPublicBlob)
                    rsaPubParams = rsaCSP.ExportParameters(False)
                    ' import public key params into instance of RSACryptoServiceProvider
                    rsaCSP.ImportParameters(rsaPubParams)
                    toEncrypt = enc.GetBytes(strdata)


                    '' The following Encrypt method works for long and short strings
                    '' =============================== ACTIVATE FOR LONG AND SHORT STRINGS =============================
                    ''The RSA algorithm works on individual blocks of unencoded bytes. In this case, the
                    ''maximum is 58 bytes. Therefore, we are required to break up the text into blocks and
                    ''encrypt each one individually. Each encrypted block will give us an output of 128 bytes.
                    ''If we do not break up the blocks in this manner, we will throw a "key not valid for use
                    ''in specified state" exception

                    ''Get the size of the final block
                    'Const RSA_BLOCKSIZE As Integer = 58
                    'Dim lastBlockLength As Integer = toEncrypt.Length Mod RSA_BLOCKSIZE
                    'Dim blockCount As Integer = CType(Math.Floor(toEncrypt.Length / RSA_BLOCKSIZE), Integer) ' CType not necessary in VB 2008
                    'Dim hasLastBlock As Boolean = False
                    'If Not lastBlockLength.Equals(0) Then
                    '    'We need to create a final block for the remaining characters
                    '    blockCount += 1
                    '    hasLastBlock = True
                    'End If

                    ''Initialize the result buffer
                    'Dim result() As Byte = New Byte() {}

                    ''Initialize the RSA Service Provider with the public key
                    ''rsaCSP.FromXmlString(strPublicBlob) 'This was taken care of already

                    ''Break the text into blocks and work on each block individually
                    'For blockIndex As Integer = 0 To blockCount - 1
                    '    Dim thisBlockLength As Integer

                    '    'If this is the last block and we have a remainder, then set the length accordingly
                    '    If blockCount.Equals(blockIndex + 1) AndAlso hasLastBlock Then
                    '        thisBlockLength = lastBlockLength
                    '    Else
                    '        thisBlockLength = RSA_BLOCKSIZE
                    '    End If
                    '    Dim startChar As Integer = blockIndex * RSA_BLOCKSIZE

                    '    'Define the block that we will be working on
                    '    Dim currentBlock(thisBlockLength - 1) As Byte
                    '    Array.Copy(toEncrypt, startChar, currentBlock, 0, thisBlockLength)

                    '    'Encrypt the current block and append it to the result stream
                    '    Dim encryptedBlock() As Byte = rsaCSP.Encrypt(currentBlock, False)
                    '    Dim originalResultLength As Integer = result.Length

                    '    ReDim Preserve result(originalResultLength + encryptedBlock.Length) ' This is for VB 2008
                    '    'Array.Resize(result, originalResultLength + encryptedBlock.Length)

                    '    encryptedBlock.CopyTo(result, originalResultLength)
                    'Next

                    'encrypted = result
                    ' =============================================================================================

                    ' The following Encrypt method works only for short strings
                    encrypted = rsaCSP.Encrypt(toEncrypt, False)
                    encryptedPlainText = Convert.ToBase64String(encrypted) ' convert to base64/Radix output

                    ' HASH AND SIGN THE SIGNATURE
                    ' GENERATE SIGNATURE BLOCK USING SENDER'S PRIVATE KEY
                    Dim signatureBlock As String
                    ' Hash the encrypted data and generate a signature block on the hash
                    ' using the sender's private key. (Signature Block)
                    ' create new instance of SHA1 hash algorithm to compute hash
                    Dim hash As New SHA1Managed
                    Dim hashedData() As Byte ' a byte array to store hash value
                    If strLeft(txtGCode.Text, 6) = "RSA512" Then
                        strPrivateBlob = strRight(txtGCode.Text, Len(txtGCode.Text) - 6)
                    Else
                        strPrivateBlob = strRight(txtGCode.Text, Len(txtGCode.Text) - 7)
                    End If
                    ' import private key params into instance of RSACryptoServiceProvider
                    rsaCSP.FromXmlString(strPrivateBlob)
                    Dim rsaPrivateParams As RSAParameters 'stores private key
                    rsaPrivateParams = rsaCSP.ExportParameters(True)
                    rsaCSP.ImportParameters(rsaPrivateParams)
                    ' compute hash with algorithm specified as here we have SHA1
                    hashedData = hash.ComputeHash(encrypted)
                    ' Sign Data using private key and  OID is simple name of the algorithm for which to get the object identifier (OID)
                    Dim signature As Byte() ' holds signatures
                    signature = rsaCSP.SignHash(hashedData, CryptoConfig.MapNameToOID("SHA1"))
                    ' convert to base64/Radix output
                    signatureBlock = Convert.ToBase64String(signature)

                    ' VERIFY SIGNATURE BLOCK USING THE SENDER'S PUBLIC KEY
                    ' VALIDATE THE STRING WITH THE PUBLIC/PRIVATE KEY PAIR
                    ' Verify the signature is authentic using the sender's public key(decrypt Signature block)
                    Dim myencrypted() As Byte
                    Dim mysignature() As Byte
                    myencrypted = Convert.FromBase64String(encryptedPlainText)
                    mysignature = Convert.FromBase64String(signatureBlock)
                    ' create new instance of SHA1 hash algorithm to compute hash
                    Dim sha1hash As New SHA1Managed
                    Dim sha1hashedData() As Byte ' a byte array to store hash value
                    ' import  public key params into instance of RSACryptoServiceProvider
                    rsaCSP.ImportParameters(rsaPubParams)
                    ' compute hash with algorithm specified as here we have SHA1
                    sha1hashedData = sha1hash.ComputeHash(myencrypted)
                    ' Sign Data using public key and  OID is simple name of the algorithm for which to get the object identifier (OID)
                    Dim verified As Boolean
                    verified = rsaCSP.VerifyHash(sha1hashedData, CryptoConfig.MapNameToOID("SHA1"), mysignature)
                    If verified Then
                        SayAjax(txtProductName.Text & " (" & txtProductVersion.Text & ") validated successfully.")
                        'MsgBox("Signature Valid", MsgBoxStyle.Information)
                    Else
                        SayAjax(txtProductName.Text & " (" & txtProductVersion.Text & ") GCode-VCode mismatch!")
                        'MsgBox("Invalid Signature", MsgBoxStyle.Exclamation)
                    End If

                    ' THE FOLLOWING CODE BLOCK IS USED TO RETRIEVE THE ORIGINAL
                    ' STRING strData BUT IS NOT NEEDED FOR THE VALIDATION PROCESS
                    ' IT'S BEEN SHOWN HERE FOR DEMONSTRATION PURPOSES
                    ' This works for short strings only
                    Dim newencrypted() As Byte
                    newencrypted = Convert.FromBase64String(encryptedPlainText)
                    Dim fromEncrypt() As Byte ' a byte array to store decrypted bytes
                    Dim roundTrip As String ' holds original message
                    ' import  private key params into instance of RSACryptoServiceProvider
                    rsaCSP.ImportParameters(rsaPrivateParams)


                    '' The following Decrypt method works for long and short strings
                    '' It's currently not functioning correctly
                    '' =============================== ACTIVATE FOR LONG AND SHORT STRINGS =============================
                    ''When we encrypt a string using RSA, it works on individual blocks of up to
                    ''58 bytes. Each block generates an output of 128 encrypted bytes. Therefore, to
                    ''decrypt the message, we need to break the encrypted stream into individual
                    ''chunks of 128 bytes and decrypt them individually

                    ''Determine how many bytes are in the encrypted stream. The input is in hex format,
                    ''so we have to divide it by 2
                    'Const RSA_DECRYPTBLOCKSIZE As Integer = 128
                    'Dim maxBytes As Integer = CType(encryptedPlainText.Length / 2, Integer)  ' CType not necessary in VB 2008

                    ''Ensure that the length of the encrypted stream is divisible by 128
                    'If Not (maxBytes Mod RSA_DECRYPTBLOCKSIZE).Equals(0) Then
                    '    Throw New System.Security.Cryptography.CryptographicException("Encrypted text is an invalid length")
                    'End If

                    ''Calculate the number of blocks we will have to work on
                    'Dim blockCount2 As Integer = CType(maxBytes / RSA_DECRYPTBLOCKSIZE, Integer)

                    ''Initialize the result buffer
                    'Dim result2() As Byte = New Byte() {}

                    ''rsaCSP.FromXmlString(strPrivateBlob) ' This was done already

                    ''Iterate through each block and decrypt it
                    'For blockIndex As Integer = 0 To blockCount2 - 1
                    '    'Get the current block to work on
                    '    Dim currentBlockHex As String = encryptedPlainText.Substring(blockIndex * (RSA_DECRYPTBLOCKSIZE * 2), RSA_DECRYPTBLOCKSIZE * 2)
                    '    Dim currentBlockBytes As Byte() = HexToBytes(currentBlockHex)

                    '    'Decrypt the current block and append it to the result stream
                    '    Dim currentBlockDecrypted() As Byte = rsaCSP.Decrypt(currentBlockBytes, False)
                    '    Dim originalResultLength As Integer = result2.Length

                    '    ReDim Preserve result2(originalResultLength + currentBlockDecrypted.Length)
                    '    'Array.Resize(result, originalResultLength + currentBlockDecrypted.Length) ' This is for VB 2008

                    '    currentBlockDecrypted.CopyTo(result2, originalResultLength)
                    'Next
                    'fromEncrypt = result2
                    ' =============================================================================================


                    ' The following Decrypt works for short strings only
                    'store decrypted data into byte array
                    fromEncrypt = rsaCSP.Decrypt(newencrypted, False)

                    'convert bytes to string
                    roundTrip = enc.GetString(fromEncrypt)
                    If roundTrip <> strdata Then
                        SayAjax(txtProductName.Text & " (" & txtProductVersion.Text & ") GCode-VCode mismatch!")
                    End If

                    'Release any resources held by the RSA Service Provider
                    rsaCSP.Clear()
                    Set_locale(regionalSymbol)

                Catch ex As Exception
                    Set_locale(regionalSymbol)
                    SayAjax(ex.Message & ex.StackTrace)
                End Try
            End If

            Exit Sub
exitValidate:
            Set_locale(regionalSymbol)
            SayAjax(txtProductName.Text & " (" & txtProductVersion.Text & ") GCode-VCode mismatch!")

        End Sub

        Private Sub cboLicenseType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboLicenseType.SelectedIndexChanged
            ResetCalendarControls()
            Select Case cboLicenseType.SelectedValue
                Case "0" 'time locked
                    lblExpiry.Text = "Expiration:   "
                    txtDays.Text = Date.Now.AddDays(365).ToString("yyyy/MM/dd")
                    'lblDays.Text = "yyyy/MM/dd"
                    'txtDays.Visible = False
                    'dtpExpireDate.Visible = True
                    'dtpExpireDate.Value = Date.Now.AddDays(30)
                    cmdSelectExpireDate.Visible = True
                    txtDays.ReadOnly = False
                    txtDays.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
                    txtDays.ForeColor = System.Drawing.Color.Black
                Case "1" 'periodic
                    cmdSelectExpireDate.Visible = False
                    lblExpiry.Text = "Expires after:"
                    txtDays.Text = "365"
                    'lblDays.Text = "Day(s)"
                    'txtDays.Visible = True
                    'dtpExpireDate.Visible = False
                    txtDays.ReadOnly = False
                    txtDays.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
                    txtDays.ForeColor = System.Drawing.Color.Black
                Case "2" 'permanent
                    lblExpiry.Text = "Permanent     "
                    cmdSelectExpireDate.Visible = False
                    txtDays.ReadOnly = True
                    txtDays.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
                    txtDays.ForeColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            End Select
        End Sub

        Private Sub cmdSelectExpireDate_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles cmdSelectExpireDate.Click
            If myCalendar.Visible = False Then
                Calendar1.Visible = True
                myCalendar.Visible = True

                Dim myIframe As New HtmlGenericControl("IFRAME")
                myIframe.Attributes.Add("style", "position: absolute;z-index:-1;border: none;margin: 0px 0px 0px 0px; background-color:Transparent; height: 180px; width: 200px;")
                myIframe.Attributes.Add("src", "javascript:false;")
                myIframe.Attributes.Add("frameBorder", "0")
                myIframe.Attributes.Add("scrolling", "no")

                myCalendar.Attributes.Add("style", "background-color: white; border:none; padding:0px; position: absolute; top: 158px; left: 308px; z-Index: 100;  height: 180px; width: 200px;")
                Calendar1.Attributes.Add("style", "background-color: white; border:solid 1px #90b0E0; padding:0px; position: absolute; top: 0px; left: 0px; z-Index: 100;")

                myIframe.Visible = True
                myCalendar.Controls.Add(myIframe)
            Else
                Calendar1.Visible = False
                myCalendar.Visible = False
            End If
        End Sub

        Private Sub Calendar1_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged
            txtDays.Text = Calendar1.SelectedDate.ToString("yyyy/MM/dd")
            Calendar1.Visible = False
            myCalendar.Visible = False
        End Sub

        Private Sub Calendar1_VisibleMonthChanged(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.MonthChangedEventArgs) Handles Calendar1.VisibleMonthChanged
            Calendar1.Visible = True
            myCalendar.Visible = True

            Dim myIframe As New HtmlGenericControl("IFRAME")
            myIframe.Attributes.Add("style", "position: absolute;z-index:-1;border: none;margin: 0px 0px 0px 0px; background-color:Transparent; height: 180px; width: 200px;")
            myIframe.Attributes.Add("src", "javascript:false;")
            myIframe.Attributes.Add("frameBorder", "0")
            myIframe.Attributes.Add("scrolling", "no")

            myCalendar.Controls.Add(myIframe)

        End Sub

        Private Sub ResetCalendarControls()
            Calendar1.Visible = False
            myCalendar.Visible = False
        End Sub

        Private Sub cboProduct_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboProduct.SelectedIndexChanged
            ResetCalendarControls()
            'product selected from products combo - update the controls
            UpdateKeyGenButtonStatus()

            If cboProducts_Array Is Nothing Then Exit Sub

            If cboProducts_Array(cboProduct.SelectedIndex) = "512" Then
                lblKeyStrength.Text = "[Key Strength: CryptoAPI RSA 512-bit]"
            ElseIf cboProducts_Array(cboProduct.SelectedIndex) = "1024" Then
                lblKeyStrength.Text = "[Key Strength: CryptoAPI RSA 1024-bit]"
            ElseIf cboProducts_Array(cboProduct.SelectedIndex) = "1536" Then
                lblKeyStrength.Text = "[Key Strength: CryptoAPI RSA 1536-bit]"
            ElseIf cboProducts_Array(cboProduct.SelectedIndex) = "2048" Then
                lblKeyStrength.Text = "[Key Strength: CryptoAPI RSA 2048-bit]"
            ElseIf cboProducts_Array(cboProduct.SelectedIndex) = "4096" Then
                lblKeyStrength.Text = "[Key Strength: CryptoAPI RSA 4096-bit]"
            ElseIf cboProducts_Array(cboProduct.SelectedIndex) = "0" Then
                lblKeyStrength.Text = "[Key Strength: ALCrypto 1024-bit]"
            End If
        End Sub

        Private Sub cboRegLevel_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboRegLevel.SelectedIndexChanged
            ResetCalendarControls()
        End Sub

        Private Sub PopulateUI(ByVal ProdInfo As ProductInfo)
            With ProdInfo
                AddRow(.Name, .Version, .VCode, .GCode, False)
            End With
        End Sub

        Private Sub AddRow(ByRef Name As String, ByRef Ver As String, ByRef Code1 As String, ByRef Code2 As String, Optional ByRef fUpdateStore As Boolean = True)

            ' Add a Product Row to the GUI.
            ' If fUpdateStore is True, then product info is also saved to the store.

            ' Update the view
            Call CreateDataSource(Name, Ver, Code1, Code2)

            Dim ProdInfo As ProductInfo
            ProdInfo = ActiveLock3AlugenGlobals_definst.CreateProductInfo(Name, Ver, Code1, Code2)
            If fUpdateStore Then
                Call GeneratorInstance.SaveProduct(ProdInfo)
                grdProducts.DataSource = dt
                grdProducts.DataBind()
                grdProducts.SelectedIndex = dt.Rows.Count - 1
            End If
            If Not MagicAjax.MagicAjaxContext.Current.IsAjaxCall Then
                'cboProduct.Items.Add(New ListItem(Name & " - " & Ver, Name & "|" & Ver))
                cboProduct.Items.Add(New ListItem(Name & " (" & Ver & ")", Name & "|" & Ver))

                ReDim Preserve cboProducts_Array(cboProduct.Items.Count - 1)
                If Code1.Contains("RSA512") Then
                    cboProducts_Array(cboProduct.Items.Count - 1) = "512"
                ElseIf Code1.Contains("RSA1024") Then
                    cboProducts_Array(cboProduct.Items.Count - 1) = "1024"
                ElseIf Code1.Contains("RSA1536") Then
                    cboProducts_Array(cboProduct.Items.Count - 1) = "1536"
                ElseIf Code1.Contains("RSA2048") Then
                    cboProducts_Array(cboProduct.Items.Count - 1) = "2048"
                ElseIf Code1.Contains("RSA4096") Then
                    cboProducts_Array(cboProduct.Items.Count - 1) = "4096"
                Else ' ALCrypto 1024-bit
                    cboProducts_Array(cboProduct.Items.Count - 1) = "0"
                End If

            End If
        End Sub

        Private Sub UpdateCodeGenButtonStatus()
            If txtProductName.Text = "" Or txtProductVersion.Text = "" Then
                cmdGenerateCode.Enabled = False
            ElseIf CheckDuplicate(txtProductName.Text, txtProductVersion.Text) Then
                cmdGenerateCode.Enabled = False
            Else
                cmdGenerateCode.Enabled = True
            End If
        End Sub
        Private Sub UpdateAddButtonStatus()
            If txtProductName.Text = "" Or txtProductVersion.Text = "" Or txtVCode.Text = "" Then
                cmdAddProduct.Enabled = False
            ElseIf CheckDuplicate(txtProductName.Text, txtProductVersion.Text) Then
                cmdAddProduct.Enabled = False
            Else
                cmdAddProduct.Enabled = True
            End If
        End Sub
        Private Sub txtLicenseKey_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLicenseKey.TextChanged
            cmdSaveLicenseFile.Enabled = CBool(txtLicenseKey.Text.Length > 0)
        End Sub
        Private Sub txtproductName_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtProductName.TextChanged
            UpdateCodeGenButtonStatus()
            UpdateAddButtonStatus()
        End Sub
        Private Sub txtProductVersion_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtProductVersion.TextChanged
            ' New product - will be processed by Add command
            UpdateCodeGenButtonStatus()
            UpdateAddButtonStatus()
        End Sub

        Private Sub txtUser_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUserName.TextChanged
            If systemEvent Then Exit Sub
            If Len(txtInstallCode.Text) <> 8 Then txtInstallCode.Text = ActiveLock.InstallationCode(Trim(txtUserName.Text))
        End Sub

        Private Sub txtInstallCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtInstallCode.TextChanged
            'extract username
            If txtInstallCode.Text.Length > 0 Then
                txtUserName.Text = GetUserFromInstallCode(txtInstallCode.Text)
                cmdGenerateLicenseKey.Enabled = True
                cmdEmailLicenseKey.Enabled = True
                cmdPrintLicenseKey.Enabled = True
                cmdSaveLicenseFile.Enabled = True
            Else
                txtUserName.Text = ""
                cmdGenerateLicenseKey.Enabled = False
                cmdEmailLicenseKey.Enabled = False
                cmdPrintLicenseKey.Enabled = False
                cmdSaveLicenseFile.Enabled = False
            End If
        End Sub

        Private Function GetUserFromInstallCode(ByVal strInstCode As String) As String
            ' Retrieves lock string and user info from the request string
            '
            Dim a() As String
            Dim i As Integer
            Dim aString As String
            Dim usedLockNone As Boolean
            Dim noKey As String
            noKey = Chr(110) & Chr(111) & Chr(107) & Chr(101) & Chr(121)

            GetUserFromInstallCode = ""
            If strInstCode Is Nothing Or strInstCode = "" Then Exit Function
            strInstCode = ActiveLock3Globals_definst.Base64Decode(strInstCode)

            If Not strInstCode Is Nothing _
              AndAlso strInstCode.Substring(0, 1) = "+" Then
                strInstCode = strInstCode.Substring(1)
                usedLockNone = True
            End If

            Dim arrProdVer() As String
            arrProdVer = Split(strInstCode, "&&&") ' Extract the software name and version
            strInstCode = arrProdVer(0)

            systemEvent = True
            'clean checkboxes
            chkLockMACaddress.Enabled = True
            chkLockComputer.Enabled = True
            chkLockHD.Enabled = True
            chkLockHDfirmware.Enabled = True
            chkLockWindows.Enabled = True
            chkLockBIOS.Enabled = True
            chkLockMotherboard.Enabled = True
            chkLockIP.Enabled = True
            chkLockExternalIP.Enabled = True
            chkLockFingerprint.Enabled = True
            chkLockMemory.Enabled = True
            chkLockCPUID.Enabled = True
            chkLockBaseboardID.Enabled = True
            chkLockVideoID.Enabled = True

            a = Split(strInstCode, vbLf)
            If usedLockNone = True Then
                For i = LBound(a) To UBound(a) - 1
                    aString = a(i)
                    If i = LBound(a) Then
                        MACaddress = aString
                        lblLockMacAddress.Text = MACaddress
                    ElseIf i = LBound(a) + 1 Then
                        ComputerName = aString
                        lblLockComputer.Text = ComputerName
                    ElseIf i = LBound(a) + 2 Then
                        VolumeSerial = aString
                        lblLockHD.Text = VolumeSerial
                    ElseIf i = LBound(a) + 3 Then
                        FirmwareSerial = aString
                        lblLockHDfirmware.Text = FirmwareSerial
                    ElseIf i = LBound(a) + 4 Then
                        WindowsSerial = aString
                        lblLockWindows.Text = WindowsSerial
                    ElseIf i = LBound(a) + 5 Then
                        BIOSserial = aString
                        lblLockBIOS.Text = BIOSserial
                    ElseIf i = LBound(a) + 6 Then
                        MotherboardSerial = aString
                        lblLockMotherboard.Text = MotherboardSerial
                    ElseIf i = LBound(a) + 7 Then
                        IPaddress = aString
                        lblLockIP.Text = IPaddress
                    ElseIf i = LBound(a) + 8 Then
                        ExternalIP = aString
                        lblLockExternalIP.Text = ExternalIP
                    ElseIf i = LBound(a) + 9 Then
                        Fingerprint = aString
                        lblLockFingerprint.Text = Fingerprint
                    ElseIf i = LBound(a) + 10 Then
                        Memory = aString
                        lblLockMemory.Text = Memory
                    ElseIf i = LBound(a) + 11 Then
                        CPUID = aString
                        lblLockCPUID.Text = CPUID
                    ElseIf i = LBound(a) + 12 Then
                        BaseboardID = aString
                        lblLockBaseboardID.Text = BaseboardID
                    ElseIf i = LBound(a) + 13 Then
                        VideoID = aString
                        lblLockVideoID.Text = VideoID
                    End If
                Next i
            Else '"+" was not used, therefore one or more lockTypes were specified in the application
                chkLockMACaddress.Enabled = False
                chkLockHD.Enabled = False
                chkLockHDfirmware.Enabled = False
                chkLockWindows.Enabled = False
                chkLockComputer.Enabled = False
                chkLockBIOS.Enabled = False
                chkLockMotherboard.Enabled = False
                chkLockIP.Enabled = False
                chkLockExternalIP.Enabled = False
                chkLockFingerprint.Enabled = False
                chkLockMemory.Enabled = False
                chkLockCPUID.Enabled = False
                chkLockBaseboardID.Enabled = False
                chkLockVideoID.Enabled = False

                chkLockMACaddress.Checked = False
                chkLockHD.Checked = False
                chkLockHDfirmware.Checked = False
                chkLockWindows.Checked = False
                chkLockComputer.Checked = False
                chkLockBIOS.Checked = False
                chkLockMotherboard.Checked = False
                chkLockIP.Checked = False
                chkLockExternalIP.Checked = False
                chkLockFingerprint.Checked = False
                chkLockMemory.Checked = False
                chkLockCPUID.Checked = False
                chkLockBaseboardID.Checked = False
                chkLockVideoID.Checked = False

                For i = LBound(a) To UBound(a) - 1
                    aString = a(i)
                    If i = LBound(a) And aString <> noKey Then
                        MACaddress = aString
                        lblLockMacAddress.Text = MACaddress
                        chkLockMACaddress.Checked = True
                    ElseIf i = (LBound(a) + 1) And aString <> noKey Then
                        ComputerName = aString
                        lblLockComputer.Text = ComputerName
                        chkLockComputer.Checked = True
                    ElseIf i = (LBound(a) + 2) And aString <> noKey Then
                        VolumeSerial = aString
                        lblLockHD.Text = VolumeSerial
                        chkLockHD.Checked = True
                    ElseIf i = (LBound(a) + 3) And aString <> noKey Then
                        FirmwareSerial = aString
                        lblLockHDfirmware.Text = FirmwareSerial
                        chkLockHDfirmware.Checked = True
                    ElseIf i = (LBound(a) + 4) And aString <> noKey Then
                        WindowsSerial = aString
                        lblLockWindows.Text = WindowsSerial
                        chkLockWindows.Checked = True
                    ElseIf i = (LBound(a) + 5) And aString <> noKey Then
                        BIOSserial = aString
                        lblLockBIOS.Text = BIOSserial
                        chkLockBIOS.Checked = True
                    ElseIf i = (LBound(a) + 6) And aString <> noKey Then
                        MotherboardSerial = aString
                        lblLockMotherboard.Text = MotherboardSerial
                        chkLockMotherboard.Checked = True
                    ElseIf i = (LBound(a) + 7) And aString <> noKey Then
                        IPaddress = aString
                        lblLockIP.Text = IPaddress
                        chkLockIP.Checked = True
                    ElseIf i = (LBound(a) + 8) And aString <> noKey Then
                        ExternalIP = aString
                        lblLockExternalIP.Text = ExternalIP
                        chkLockExternalIP.Checked = True
                    ElseIf i = (LBound(a) + 9) And aString <> noKey Then
                        Fingerprint = aString
                        lblLockFingerprint.Text = Fingerprint
                        chkLockFingerprint.Checked = True
                    ElseIf i = (LBound(a) + 10) And aString <> noKey Then
                        Memory = aString
                        lblLockMemory.Text = Memory
                        chkLockMemory.Checked = True
                    ElseIf i = (LBound(a) + 11) And aString <> noKey Then
                        CPUID = aString
                        lblLockCPUID.Text = CPUID
                        chkLockCPUID.Checked = True
                    ElseIf i = (LBound(a) + 12) And aString <> noKey Then
                        BaseboardID = aString
                        lblLockBaseboardID.Text = BaseboardID
                        chkLockBaseboardID.Checked = True
                    ElseIf i = (LBound(a) + 13) And aString <> noKey Then
                        VideoID = aString
                        lblLockVideoID.Text = VideoID
                        chkLockVideoID.Checked = True
                    End If

                Next i
            End If

            If VolumeSerial = "Not Available" Or VolumeSerial = "0000-0000" Then
                chkLockHD.Enabled = False
                chkLockHD.Checked = False
            End If
            If MACaddress = "00 00 00 00 00 00" Or MACaddress = "00-00-00-00-00-00" Or MACaddress = "" Or MACaddress = "Not Available" Then
                chkLockMACaddress.Enabled = False
                chkLockMACaddress.Checked = False
            End If
            If FirmwareSerial = "Not Available" Then
                chkLockHDfirmware.Enabled = False
                chkLockHDfirmware.Checked = False
            End If
            If BIOSserial = "Not Available" Then
                chkLockBIOS.Enabled = False
                chkLockBIOS.Checked = False
            End If
            If MotherboardSerial = "Not Available" Then
                chkLockMotherboard.Enabled = False
                chkLockMotherboard.Checked = False
            End If
            If IPaddress = "Not Available" Then
                chkLockIP.Enabled = False
                chkLockIP.Checked = False
            End If
            If ExternalIP = "Not Available" Then
                chkLockExternalIP.Enabled = False
                chkLockExternalIP.Checked = False
            End If
            If Fingerprint = "Not Available" Then
                chkLockFingerprint.Enabled = False
                chkLockFingerprint.Checked = False
            End If
            If Memory = "Not Available" Then
                chkLockMemory.Enabled = False
                chkLockMemory.Checked = False
            End If
            If CPUID = "Not Available" Then
                chkLockCPUID.Enabled = False
                chkLockCPUID.Checked = False
            End If
            If BaseboardID = "Not Available" Then
                chkLockBaseboardID.Enabled = False
                chkLockBaseboardID.Checked = False
            End If
            If VideoID = "Not Available" Then
                chkLockVideoID.Enabled = False
                chkLockVideoID.Checked = False
            End If

            GetUserFromInstallCode = a(a.Length - 1)
            systemEvent = False

        End Function

        Private Sub loadRegisteredLevels()
            'load RegisteredLevels
            If Not File.Exists(strRegisteredLevelDBName) Then
                With cboRegLevel
                    .Items.Clear()
                    .Items.Add(New ListItem("Limited A", "0"))
                    .Items.Add(New ListItem("Limited B", "1"))
                    .Items.Add(New ListItem("Limited C", "2"))
                    .Items.Add(New ListItem("Limited D", "3"))
                    .Items.Add(New ListItem("Limited E", "4"))
                    .Items.Add(New ListItem("No Print/Save", "5"))
                    .Items.Add(New ListItem("Educational A", "6"))
                    .Items.Add(New ListItem("Educational B", "7"))
                    .Items.Add(New ListItem("Educational C", "8"))
                    .Items.Add(New ListItem("Educational D", "9"))
                    .Items.Add(New ListItem("Educational E", "10"))
                    .Items.Add(New ListItem("Level 1", "11"))
                    .Items.Add(New ListItem("Level 2", "12"))
                    .Items.Add(New ListItem("Level 3", "13"))
                    .Items.Add(New ListItem("Level 4", "14"))
                    .Items.Add(New ListItem("Light Version", "15"))
                    .Items.Add(New ListItem("Pro Version", "16"))
                    .Items.Add(New ListItem("Enterprise Version", "17"))
                    .Items.Add(New ListItem("Demo Only", "18"))
                    .Items.Add(New ListItem("Full Version-Europe", "19"))
                    .Items.Add(New ListItem("Full Version-Africa", "20"))
                    .Items.Add(New ListItem("Full Version-Asia", "21"))
                    .Items.Add(New ListItem("Full Version-USA", "22"))
                    .Items.Add(New ListItem("Full Version-International", "23"))
                    .SelectedIndex = 0
                    SaveComboBox(strRegisteredLevelDBName, CType(cboRegLevel, ListControl), True)
                End With
            Else
                LoadComboBox(strRegisteredLevelDBName, CType(cboRegLevel, ListControl), True)
            End If
        End Sub

        Private Sub cmdGenerateLicenseKey_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGenerateLicenseKey.Click
            'generate license key
            Dim usedVCode As String = Nothing
            Dim licFlag As ActiveLock3_6NET.ProductLicense.LicFlags, maximumUsers As Short

            If txtInstallCode.Text.Length < 8 Then Exit Sub

            ' Get the current date format and save it to regionalSymbol variable
            Get_locale()
            ' Use this trick to temporarily set the date format to "yyyy/MM/dd"
            Set_locale("")

            If cboLicenseType.Text = "Time Locked" Then
                ' Check to see if there's a valid expiration date
                If CDate(CType(txtDays.Text, DateTime).ToString("yyyy/MM/dd")) < CDate(Format(Date.Now, "yyyy/MM/dd")) Then
                    Set_locale(regionalSymbol)
                    MsgBox("Entered date occurs in the past.", vbExclamation)
                    Exit Sub
                End If
            End If

            If txtInstallCode.Text.Length <> 8 Then  'Not a Short Key License
                If chkLockMACaddress.Checked = False _
                  And chkLockComputer.Checked = False _
                  And chkLockHD.Checked = False _
                  And chkLockHDfirmware.Checked = False _
                  And chkLockWindows.Checked = False _
                  And chkLockBIOS.Checked = False _
                  And chkLockMotherboard.Checked = False _
                  And chkLockExternalIP.Checked = False _
                  And chkLockFingerprint.Checked = False _
                  And chkLockMemory.Checked = False _
                  And chkLockCPUID.Checked = False _
                  And chkLockBaseboardID.Checked = False _
                  And chkLockVideoID.Checked = False _
                  And chkLockIP.Checked = False Then
                    SayAjax("Warning: You did not select any hardware keys to lock the license." & vbCrLf & "This license will be machine independent. License will be locked to the username only !!!")
                End If
            End If

            systemEvent = True
            If Len(txtInstallCode.Text) <> 8 Then
                Dim dummyUser As String = GetUserFromInstallCode(txtInstallCode.Text)
                txtInstallCode.Text = ReconstructedInstallationCode()
            End If
            systemEvent = False

            Try
                Dim strName, strVer As String
                Dim itemProduct As ListItem = cboProduct.SelectedItem
                Dim a As String()

                a = itemProduct.Value.Split(Convert.ToChar("|"))
                strName = a(0)
                strVer = a(1)

                ActiveLock = ActiveLock3Globals_definst.NewInstance()
                With ActiveLock
                    .SoftwareName = strName
                    .SoftwareVersion = strVer
                End With

                Dim Generator As New AlugenGlobals
                GeneratorInstance = Generator.GeneratorInstance(IActiveLock.ProductsStoreType.alsINIFile)
                GeneratorInstance.StoragePath = AppPath() & "\licenses.ini"

                Dim varLicType As ProductLicense.ALLicType

                If cboLicenseType.SelectedValue = "1" Then 'Periodic
                    varLicType = ProductLicense.ALLicType.allicPeriodic
                ElseIf cboLicenseType.SelectedValue = "2" Then 'Permanent
                    varLicType = ProductLicense.ALLicType.allicPermanent
                ElseIf cboLicenseType.SelectedValue = "0" Then 'Time Locked
                    varLicType = ProductLicense.ALLicType.allicTimeLocked
                Else
                    varLicType = ProductLicense.ALLicType.allicNone
                End If

                Dim strExpire As String
                strExpire = GetExpirationDate()
                Dim strRegDate As String
                strRegDate = Date.UtcNow.ToString("yyyy/MM/dd")

                Dim Lic As ProductLicense
                'generate license object
                Dim selRegLevel As ListItem = cboRegLevel.SelectedItem
                Dim selRegLevelType As String
                If chkUseItemData.Checked Then
                    selRegLevelType = selRegLevel.Value
                Else
                    selRegLevelType = selRegLevel.Text
                End If

                'Take care of the networked licenses
                If chkNetworkedLicense.Checked = True Then
                    licFlag = ProductLicense.LicFlags.alfMulti
                Else
                    licFlag = ProductLicense.LicFlags.alfSingle
                End If
                If txtMaxCount.Text = "" Then
                    maximumUsers = 1
                Else
                    maximumUsers = CShort(txtMaxCount.Text)
                End If

                Lic = ActiveLock3Globals_definst.CreateProductLicense(strName, strVer, "", _
                          licFlag, varLicType, "", _
                          selRegLevelType, _
                          strExpire, , strRegDate, , maximumUsers)

                Dim strLibKey As String, i As Integer
                If Len(txtInstallCode.Text) = 8 Then  'Short Key License
                    Dim arrProdInfos() As ProductInfo
                    Dim MyGen As New AlugenGlobals
                    GeneratorInstance = MyGen.GeneratorInstance(IActiveLock.ProductsStoreType.alsINIFile)
                    GeneratorInstance.StoragePath = AppPath() & "\licenses.ini"
                    arrProdInfos = GeneratorInstance.RetrieveProducts()
                    For i = 0 To arrProdInfos.Length - 1
                        If arrProdInfos(i).Name = strName Then
                            usedVCode = arrProdInfos(i).VCode
                        End If
                    Next

                    'For i = 0 To grdProducts.Items.Count
                    '    If strName = grdProducts.Items(i).Text And strVer = lstvwProducts.Items(i).SubItems(1).Text Then
                    '        usedVCode = grdProducts.Items(i).SubItems(2).Text
                    '        Exit For
                    '    End If
                    'Next
                    strLibKey = ActiveLock.GenerateShortKey(usedVCode, txtInstallCode.Text, Trim(txtUserName.Text), strExpire, varLicType, cboRegLevel.SelectedIndex + 200, maximumUsers)
                    txtLicenseKey.Text = strLibKey
                Else 'ALCrypto License Key
                    ' Pass it to IALUGenerator to generate the key
                    Dim selectedRegisteredLevel As String
                    Dim mList As ListItem
                    mList = cboRegLevel.SelectedItem
                    If chkUseItemData.Checked Then
                        selectedRegisteredLevel = mList.Value
                    Else
                        selectedRegisteredLevel = mList.Text
                    End If

                    strLibKey = GeneratorInstance.GenKey(Lic, txtInstallCode.Text, selectedRegisteredLevel)
                    'split license key into 64byte chunks
                    txtLicenseKey.Text = Make64ByteChunks(strLibKey & "aLck" & txtInstallCode.Text)
                    'save the license to local server file - for use in save
                    SaveLicenseKey(txtLicenseKey.Text, strName & strVer & "_" & Session.SessionID.ToString & ".all")

                    If sender Is cmdGenerateLicenseKey Then
                        SayAjax("License code generated successfuly!")
                    End If
                End If
                Set_locale(regionalSymbol)

            Catch ex As Exception
                Set_locale(regionalSymbol)
                SayAjax("Error: " & ex.Message)
            Finally
                Set_locale(regionalSymbol)
            End Try
        End Sub

        Private Function GetExpirationDate() As String
            If cboLicenseType.SelectedValue = "0" Then 'Time Locked
                If txtDays.Text.Trim.Length = 0 Then txtDays.Text = Date.UtcNow.AddDays(30).ToString("yyyy/MM/dd")
                GetExpirationDate = CType(txtDays.Text, DateTime).ToString("yyyy/MM/dd")
            Else
                If txtDays.Text.Trim.Length = 0 Then txtDays.Text = "30"
                GetExpirationDate = Date.UtcNow.AddDays(CShort(txtDays.Text)).ToString("yyyy/MM/dd")
            End If
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

        Private Sub cmdEmailLicenseKey_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEmailLicenseKey.Click
            Dim mailToString As String
            Dim emailAddress As String = "user@company.com"
            Dim strSubject As String
            Dim strBodyMessage As String
            Dim strNewLine As String = "%0D%0A"

            'ensure it is correct key
            cmdGenerateLicenseKey_Click(sender, e)

            strSubject = String.Format("License key for application {0}, user [{1}]", cboProduct.SelectedItem.Text, txtUserName.Text)
            strBodyMessage = strNewLine & String.Format("Install code:" & strNewLine & "{0}", txtInstallCode.Text)
            strBodyMessage = strBodyMessage & strNewLine & strNewLine & String.Format("License key:" & strNewLine & "{0}", txtLicenseKey.Text.Replace(Chr(13), strNewLine).Replace(Chr(10), ""))

            'final constructor
            mailToString = String.Format("mailto:{0}?subject={1}&body={2}", emailAddress, strSubject, strBodyMessage)

            AjaxCallHelper.Redirect(mailToString)
        End Sub

        Private Sub cmdSaveLicenseFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSaveLicenseFile.Click
            Dim strName, strVer As String
            Dim itemProduct As ListItem = cboProduct.SelectedItem
            Dim a As String()

            a = itemProduct.Value.Split(Convert.ToChar("|"))
            strName = a(0)
            strVer = a(1)

            'ensure it is correct key and license file generated
            cmdGenerateLicenseKey_Click(sender, e)
            'dowload file
            DownloadFile(AppPath() & "\" & "GenKeys" & "\" & strName & strVer & "_" & Session.SessionID.ToString & ".all", True)

        End Sub

        Private Sub DownloadFile(ByVal fname As String, ByVal forceDownload As Boolean)
            Dim fullpath As String = Path.GetFullPath(fname)
            'Dim name As String = path.GetFileName(fullpath)
            Dim ext As String = Path.GetExtension(fullpath)
            Dim contenttype As String = String.Empty

            If Not IsDBNull(ext) Then
                ext = LCase(ext)
            End If

            Select Case ext
                Case ".htm", ".html"
                    contenttype = "text/HTML"
                Case ".txt"
                    contenttype = "text/plain"
                Case ".doc", ".rtf"
                    contenttype = "Application/msword"
                Case ".csv", ".xls"
                    contenttype = "Application/x-msexcel"
                Case Else
                    'contenttype = "text/plain"
                    contenttype = "application\octet-stream" '"application/zip"
            End Select

            HttpContext.Current.Response.Clear()
            HttpContext.Current.Response.ClearHeaders()
            HttpContext.Current.Response.ClearContent()

            Response.ContentType = contenttype

            If (forceDownload) Then
                Response.AppendHeader("content-disposition", _
                "attachment; filename=ASPNETAlugen3" & ext)
            End If

            Response.WriteFile(fullpath)
            Response.End()
        End Sub

        Private Sub SaveLicenseKey(ByVal sLibKey As String, ByVal sFileName As String)
            Dim fp As StreamWriter
            Try
                fp = File.CreateText(Server.MapPath(".\" & "GenKeys" & "\") & sFileName)
                fp.WriteLine(sLibKey)
                fp.Close()
            Catch err As Exception
                SayAjax(err.Message)
            Finally
            End Try
        End Sub

        Private Sub cmdPrintLicenseKey_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdPrintLicenseKey.Click
            Dim msgToPrint As New StringBuilder(String.Empty)
            Dim mWidth As Integer = 600
            Dim mHeight As Integer = 340

            'ensure it is correct key
            cmdGenerateLicenseKey_Click(sender, e)

            'start html
            msgToPrint.Append("<html><head>")
            msgToPrint.Append("<script language='Javascript1.2'>function printPage() { window.print(); }</script>")
            msgToPrint.Append("</head><body onload='printPage();' style='font-family:Courier;font-size:0.8em;'>")
            'body content
            msgToPrint.Append("<table width='100%' border='1' cellpadding='0' cellspacing='0' style='font-family:Courier;font-size:0.8em;'>")
            'title
            msgToPrint.Append("<tr>")
            msgToPrint.Append("<td colspan='2' valign=middle style='background-color:#808080;color:#f0f0f0;font-family:Courier;font-size:1.2em;text-align:center;height:30;'>")
            msgToPrint.Append("License Key")
            msgToPrint.Append("</td>")
            msgToPrint.Append("</tr>")
            'Software name
            msgToPrint.Append("<tr>")
            msgToPrint.Append("<td width=150>")
            msgToPrint.Append("Software name")
            msgToPrint.Append("</td>")
            msgToPrint.Append("<td>")
            msgToPrint.Append(cboProduct.SelectedItem.Text)
            msgToPrint.Append("</td>")
            msgToPrint.Append("</tr>")
            'Registered level
            msgToPrint.Append("<tr>")
            msgToPrint.Append("<td width=150>")
            msgToPrint.Append("Reg level")
            msgToPrint.Append("</td>")
            msgToPrint.Append("<td>")
            msgToPrint.Append(cboRegLevel.SelectedItem.Text)
            msgToPrint.Append("</td>")
            msgToPrint.Append("</tr>")
            'License type
            msgToPrint.Append("<tr>")
            msgToPrint.Append("<td width=150>")
            msgToPrint.Append("License type")
            msgToPrint.Append("</td>")
            msgToPrint.Append("<td>")
            msgToPrint.Append(cboLicenseType.SelectedItem.Text)
            msgToPrint.Append("</td>")
            msgToPrint.Append("</tr>")
            'Expire date / Afted days
            msgToPrint.Append("<tr>")
            msgToPrint.Append("<td width=150>")
            Select Case cboLicenseType.SelectedValue
                Case "0" 'time locked
                    msgToPrint.Append("Expire on date")
                Case "1" 'periodic
                    msgToPrint.Append("Expire after")
                Case "2" 'permanent
                    msgToPrint.Append("Not expire")
            End Select
            msgToPrint.Append("</td>")
            msgToPrint.Append("<td>")
            msgToPrint.Append(Server.HtmlEncode(txtDays.Text))
            msgToPrint.Append("</td>")
            msgToPrint.Append("</tr>")
            'UserName
            msgToPrint.Append("<tr>")
            msgToPrint.Append("<td width=150>")
            msgToPrint.Append("User Name")
            msgToPrint.Append("</td>")
            msgToPrint.Append("<td>")
            msgToPrint.Append(Server.HtmlEncode(txtUserName.Text))
            msgToPrint.Append("</td>")
            msgToPrint.Append("</tr>")
            'InstallCode
            msgToPrint.Append("<tr>")
            msgToPrint.Append("<td width=150>")
            msgToPrint.Append("Install code")
            msgToPrint.Append("</td>")
            msgToPrint.Append("<td>")
            msgToPrint.Append(Server.HtmlEncode(txtInstallCode.Text))
            msgToPrint.Append("</td>")
            msgToPrint.Append("</tr>")
            'License key
            msgToPrint.Append("<tr>")
            msgToPrint.Append("<td width=150>")
            msgToPrint.Append("License key")
            msgToPrint.Append("</td>")
            msgToPrint.Append("<td>")
            msgToPrint.Append(txtLicenseKey.Text.Replace(Chr(13), "</br>").Replace(Chr(10), ""))
            msgToPrint.Append("</td>")
            msgToPrint.Append("</tr>")
            'end table
            msgToPrint.Append("</table>")
            msgToPrint.Append("</br>")
            msgToPrint.Append("<p align='right'>")
            msgToPrint.Append("<a href='javascript:printPage();'>Print me</a>")
            msgToPrint.Append("&nbsp;")
            msgToPrint.Append("<a href='javascript:self.close()'>Close me</a>")
            msgToPrint.Append("</p>")
            msgToPrint.Append("</br>")
            'end html
            msgToPrint.Append("</body></html>")

            'write to page
            AjaxCallHelper.WriteLine("javascript:PrintLicenseKey(""" & msgToPrint.ToString & """, " & mWidth.ToString & ", " & mHeight.ToString & ");")
        End Sub
        Private Sub InitActiveLock()
            On Error GoTo InitForm_Error
            ActiveLock = ActiveLock3Globals_definst.NewInstance()
            ActiveLock.KeyStoreType = mKeyStoreType

            Dim MyAL As New Globals
            Dim MyGen As New AlugenGlobals

            'Use the following for ASP.NET applications
            'ActiveLock.Init(Application.StartupPath & "\bin")
            'Use the following for the VB.NET applications
            ActiveLock.Init(AppPath)

            ' Initialize Generator
            GeneratorInstance = MyGen.GeneratorInstance(mProductsStoreType)
            If File.Exists(mProductsStoragePath) = False Then
                Select Case mProductsStoreType
                    Case IActiveLock.ProductsStoreType.alsINIFile
                        mProductsStoragePath = AppPath() & "\licenses.ini"
                    Case IActiveLock.ProductsStoreType.alsMDBFile
                        mProductsStoragePath = AppPath() & "\licenses.mdb"
                    Case IActiveLock.ProductsStoreType.alsXMLFile
                        mProductsStoragePath = AppPath() & "\licenses.xml"
                End Select
            End If
            GeneratorInstance.StoragePath = mProductsStoragePath

            On Error GoTo 0
            Exit Sub

InitForm_Error:

            SayAjax("Error " & Err.Number & " (" & Err.Description & ") in procedure in Page Load")
        End Sub
        Private Sub cmdPaste_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdPaste.Click
            If txtInstallCode.Text.Length < 8 Then GoTo continueHere

            If txtInstallCode.Text.Substring(0, 8).ToLower = "you must" Then 'short key license
                Dim arrProdVer() As String
                arrProdVer = Split(txtInstallCode.Text, vbLf)
                systemEvent = True
                txtInstallCode.Text = (arrProdVer(1).Substring(15, 8)).Trim
                txtUserName.Text = (arrProdVer(3).Substring(11, arrProdVer(3).Length - 11)).Trim
                systemEvent = False
                HandleInstallationCode()
            Else
continueHere:
                'If Clipboard.GetDataObject.GetDataPresent(DataFormats.Text) Then
                '    txtInstallCode.Text = CType(Clipboard.GetDataObject.GetData(DataFormats.Text), String)
                UpdateKeyGenButtonStatus()
                HandleInstallationCode()
            End If

        End Sub

        '        Protected Sub cmdPasteInstallCode_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdPasteInstallCode.Init

        '            If txtInstallCode.Text.Length < 8 Then GoTo continueHere

        '            If txtInstallCode.Text.Substring(0, 8).ToLower = "you must" Then 'short key license
        '                Dim arrProdVer() As String
        '                arrProdVer = Split(txtInstallCode.Text, vbLf)
        '                systemEvent = True
        '                txtInstallCode.Text = (arrProdVer(1).Substring(15, 8)).Trim
        '                txtUserName.Text = (arrProdVer(3).Substring(11, arrProdVer(3).Length - 11)).Trim
        '                systemEvent = False
        '                HandleInstallationCode()
        '            Else
        'continueHere:
        '                'If Clipboard.GetDataObject.GetDataPresent(DataFormats.Text) Then
        '                '    txtInstallCode.Text = CType(Clipboard.GetDataObject.GetData(DataFormats.Text), String)
        '                UpdateKeyGenButtonStatus()
        '                HandleInstallationCode()
        '            End If
        '        End Sub

        Private Sub txtDays_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDays.TextChanged
            SaveFormSettings()
        End Sub

        Private Sub chkLockIP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockIP.CheckedChanged
            If systemEvent Then Exit Sub
            systemEvent = True
            txtInstallCode.Text = ReconstructedInstallationCode()
            systemEvent = False
            If chkLockIP.Checked Then
                MsgBox("Warning: Use Local IP addresses cautiously since they may not be static.", MsgBoxStyle.Exclamation, "Static IP Address Warning")
            End If
        End Sub
        Private Sub chkLockexternalIP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockExternalIP.CheckedChanged
            If systemEvent Then Exit Sub
            systemEvent = True
            txtInstallCode.Text = ReconstructedInstallationCode()
            systemEvent = False
            If chkLockExternalIP.Checked Then
                MsgBox("Warning: Use External IP addresses cautiously since they may not be static.", MsgBoxStyle.Exclamation, "Static IP Address Warning")
            End If
        End Sub

        Private Sub cmdCheckAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCheckAll.Click
            chkLockMACaddress.Checked = True
            chkLockComputer.Checked = True
            chkLockHD.Checked = True
            chkLockHDfirmware.Checked = True
            chkLockWindows.Checked = True
            chkLockBIOS.Checked = True
            chkLockMotherboard.Checked = True
            chkLockIP.Checked = True
            chkLockExternalIP.Checked = True
            chkLockFingerprint.Checked = True
            chkLockMemory.Checked = True
            chkLockCPUID.Checked = True
            chkLockBaseboardID.Checked = True
            chkLockVideoID.Checked = True
            SaveFormSettings()
        End Sub

        Private Sub cmdUncheckAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUncheckAll.Click
            chkLockMACaddress.Checked = False
            chkLockComputer.Checked = False
            chkLockHD.Checked = False
            chkLockHDfirmware.Checked = False
            chkLockWindows.Checked = False
            chkLockBIOS.Checked = False
            chkLockMotherboard.Checked = False
            chkLockIP.Checked = False
            chkLockExternalIP.Checked = False
            chkLockFingerprint.Checked = False
            chkLockMemory.Checked = False
            chkLockCPUID.Checked = False
            chkLockBaseboardID.Checked = False
            chkLockVideoID.Checked = False
            SaveFormSettings()
        End Sub

        Private Sub chkLockMACaddress_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockMACaddress.CheckedChanged
            If systemEvent Then Exit Sub
            systemEvent = True
            txtInstallCode.Text = ReconstructedInstallationCode()
            systemEvent = False
            SaveFormSettings()
        End Sub

        Private Sub chkLockComputer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockComputer.CheckedChanged
            If systemEvent Then Exit Sub
            systemEvent = True
            txtInstallCode.Text = ReconstructedInstallationCode()
            systemEvent = False
            SaveFormSettings()
        End Sub

        Private Sub chkLockHD_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockHD.CheckedChanged
            If systemEvent Then Exit Sub
            systemEvent = True
            txtInstallCode.Text = ReconstructedInstallationCode()
            systemEvent = False
            SaveFormSettings()
        End Sub

        Private Sub chkLockHDfirmware_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockHDfirmware.CheckedChanged
            If systemEvent Then Exit Sub
            systemEvent = True
            txtInstallCode.Text = ReconstructedInstallationCode()
            systemEvent = False
            SaveFormSettings()
        End Sub

        Private Sub chkLockWindows_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockWindows.CheckedChanged
            If systemEvent Then Exit Sub
            systemEvent = True
            txtInstallCode.Text = ReconstructedInstallationCode()
            systemEvent = False
            SaveFormSettings()
        End Sub

        Private Sub chkLockBIOS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockBIOS.CheckedChanged
            If systemEvent Then Exit Sub
            systemEvent = True
            txtInstallCode.Text = ReconstructedInstallationCode()
            systemEvent = False
            SaveFormSettings()
        End Sub

        Private Sub chkLockMotherboard_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockMotherboard.CheckedChanged
            If systemEvent Then Exit Sub
            systemEvent = True
            txtInstallCode.Text = ReconstructedInstallationCode()
            systemEvent = False
            SaveFormSettings()
        End Sub

        Private Sub chkLockVideoID_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockVideoID.CheckedChanged
            If systemEvent Then Exit Sub
            systemEvent = True
            txtInstallCode.Text = ReconstructedInstallationCode()
            systemEvent = False
            SaveFormSettings()
        End Sub

        Private Sub chkLockMemory_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockMemory.CheckedChanged
            If systemEvent Then Exit Sub
            systemEvent = True
            txtInstallCode.Text = ReconstructedInstallationCode()
            systemEvent = False
            SaveFormSettings()
        End Sub

        Private Sub chkLockCPUID_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockCPUID.CheckedChanged
            If systemEvent Then Exit Sub
            systemEvent = True
            txtInstallCode.Text = ReconstructedInstallationCode()
            systemEvent = False
            SaveFormSettings()
        End Sub

        Private Sub chkLockBaseboardID_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockBaseboardID.CheckedChanged
            If systemEvent Then Exit Sub
            systemEvent = True
            txtInstallCode.Text = ReconstructedInstallationCode()
            systemEvent = False
            SaveFormSettings()
        End Sub

        Private Sub chkLockFingerprint_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockFingerprint.CheckedChanged
            If systemEvent Then Exit Sub
            systemEvent = True
            txtInstallCode.Text = ReconstructedInstallationCode()
            systemEvent = False
            SaveFormSettings()
        End Sub

        Public Sub HandleInstallationCode()

            If systemEvent Then Exit Sub
            If txtInstallCode.Text.Length < 8 Then Exit Sub

            If txtInstallCode.Text.Substring(0, 8).ToLower = "you must" Then 'short key license
                Dim arrProdVer() As String
                arrProdVer = Split(txtInstallCode.Text, vbLf)
                systemEvent = True
                txtInstallCode.Text = (arrProdVer(1).Substring(15, 8)).Trim
                txtUserName.Text = (arrProdVer(3).Substring(11, arrProdVer(3).Length - 11)).Trim
                systemEvent = False
            End If

            If txtInstallCode.Text.Length = 8 Then 'Short key authorization is much simpler
                UpdateKeyGenButtonStatus()
                'If fDisableNotifications Then Exit Sub

                chkLockMACaddress.Visible = False
                chkLockComputer.Visible = False
                chkLockHD.Visible = False
                chkLockHDfirmware.Visible = False
                chkLockWindows.Visible = False
                chkLockBIOS.Visible = False
                chkLockMotherboard.Visible = False
                chkLockIP.Visible = False
                chkLockExternalIP.Visible = False
                chkLockFingerprint.Visible = False
                chkLockMemory.Visible = False
                chkLockCPUID.Visible = False
                chkLockBaseboardID.Visible = False
                chkLockVideoID.Visible = False

                'txtUserName.Text = ""
                txtUserName.Enabled = True
                txtUserName.ReadOnly = False
                txtUserName.BackColor = Color.White

                ' Label5.Visible = False
                'txtLicenseFile.Visible = False
                'cmdBrowse.Visible = False
                cmdSaveLicenseFile.Visible = False
                Exit Sub

            Else 'ALCrypto

                chkLockMACaddress.Visible = True
                chkLockComputer.Visible = True
                chkLockHD.Visible = True
                chkLockHDfirmware.Visible = True
                chkLockWindows.Visible = True
                chkLockBIOS.Visible = True
                chkLockMotherboard.Visible = True
                chkLockIP.Visible = True
                chkLockExternalIP.Visible = True
                chkLockFingerprint.Visible = True
                chkLockMemory.Visible = True
                chkLockCPUID.Visible = True
                chkLockBaseboardID.Visible = True
                chkLockVideoID.Visible = True
                txtUserName.Enabled = False
                txtUserName.ReadOnly = True
                txtUserName.BackColor = System.Drawing.SystemColors.Control

                'Label5.Visible = True
                'txtLicenseFile.Visible = True
                'cmdBrowse.Visible = True
                cmdSaveLicenseFile.Visible = True

                If Len(txtInstallCode.Text) > 0 Then
                    If systemEvent Then Exit Sub
                    UpdateKeyGenButtonStatus()
                    'If fDisableNotifications Then Exit Sub

                    'fDisableNotifications = True
                    txtUserName.Text = GetUserFromInstallCode(txtInstallCode.Text)
                    'fDisableNotifications = False

                    Dim installNameandVersion As String
                    Dim i As Integer, success As Boolean
                    installNameandVersion = GetUserSoftwareNameandVersionfromInstallCode(txtInstallCode.Text)
                    For i = 0 To cboProduct.Items.Count - 1
                        cboProduct.SelectedIndex = i
                        If installNameandVersion = cboProduct.Items(cboProduct.SelectedIndex).Text Then
                            success = True
                            Exit For
                        End If
                    Next i
                    If Not success Then
                        MsgBox("There's no matching Software Name and Version Number for this Installation Code.", MsgBoxStyle.Exclamation)
                    End If
                Else
                    'fDisableNotifications = True
                    chkLockComputer.Enabled = True
                    chkLockComputer.Text = "Lock to Computer Name"
                    chkLockHD.Enabled = True
                    chkLockHD.Text = "Lock to HDD Volume Serial"
                    chkLockHDfirmware.Enabled = True
                    chkLockHDfirmware.Text = "Lock to HDD Firmware Serial"
                    chkLockMACaddress.Enabled = True
                    chkLockMACaddress.Text = "Lock to MAC Address"
                    chkLockWindows.Enabled = True
                    chkLockWindows.Text = "Lock to Windows Serial"
                    chkLockBIOS.Enabled = True
                    chkLockBIOS.Text = "Lock to BIOS Version"
                    chkLockMotherboard.Enabled = True
                    chkLockMotherboard.Text = "Lock to Motherboard Serial"
                    chkLockIP.Enabled = True
                    chkLockIP.Text = "Lock to Local IP Address"
                    chkLockExternalIP.Enabled = True
                    chkLockExternalIP.Text = "Lock to External IP Address"
                    chkLockFingerprint.Enabled = True
                    chkLockFingerprint.Text = "Lock to Computer Fingerprint [VB.NET]"
                    chkLockMemory.Enabled = True
                    chkLockMemory.Text = "Lock to Memory"
                    chkLockCPUID.Enabled = True
                    chkLockCPUID.Text = "Lock to CPU ID"
                    chkLockBaseboardID.Enabled = True
                    chkLockBaseboardID.Text = "Lock to Baseboard ID"
                    chkLockVideoID.Enabled = True
                    chkLockVideoID.Text = "Lock to Video Controller ID"
                    txtUserName.Text = ""
                    'fDisableNotifications = False
                End If
            End If

        End Sub

        Private Sub txtName_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtProductName.TextChanged
            'fDisableNotifications = False
            UpdateCodeGenButtonStatus()
            UpdateAddButtonStatus()
        End Sub
        Private Sub txtVer_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtProductVersion.TextChanged
            ' New product - will be processed by Add command
            'fDisableNotifications = False
            UpdateCodeGenButtonStatus()
            UpdateAddButtonStatus()
        End Sub

        Private Sub chkNetworkedLicense_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkNetworkedLicense.CheckedChanged
            If chkNetworkedLicense.Checked = True Then
                lblNumberOfUsers.Visible = True
                txtMaxCount.Visible = True
            Else
                lblNumberOfUsers.Visible = False
                txtMaxCount.Visible = False
            End If
            SaveFormSettings()
        End Sub
        Private Sub cmdValidate2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdValidate2.Click

            ' ------------------ begin Message from Ismail ------------------
            ' This code block is used to Sign and then Validate any size text
            ' If you try to do the same with the typical RSA and with long strings, 
            ' routines will fail with a "Bad Length" error
            ' I am providing this sample here to show a second type of sign/verify scheme
            ' Although RSA is usually intended for signing/verifying short keys,
            ' the routine below with sign/verify any length string
            ' This 2nd type of validation button is hidden and is intended for 
            ' developers to test and learn.
            ' Note that no facility exists to retrieve the original data.
            ' Similar code can be found under SourceForge in project NCrypto
            ' ------------------ end Message from Ismail ------------------

            ' Get the current date format and save it to regionalSymbol variable
            Get_locale()
            ' Use this trick to temporarily set the date format to "yyyy/MM/dd"
            Set_locale("")

            If strLeft(txtVCode.Text, 3) = "RSA" Then

                Try
                    Dim strData As String
                    strData = "TestApp" & vbCrLf & "3" & vbCrLf & "Single" & vbCrLf & "1" & vbCrLf & "Evaluation User" & vbCrLf & "0" & vbCrLf & "2006/11/22" & vbCrLf & "2006/12/22" & vbCrLf & "5" & vbLf & "+00 10 18 09 71 85" & vbCrLf & "MYSWEETBABY" & vbCrLf & "5CA9-4B2A" & vbCrLf & "3JT26AA0" & vbCrLf & "55274-OEM-0011903-00102" & vbCrLf & "DELL   - 7" & vbCrLf & "BFWB741" & vbCrLf & "192.168.0.1"
                    Dim rsaCSP As New System.Security.Cryptography.RSACryptoServiceProvider
                    Dim rsaPubParams As RSAParameters 'stores public key
                    Dim strPublicBlob, strPrivateBlob As String

                    strPublicBlob = txtVCode.Text
                    strPrivateBlob = txtGCode.Text

                    If strLeft(txtGCode.Text, 6) = "RSA512" Then
                        strPrivateBlob = strRight(txtGCode.Text, Len(txtGCode.Text) - 6)
                    Else
                        strPrivateBlob = strRight(txtGCode.Text, Len(txtGCode.Text) - 7)
                    End If
                    ' import private key params into instance of RSACryptoServiceProvider
                    rsaCSP.FromXmlString(strPrivateBlob)
                    Dim rsaPrivateParams As RSAParameters 'stores private key
                    rsaPrivateParams = rsaCSP.ExportParameters(True)
                    rsaCSP.ImportParameters(rsaPrivateParams)

                    Dim userData As Byte() = Encoding.UTF8.GetBytes(strData)
                    Dim asf As AsymmetricSignatureFormatter = New RSAPKCS1SignatureFormatter(rsaCSP)
                    Dim algorithm As HashAlgorithm = New SHA1Managed
                    asf.SetHashAlgorithm(algorithm.ToString)
                    Dim myhashedData() As Byte ' a byte array to store hash value
                    Dim myhashedDataString As String
                    myhashedData = algorithm.ComputeHash(userData)
                    myhashedDataString = BitConverter.ToString(myhashedData).Replace("-", String.Empty)
                    Dim mysignature As Byte() ' holds signatures
                    mysignature = asf.CreateSignature(algorithm)
                    Dim mySignatureBlock As String
                    mySignatureBlock = Convert.ToBase64String(mysignature)

                    ' Verify Signature
                    If strLeft(txtVCode.Text, 6) = "RSA512" Then
                        strPublicBlob = strRight(txtVCode.Text, Len(txtVCode.Text) - 6)
                    Else
                        strPublicBlob = strRight(txtVCode.Text, Len(txtVCode.Text) - 7)
                    End If
                    rsaCSP.FromXmlString(strPublicBlob)
                    rsaPubParams = rsaCSP.ExportParameters(False)
                    ' import public key params into instance of RSACryptoServiceProvider
                    rsaCSP.ImportParameters(rsaPubParams)

                    ' Also could use the following to check if the string is a base64 string
                    If ExpBase64.IsMatch(mySignatureBlock) Then
                    End If

                    Dim newsignature() As Byte
                    newsignature = Convert.FromBase64String(mySignatureBlock)
                    Dim asd As AsymmetricSignatureDeformatter = New RSAPKCS1SignatureDeformatter(rsaCSP)
                    asd.SetHashAlgorithm(algorithm.ToString)
                    Dim newhashedData() As Byte ' a byte array to store hash value
                    Dim newhashedDataString As String
                    newhashedData = algorithm.ComputeHash(userData)
                    newhashedDataString = BitConverter.ToString(newhashedData).Replace("-", String.Empty)
                    Dim verified As Boolean
                    verified = asd.VerifySignature(algorithm, newsignature)
                    If verified Then
                        SayAjax(txtProductName.Text & " (" & txtProductVersion.Text & ") validated successfully.")
                        'MsgBox("Signature Valid", MsgBoxStyle.Information)
                    Else
                        SayAjax(txtProductName.Text & " (" & txtProductVersion.Text & ") GCode-VCode mismatch!")
                        'MsgBox("Invalid Signature", MsgBoxStyle.Exclamation)
                    End If

                    'Release any resources held by the RSA Service Provider
                    rsaCSP.Clear()
                    Set_locale(regionalSymbol)

                Catch ex As Exception
                    Set_locale(regionalSymbol)
                    SayAjax(ex.Message)
                End Try
            End If

        End Sub

    End Class

End Namespace

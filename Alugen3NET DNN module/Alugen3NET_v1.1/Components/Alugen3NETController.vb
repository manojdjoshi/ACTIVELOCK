Imports System
Imports System.Xml
Imports System.Data
Imports DotNetNuke
Imports DotNetNuke.Framework
Imports DotNetNuke.UI.Utilities
Imports DotNetNuke.Services.Exceptions
Imports FriendSoftware.DNN.Modules.Alugen3NET.Data

Namespace FriendSoftware.DNN.Modules.Alugen3NET.Business

  'main controller
  Public Class AlugenMainController
    Implements Entities.Modules.ISearchable
    Implements Entities.Modules.IPortable

#Region "Optional Interfaces"
    Public Function GetSearchItems(ByVal ModInfo As DotNetNuke.Entities.Modules.ModuleInfo) As DotNetNuke.Services.Search.SearchItemInfoCollection Implements DotNetNuke.Entities.Modules.ISearchable.GetSearchItems

    End Function

    Public Function ExportModule(ByVal ModuleID As Integer) As String Implements DotNetNuke.Entities.Modules.IPortable.ExportModule
      Dim strXML As String = ""
      'export products
      Dim objAlugenProductsController As New AlugenProductsController
      Dim arrProducts As ArrayList = objAlugenProductsController.List(GetPortalSettings().PortalId, ModuleID)
      strXML += "<Alugen3NET>"
      If arrProducts.Count <> 0 Then
        strXML += "<Alugen3NETProducts>"
        Dim objAlugenProductsInfo As AlugenProductsInfo
        For Each objAlugenProductsInfo In arrProducts
          strXML += "<Alugen3NETProduct>"
          strXML += "<ProductName>" & XmlConvert.EncodeName(objAlugenProductsInfo.ProductName) & "</ProductName>"
          strXML += "<ProductVersion>" & XmlConvert.EncodeName(objAlugenProductsInfo.ProductVersion) & "</ProductVersion>"
          strXML += "<ProductVCode>" & XmlConvert.EncodeName(objAlugenProductsInfo.ProductVCode) & "</ProductVCode>"
          strXML += "<ProductGCode>" & XmlConvert.EncodeName(objAlugenProductsInfo.ProductGCode) & "</ProductGCode>"
          strXML += "<ProductPrice>" & objAlugenProductsInfo.ProductPrice & "</ProductPrice>"
          strXML += "</Alugen3NETProduct>"
        Next
        strXML += "</Alugen3NETProducts>"
      End If
      'export customers
      Dim objAlugenCustomersController As New AlugenCustomersController
      Dim arrCustomers As ArrayList = objAlugenCustomersController.List(GetPortalSettings().PortalId, ModuleID)
      If arrCustomers.Count <> 0 Then
        strXML += "<Alugen3NETCustomers>"
        Dim objAlugenCustomersInfo As AlugenCustomersInfo
        For Each objAlugenCustomersInfo In arrCustomers
          strXML += "<Alugen3NETCustomer>"
          strXML += "<CustomerName>" & XmlConvert.EncodeName(objAlugenCustomersInfo.CustomerName) & "</CustomerName>"
          strXML += "<CustomerAddress>" & XmlConvert.EncodeName(objAlugenCustomersInfo.CustomerAddress) & "</CustomerAddress>"
          strXML += "<CustomerContactPerson>" & XmlConvert.EncodeName(objAlugenCustomersInfo.CustomerContactPerson) & "</CustomerContactPerson>"
          strXML += "<CustomerPhone>" & XmlConvert.EncodeName(objAlugenCustomersInfo.CustomerPhone) & "</CustomerPhone>"
          strXML += "<CustomerEmail>" & XmlConvert.EncodeName(objAlugenCustomersInfo.CustomerEmail) & "</CustomerEmail>"
          strXML += "<AssociatedUser>" & XmlConvert.EncodeName(objAlugenCustomersInfo.AssociatedUser) & "</AssociatedUser>"
          strXML += "<UseUserEmail>" & objAlugenCustomersInfo.UseUserEmail.ToString() & "</UseUserEmail>"
          strXML += "</Alugen3NETCustomer>"
        Next
        strXML += "</Alugen3NETCustomers>"
      End If
      'export codes
      Dim objAlugenCodeController As New AlugenCodeController
      Dim arrCodes As ArrayList = objAlugenCodeController.List(GetPortalSettings().PortalId, ModuleID)
      If arrCodes.Count <> 0 Then
        strXML += "<Alugen3NETCodes>"
        Dim objAlugenCodeInfo As AlugenCodeInfo
        For Each objAlugenCodeInfo In arrCodes
          strXML += "<Alugen3NETCode>"
          Dim strCustomerName As String = objAlugenCustomersController.Get(objAlugenCodeInfo.CustomerID).CustomerName
          strXML += "<CustomerName>" & XmlConvert.EncodeName(strCustomerName) & "</CustomerName>"
          Dim strProductName As String = objAlugenProductsController.Get(objAlugenCodeInfo.ProductID).ProductName
          strXML += "<ProductName>" & XmlConvert.EncodeName(strProductName) & "</ProductName>"
          Dim strProductVersion As String = objAlugenProductsController.Get(objAlugenCodeInfo.ProductID).ProductVersion
          strXML += "<ProductVersion>" & XmlConvert.EncodeName(strProductVersion) & "</ProductVersion>"
          strXML += "<CodeUserName>" & XmlConvert.EncodeName(objAlugenCodeInfo.CodeUserName) & "</CodeUserName>"
          strXML += "<CodeInstalationCode>" & XmlConvert.EncodeName(objAlugenCodeInfo.CodeInstalationCode) & "</CodeInstalationCode>"
          strXML += "<CodeActivationCode>" & XmlConvert.EncodeName(objAlugenCodeInfo.CodeActivationCode) & "</CodeActivationCode>"
          strXML += "<CodeLicenseType>" & objAlugenCodeInfo.CodeLicenseType & "</CodeLicenseType>"
          strXML += "<CodeExpireDate>" & objAlugenCodeInfo.CodeExpireDate.ToString("d") & "</CodeExpireDate>"
          strXML += "<CodeDays>" & objAlugenCodeInfo.CodeDays & "</CodeDays>"
          strXML += "<CodeRegisteredLevel>" & objAlugenCodeInfo.CodeRegisteredLevel & "</CodeRegisteredLevel>"
          strXML += "<GenByCustomer>" & objAlugenCodeInfo.GenByCustomer.ToString() & "</GenByCustomer>"
          strXML += "<CreatedDate>" & objAlugenCodeInfo.CreatedDate.ToString("d") & "</CreatedDate>"
          strXML += "</Alugen3NETCode>"
        Next
        strXML += "</Alugen3NETCodes>"
      End If
      strXML += "</Alugen3NET>"
      'return xml content
      Return strXML
    End Function

    Public Sub ImportModule(ByVal ModuleID As Integer, ByVal Content As String, ByVal Version As String, ByVal UserID As Integer) Implements DotNetNuke.Entities.Modules.IPortable.ImportModule
      'import products
      Dim objAlugenProductsController As New AlugenProductsController
      Dim xmlProduct As XmlNode

      Dim xmlProducts As XmlNode = GetContent(Content, "Alugen3NET/Alugen3NETProducts")
      For Each xmlProduct In xmlProducts.SelectNodes("Alugen3NETProduct")
        Dim objAlugenProductsInfo As New AlugenProductsInfo
        objAlugenProductsInfo.ModuleID = ModuleID
        objAlugenProductsInfo.PortalID = GetPortalSettings().PortalId
        'real data
        objAlugenProductsInfo.ProductName = XmlConvert.DecodeName(xmlProduct.Item("ProductName").InnerText)
        objAlugenProductsInfo.ProductVersion = XmlConvert.DecodeName(xmlProduct.Item("ProductVersion").InnerText)
        objAlugenProductsInfo.ProductVCode = XmlConvert.DecodeName(xmlProduct.Item("ProductVCode").InnerText)
        objAlugenProductsInfo.ProductGCode = XmlConvert.DecodeName(xmlProduct.Item("ProductGCode").InnerText)
        objAlugenProductsInfo.ProductPrice = CType(xmlProduct.Item("ProductPrice").InnerText, Decimal)

        objAlugenProductsController.Add(objAlugenProductsInfo)
      Next
      'import customers
      Dim objAlugenCustomersController As New AlugenCustomersController
      Dim xmlCustomer As XmlNode
      Dim xmlCustomers As XmlNode = GetContent(Content, "Alugen3NET/Alugen3NETCustomers")
      For Each xmlCustomer In xmlCustomers.SelectNodes("Alugen3NETCustomer")
        Dim objAlugenCustomersInfo As New AlugenCustomersInfo
        objAlugenCustomersInfo.ModuleID = ModuleID
        objAlugenCustomersInfo.PortalID = GetPortalSettings().PortalId
        'real info
        objAlugenCustomersInfo.CustomerName = XmlConvert.DecodeName(xmlCustomer.Item("CustomerName").InnerText)
        objAlugenCustomersInfo.CustomerAddress = XmlConvert.DecodeName(xmlCustomer.Item("CustomerAddress").InnerText)
        objAlugenCustomersInfo.CustomerContactPerson = XmlConvert.DecodeName(xmlCustomer.Item("CustomerContactPerson").InnerText)
        objAlugenCustomersInfo.CustomerPhone = XmlConvert.DecodeName(xmlCustomer.Item("CustomerPhone").InnerText)
        objAlugenCustomersInfo.CustomerEmail = XmlConvert.DecodeName(xmlCustomer.Item("CustomerEmail").InnerText)
        objAlugenCustomersInfo.AssociatedUser = XmlConvert.DecodeName(xmlCustomer.Item("AssociatedUser").InnerText)
        objAlugenCustomersInfo.UseUserEmail = CType(xmlCustomer.Item("UseUserEmail").InnerText, Boolean)

        objAlugenCustomersController.Add(objAlugenCustomersInfo)
      Next
      'import codes
      Dim objAlugenCodeController As New AlugenCodeController
      Dim xmlCode As XmlNode
      Dim xmlCodes As XmlNode = GetContent(Content, "Alugen3NET/Alugen3NETCodes")
      For Each xmlCode In xmlCodes.SelectNodes("Alugen3NETCode")
        Dim objAlugenCodeInfo As New AlugenCodeInfo
        objAlugenCodeInfo.ModuleID = ModuleID
        objAlugenCodeInfo.PortalID = GetPortalSettings().PortalId
        'real info
        Dim objAlugenCustomersInfo As AlugenCustomersInfo
        objAlugenCustomersInfo = objAlugenCustomersController.Get(XmlConvert.DecodeName(xmlCode.Item("CustomerName").InnerText), GetPortalSettings().PortalId, ModuleID)
        objAlugenCodeInfo.CustomerID = objAlugenCustomersInfo.CustomerID
        Dim objAlugenProductsInfo As New AlugenProductsInfo
        objAlugenProductsInfo = objAlugenProductsController.Get(XmlConvert.DecodeName(xmlCode.Item("ProductName").InnerText), XmlConvert.DecodeName(xmlCode.Item("ProductVersion").InnerText), GetPortalSettings().PortalId, ModuleID)
        objAlugenCodeInfo.ProductID = objAlugenProductsInfo.ProductID
        objAlugenCodeInfo.CodeUserName = XmlConvert.DecodeName(xmlCode.Item("CodeUserName").InnerText)
        objAlugenCodeInfo.CodeInstalationCode = XmlConvert.DecodeName(xmlCode.Item("CodeInstalationCode").InnerText)
        objAlugenCodeInfo.CodeActivationCode = XmlConvert.DecodeName(xmlCode.Item("CodeActivationCode").InnerText)
        objAlugenCodeInfo.CodeLicenseType = CType(xmlCode.Item("CodeLicenseType").InnerText, Integer)
        objAlugenCodeInfo.CodeExpireDate = CType(xmlCode.Item("CodeExpireDate").InnerText, Date)
        objAlugenCodeInfo.CodeDays = CType(xmlCode.Item("CodeDays").InnerText, Integer)
        objAlugenCodeInfo.CodeRegisteredLevel = CType(xmlCode.Item("CodeRegisteredLevel").InnerText, Integer)
        objAlugenCodeInfo.CreatedDate = CType(xmlCode.Item("CreatedDate").InnerText, Date)
        objAlugenCodeInfo.GenByCustomer = CType(xmlCode.Item("GenByCustomer").InnerText, Boolean)
        objAlugenCodeInfo.CreatedByUser = UserController.GetCurrentUserInfo.Username

        objAlugenCodeController.Add(objAlugenCodeInfo)
      Next
    End Sub

#End Region

  End Class

  'products controller
  Public Class AlugenProductsController
    Implements Entities.Modules.ISearchable
    Implements Entities.Modules.IPortable


#Region "Public Methods"
    Public Function [Get](ByVal productID As Integer) As AlugenProductsInfo

      Return CType(CBO.FillObject(DataProvider.Instance().Get_AlugenProducts(productID), GetType(AlugenProductsInfo)), AlugenProductsInfo)

    End Function
    Public Function [Get](ByVal productName As String, ByVal productVersion As String, ByVal portalID As Integer, ByVal moduleID As Integer) As AlugenProductsInfo

      Return CType(CBO.FillObject(DataProvider.Instance().Get_AlugenProducts(productName, productVersion, portalID, moduleID), GetType(AlugenProductsInfo)), AlugenProductsInfo)

    End Function

    Public Function List(ByVal portalID As Integer, ByVal moduleId As Integer) As ArrayList

      Return CBO.FillCollection(DataProvider.Instance().List_AlugenProducts(portalID, moduleId), GetType(AlugenProductsInfo))

    End Function

    Public Function ListFull(ByVal portalID As Integer, ByVal moduleId As Integer) As ArrayList

      Return CBO.FillCollection(DataProvider.Instance().ListFull_AlugenProducts(portalID, moduleId), GetType(AlugeProductsFullInfo))

    End Function

    Public Function Add(ByVal objfs_Alugen3NET As AlugenProductsInfo) As Integer

      Return CType(DataProvider.Instance().Add_AlugenProducts(objfs_Alugen3NET.PortalID, objfs_Alugen3NET.ModuleID, objfs_Alugen3NET.UserName, objfs_Alugen3NET.ProductName, objfs_Alugen3NET.ProductVersion, objfs_Alugen3NET.ProductVCode, objfs_Alugen3NET.ProductGCode, objfs_Alugen3NET.ProductPrice), Integer)

    End Function

    Public Sub Update(ByVal objfs_Alugen3NET As AlugenProductsInfo)

      DataProvider.Instance().Update_AlugenProducts(objfs_Alugen3NET.ProductID, objfs_Alugen3NET.PortalID, objfs_Alugen3NET.ModuleID, objfs_Alugen3NET.ProductName, objfs_Alugen3NET.ProductVersion, objfs_Alugen3NET.ProductVCode, objfs_Alugen3NET.ProductGCode, objfs_Alugen3NET.ProductPrice, objfs_Alugen3NET.UserName)

    End Sub

    Public Sub Delete(ByVal productID As Integer)

      DataProvider.Instance().Delete_AlugenProducts(productID)

    End Sub
#End Region

#Region "Optional Interfaces"
    Public Function GetSearchItems(ByVal ModInfo As Entities.Modules.ModuleInfo) As Services.Search.SearchItemInfoCollection Implements Entities.Modules.ISearchable.GetSearchItems

    End Function
    Public Function ExportModule(ByVal ModuleID As Integer) As String Implements Entities.Modules.IPortable.ExportModule

    End Function

    Public Sub ImportModule(ByVal ModuleID As Integer, ByVal Content As String, ByVal Version As String, ByVal UserId As Integer) Implements Entities.Modules.IPortable.ImportModule

    End Sub
#End Region

  End Class


  'customers controller
  Public Class AlugenCustomersController
    Implements Entities.Modules.ISearchable
    Implements Entities.Modules.IPortable

#Region "Public Methods"
    Public Function [Get](ByVal customerID As Integer) As AlugenCustomersInfo

      Return CType(CBO.FillObject(DataProvider.Instance().Getdnn_AlugenCustomers(customerID), GetType(AlugenCustomersInfo)), AlugenCustomersInfo)

    End Function
    Public Function [Get](ByVal customerName As String, ByVal portalID As Integer, ByVal moduleID As Integer) As AlugenCustomersInfo

      Return CType(CBO.FillObject(DataProvider.Instance().Getdnn_AlugenCustomers(customerName, portalID, moduleID), GetType(AlugenCustomersInfo)), AlugenCustomersInfo)

    End Function

    Public Function List(ByVal portalID As Integer, ByVal moduleId As Integer) As ArrayList

      Return CBO.FillCollection(DataProvider.Instance().Listdnn_AlugenCustomers(portalID, moduleId), GetType(AlugenCustomersInfo))

    End Function

    Public Function List(ByVal moduleId As Integer) As ArrayList

      Return CBO.FillCollection(DataProvider.Instance().Listdnn_AlugenCustomers(moduleId), GetType(AlugenCustomersInfo))

    End Function

    Public Function Add(ByVal objdnn_AlugenCustomers As AlugenCustomersInfo) As Integer

      Return CType(DataProvider.Instance().Adddnn_AlugenCustomers(objdnn_AlugenCustomers.PortalID, objdnn_AlugenCustomers.ModuleID, objdnn_AlugenCustomers.CustomerName, objdnn_AlugenCustomers.CustomerAddress, objdnn_AlugenCustomers.CustomerContactPerson, objdnn_AlugenCustomers.CustomerPhone, objdnn_AlugenCustomers.CustomerEmail, objdnn_AlugenCustomers.AssociatedUser, objdnn_AlugenCustomers.UseUserEmail), Integer)

    End Function

    Public Sub Update(ByVal objdnn_AlugenCustomers As AlugenCustomersInfo)

      DataProvider.Instance().Updatednn_AlugenCustomers(objdnn_AlugenCustomers.CustomerID, objdnn_AlugenCustomers.PortalID, objdnn_AlugenCustomers.ModuleID, objdnn_AlugenCustomers.CustomerName, objdnn_AlugenCustomers.CustomerAddress, objdnn_AlugenCustomers.CustomerContactPerson, objdnn_AlugenCustomers.CustomerPhone, objdnn_AlugenCustomers.CustomerEmail, objdnn_AlugenCustomers.AssociatedUser, objdnn_AlugenCustomers.UseUserEmail)

    End Sub

    Public Sub Delete(ByVal customerID As Integer)

      DataProvider.Instance().Deletednn_AlugenCustomers(customerID)

    End Sub
#End Region

#Region "Optional Interfaces"
    Public Function GetSearchItems(ByVal ModInfo As Entities.Modules.ModuleInfo) As Services.Search.SearchItemInfoCollection Implements Entities.Modules.ISearchable.GetSearchItems

    End Function
    Public Function ExportModule(ByVal ModuleID As Integer) As String Implements Entities.Modules.IPortable.ExportModule

    End Function

    Public Sub ImportModule(ByVal ModuleID As Integer, ByVal Content As String, ByVal Version As String, ByVal UserId As Integer) Implements Entities.Modules.IPortable.ImportModule

    End Sub
#End Region

  End Class


  'codes controller
  Public Class AlugenCodeController
    Implements Entities.Modules.ISearchable
    Implements Entities.Modules.IPortable

#Region "Public Methods"
    Public Function [Get](ByVal codeID As Integer) As AlugenCodeInfo

      Return CType(CBO.FillObject(DataProvider.Instance().Getdnn_AlugenCode(codeID), GetType(AlugenCodeInfo)), AlugenCodeInfo)

    End Function

    Public Function List() As ArrayList

      Return CBO.FillCollection(DataProvider.Instance().Listdnn_AlugenCode(), GetType(AlugenCodeInfo))

    End Function

    Public Function List(ByVal portalID As Integer, ByVal moduleId As Integer) As ArrayList

      Return CBO.FillCollection(DataProvider.Instance().Listdnn_AlugenCode(portalID, moduleId), GetType(AlugenCodeInfo))

    End Function

    Public Function ListByCustomerID(ByVal portalID As Integer, ByVal moduleId As Integer, ByVal customerID As Integer) As ArrayList

      Return CBO.FillCollection(DataProvider.Instance().Listdnn_AlugenCodeByCustomerID(portalID, moduleId, customerID), GetType(AlugenCodeInfo))

    End Function

    Public Function ListByProductID(ByVal portalID As Integer, ByVal moduleId As Integer, ByVal productID As Integer) As ArrayList

      Return CBO.FillCollection(DataProvider.Instance().Listdnn_AlugenCodeByProductID(portalID, moduleId, productID), GetType(AlugenCodeInfo))

    End Function

    Public Function ListFull(ByVal portalID As Integer, ByVal moduleId As Integer) As ArrayList

      Return CBO.FillCollection(DataProvider.Instance().ListFulldnn_AlugenCode(portalID, moduleId), GetType(AlugenCodeFullInfo))

    End Function

    Public Function ListFullForUser(ByVal forUser As String, ByVal portalID As Integer, ByVal moduleId As Integer) As ArrayList

      Return CBO.FillCollection(DataProvider.Instance().ListFullForUserdnn_AlugenCode(forUser, portalID, moduleId), GetType(AlugenCodeFullInfo))

    End Function

    Public Function Add(ByVal objdnn_AlugenCode As AlugenCodeInfo) As Integer

      Return CType(DataProvider.Instance().Adddnn_AlugenCode(objdnn_AlugenCode.PortalID, objdnn_AlugenCode.ModuleID, objdnn_AlugenCode.CustomerID, objdnn_AlugenCode.ProductID, objdnn_AlugenCode.CodeUserName, objdnn_AlugenCode.CodeInstalationCode, objdnn_AlugenCode.CodeActivationCode, objdnn_AlugenCode.CodeLicenseType, objdnn_AlugenCode.CodeExpireDate, objdnn_AlugenCode.CodeDays, objdnn_AlugenCode.CodeRegisteredLevel, objdnn_AlugenCode.GenByCustomer, objdnn_AlugenCode.CreatedByUser, objdnn_AlugenCode.CreatedDate), Integer)

    End Function

    Public Sub Update(ByVal objdnn_AlugenCode As AlugenCodeInfo)

      DataProvider.Instance().Updatednn_AlugenCode(objdnn_AlugenCode.CodeID, objdnn_AlugenCode.PortalID, objdnn_AlugenCode.ModuleID, objdnn_AlugenCode.CustomerID, objdnn_AlugenCode.ProductID, objdnn_AlugenCode.CodeUserName, objdnn_AlugenCode.CodeInstalationCode, objdnn_AlugenCode.CodeActivationCode, objdnn_AlugenCode.CodeLicenseType, objdnn_AlugenCode.CodeExpireDate, objdnn_AlugenCode.CodeDays, objdnn_AlugenCode.CodeRegisteredLevel, objdnn_AlugenCode.GenByCustomer, objdnn_AlugenCode.CreatedByUser, objdnn_AlugenCode.CreatedDate)

    End Sub

    Public Sub Delete(ByVal codeID As Integer)

      DataProvider.Instance().Deletednn_AlugenCode(codeID)

    End Sub
#End Region

#Region "Optional Interfaces"
    Public Function GetSearchItems(ByVal ModInfo As Entities.Modules.ModuleInfo) As Services.Search.SearchItemInfoCollection Implements Entities.Modules.ISearchable.GetSearchItems

    End Function
    Public Function ExportModule(ByVal ModuleID As Integer) As String Implements Entities.Modules.IPortable.ExportModule

    End Function

    Public Sub ImportModule(ByVal ModuleID As Integer, ByVal Content As String, ByVal Version As String, ByVal UserId As Integer) Implements Entities.Modules.IPortable.ImportModule

    End Sub
#End Region

  End Class

End Namespace

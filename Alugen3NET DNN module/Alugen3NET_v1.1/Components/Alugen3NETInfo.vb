Imports System
Imports System.Data


Namespace FriendSoftware.DNN.Modules.Alugen3NET.Business

  Public Class AlugenProductsInfo

#Region "Private Members"
    Dim _productID As Integer
    Dim _portalID As Integer
    Dim _moduleID As Integer
    Dim _productName As String
    Dim _productVersion As String
    Dim _productVCode As String
    Dim _productGCode As String
    Dim _productPrice As Decimal
    Dim _userName As String
#End Region

#Region "Constructors"
    Public Sub New()
    End Sub

    Public Sub New(ByVal productID As Integer, ByVal portalID As Integer, ByVal moduleID As Integer, ByVal productName As String, ByVal productVersion As String, ByVal productVCode As String, ByVal productGCode As String, ByVal productPrice As Decimal, ByVal userName As String)
      Me.ProductID = productID
      Me.PortalID = portalID
      Me.ModuleID = moduleID
      Me.ProductName = productName
      Me.ProductVersion = productVersion
      Me.ProductVCode = productVCode
      Me.ProductGCode = productGCode
      Me.ProductPrice = productPrice
      Me.UserName = userName
    End Sub
#End Region

#Region "Public Properties"
    Public Property ProductID() As Integer
      Get
        Return _productID
      End Get
      Set(ByVal Value As Integer)
        _productID = Value
      End Set
    End Property

    Public Property PortalID() As Integer
      Get
        Return _portalID
      End Get
      Set(ByVal Value As Integer)
        _portalID = Value
      End Set
    End Property

    Public Property ModuleID() As Integer
      Get
        Return _moduleID
      End Get
      Set(ByVal Value As Integer)
        _moduleID = Value
      End Set
    End Property

    Public Property ProductName() As String
      Get
        Return _productName
      End Get
      Set(ByVal Value As String)
        _productName = Value
      End Set
    End Property

    Public Property ProductVersion() As String
      Get
        Return _productVersion
      End Get
      Set(ByVal Value As String)
        _productVersion = Value
      End Set
    End Property

    Public Property ProductVCode() As String
      Get
        Return _productVCode
      End Get
      Set(ByVal Value As String)
        _productVCode = Value
      End Set
    End Property

    Public Property ProductGCode() As String
      Get
        Return _productGCode
      End Get
      Set(ByVal Value As String)
        _productGCode = Value
      End Set
    End Property

    Public Property ProductPrice() As Decimal
      Get
        Return _productPrice
      End Get
      Set(ByVal Value As Decimal)
        _productPrice = Value
      End Set
    End Property

    Public Property UserName() As String
      Get
        Return _userName
      End Get
      Set(ByVal Value As String)
        _userName = Value
      End Set
    End Property

#End Region

  End Class

  Public Class AlugeProductsFullInfo

#Region "Private Members"
    Dim _productID As Integer
    Dim _portalID As Integer
    Dim _moduleID As Integer
    Dim _productName As String
    Dim _productVersion As String
    Dim _productVCode As String
    Dim _productGCode As String
    Dim _productPrice As Decimal
    Dim _userName As String
    Dim _productNameVersion As String
#End Region

#Region "Constructors"
    Public Sub New()
    End Sub

    Public Sub New(ByVal productID As Integer, ByVal portalID As Integer, ByVal moduleID As Integer, ByVal productName As String, ByVal productVersion As String, ByVal productVCode As String, ByVal productGCode As String, ByVal productPrice As Decimal, ByVal userName As String, ByVal productNameVersion As String)
      Me.ProductID = productID
      Me.PortalID = portalID
      Me.ModuleID = moduleID
      Me.ProductName = productName
      Me.ProductVersion = productVersion
      Me.ProductVCode = productVCode
      Me.ProductGCode = productGCode
      Me.ProductPrice = productPrice
      Me.UserName = userName
      Me.ProductNameVersion = productNameVersion
    End Sub
#End Region

#Region "Public Properties"
    Public Property ProductNameVersion() As String
      Get
        Return _productNameVersion
      End Get
      Set(ByVal Value As String)
        _productNameVersion = Value
      End Set
    End Property
    Public Property ProductID() As Integer
      Get
        Return _productID
      End Get
      Set(ByVal Value As Integer)
        _productID = Value
      End Set
    End Property

    Public Property PortalID() As Integer
      Get
        Return _portalID
      End Get
      Set(ByVal Value As Integer)
        _portalID = Value
      End Set
    End Property

    Public Property ModuleID() As Integer
      Get
        Return _moduleID
      End Get
      Set(ByVal Value As Integer)
        _moduleID = Value
      End Set
    End Property

    Public Property ProductName() As String
      Get
        Return _productName
      End Get
      Set(ByVal Value As String)
        _productName = Value
      End Set
    End Property

    Public Property ProductVersion() As String
      Get
        Return _productVersion
      End Get
      Set(ByVal Value As String)
        _productVersion = Value
      End Set
    End Property

    Public Property ProductVCode() As String
      Get
        Return _productVCode
      End Get
      Set(ByVal Value As String)
        _productVCode = Value
      End Set
    End Property

    Public Property ProductGCode() As String
      Get
        Return _productGCode
      End Get
      Set(ByVal Value As String)
        _productGCode = Value
      End Set
    End Property

    Public Property ProductPrice() As Decimal
      Get
        Return _productPrice
      End Get
      Set(ByVal Value As Decimal)
        _productPrice = Value
      End Set
    End Property

    Public Property UserName() As String
      Get
        Return _userName
      End Get
      Set(ByVal Value As String)
        _userName = Value
      End Set
    End Property

#End Region

  End Class

  Public Class AlugenCustomersInfo

#Region "Private Members"
    Dim _customerID As Integer
    Dim _portalID As Integer
    Dim _moduleID As Integer
    Dim _customerName As String
    Dim _customerAddress As String
    Dim _customerContactPerson As String
    Dim _customerPhone As String
    Dim _customerEmail As String
    Dim _associatedUser As String
    Dim _useUserEmail As Boolean
#End Region

#Region "Constructors"
    Public Sub New()
    End Sub

    Public Sub New(ByVal customerID As Integer, ByVal portalID As Integer, ByVal moduleID As Integer, ByVal customerName As String, ByVal customerAddress As String, ByVal customerContactPerson As String, ByVal customerPhone As String, ByVal customerEmail As String, ByVal associatedUser As String, ByVal useUserEmail As Boolean)
      Me.CustomerID = customerID
      Me.PortalID = portalID
      Me.ModuleID = moduleID
      Me.CustomerName = customerName
      Me.CustomerAddress = customerAddress
      Me.CustomerContactPerson = customerContactPerson
      Me.CustomerPhone = customerPhone
      Me.CustomerEmail = customerEmail
      Me.AssociatedUser = associatedUser
    End Sub
#End Region

#Region "Public Properties"
    Public Property UseUserEmail() As Boolean
      Get
        Return _useUserEmail
      End Get
      Set(ByVal Value As Boolean)
        _useUserEmail = Value
      End Set
    End Property
    Public Property AssociatedUser() As String
      Get
        Return _associatedUser
      End Get
      Set(ByVal Value As String)
        _associatedUser = Value
      End Set
    End Property
    Public Property CustomerID() As Integer
      Get
        Return _customerID
      End Get
      Set(ByVal Value As Integer)
        _customerID = Value
      End Set
    End Property

    Public Property PortalID() As Integer
      Get
        Return _portalID
      End Get
      Set(ByVal Value As Integer)
        _portalID = Value
      End Set
    End Property

    Public Property ModuleID() As Integer
      Get
        Return _moduleID
      End Get
      Set(ByVal Value As Integer)
        _moduleID = Value
      End Set
    End Property

    Public Property CustomerName() As String
      Get
        Return _customerName
      End Get
      Set(ByVal Value As String)
        _customerName = Value
      End Set
    End Property

    Public Property CustomerAddress() As String
      Get
        Return _customerAddress
      End Get
      Set(ByVal Value As String)
        _customerAddress = Value
      End Set
    End Property

    Public Property CustomerContactPerson() As String
      Get
        Return _customerContactPerson
      End Get
      Set(ByVal Value As String)
        _customerContactPerson = Value
      End Set
    End Property

    Public Property CustomerPhone() As String
      Get
        Return _customerPhone
      End Get
      Set(ByVal Value As String)
        _customerPhone = Value
      End Set
    End Property

    Public Property CustomerEmail() As String
      Get
        Return _customerEmail
      End Get
      Set(ByVal Value As String)
        _customerEmail = Value
      End Set
    End Property
#End Region

  End Class


  Public Class AlugenCodeInfo

#Region "Private Members"
    Dim _codeID As Integer
    Dim _portalID As Integer
    Dim _moduleID As Integer
    Dim _customerID As Integer
    Dim _productID As Integer
    Dim _codeUserName As String
    Dim _codeInstalationCode As String
    Dim _codeActivationCode As String
    Dim _codeLicenseType As Integer
    Dim _codeExpireDate As DateTime
    Dim _codeDays As Integer
    Dim _codeRegisteredLevel As Integer
    Dim _codeGenByCustomer As Boolean
    Dim _createdByUser As String
    Dim _createdDate As DateTime
#End Region

#Region "Constructors"
    Public Sub New()
    End Sub

    Public Sub New(ByVal codeID As Integer, ByVal portalID As Integer, ByVal moduleID As Integer, ByVal customerID As Integer, ByVal productID As Integer, ByVal codeUserName As String, ByVal codeInstalationCode As String, ByVal codeActivationCode As String, ByVal codeLicenseType As Integer, ByVal codeExpireDate As DateTime, ByVal codeRegisteredLevel As Integer, ByVal codeGenByCustomer As Boolean, ByVal createdByUser As String, ByVal createdDate As DateTime)
      Me.CodeID = codeID
      Me.PortalID = portalID
      Me.ModuleID = moduleID
      Me.CustomerID = customerID
      Me.ProductID = productID
      Me.CodeUserName = codeUserName
      Me.CodeInstalationCode = codeInstalationCode
      Me.CodeActivationCode = codeActivationCode
      Me.CodeLicenseType = codeLicenseType
      Me.CodeExpireDate = codeExpireDate
      Me.CodeRegisteredLevel = codeRegisteredLevel
      Me.CreatedByUser = createdByUser
      Me.CreatedDate = createdDate
      Me.CodeDays = _codeDays
      Me.GenByCustomer = _codeGenByCustomer
    End Sub
#End Region

#Region "Public Properties"
    Public Property GenByCustomer() As Boolean
      Get
        Return _codeGenByCustomer
      End Get
      Set(ByVal Value As Boolean)
        _codeGenByCustomer = Value
      End Set
    End Property
    Public Property CodeDays() As Integer
      Get
        Return _codeDays
      End Get
      Set(ByVal Value As Integer)
        _codeDays = Value
      End Set
    End Property
    Public Property CodeID() As Integer
      Get
        Return _codeID
      End Get
      Set(ByVal Value As Integer)
        _codeID = Value
      End Set
    End Property

    Public Property PortalID() As Integer
      Get
        Return _portalID
      End Get
      Set(ByVal Value As Integer)
        _portalID = Value
      End Set
    End Property

    Public Property ModuleID() As Integer
      Get
        Return _moduleID
      End Get
      Set(ByVal Value As Integer)
        _moduleID = Value
      End Set
    End Property

    Public Property CustomerID() As Integer
      Get
        Return _customerID
      End Get
      Set(ByVal Value As Integer)
        _customerID = Value
      End Set
    End Property

    Public Property ProductID() As Integer
      Get
        Return _productID
      End Get
      Set(ByVal Value As Integer)
        _productID = Value
      End Set
    End Property

    Public Property CodeUserName() As String
      Get
        Return _codeUserName
      End Get
      Set(ByVal Value As String)
        _codeUserName = Value
      End Set
    End Property

    Public Property CodeInstalationCode() As String
      Get
        Return _codeInstalationCode
      End Get
      Set(ByVal Value As String)
        _codeInstalationCode = Value
      End Set
    End Property

    Public Property CodeActivationCode() As String
      Get
        Return _codeActivationCode
      End Get
      Set(ByVal Value As String)
        _codeActivationCode = Value
      End Set
    End Property

    Public Property CodeLicenseType() As Integer
      Get
        Return _codeLicenseType
      End Get
      Set(ByVal Value As Integer)
        _codeLicenseType = Value
      End Set
    End Property

    Public Property CodeExpireDate() As DateTime
      Get
        Return _codeExpireDate
      End Get
      Set(ByVal Value As DateTime)
        _codeExpireDate = Value
      End Set
    End Property

    Public Property CodeRegisteredLevel() As Integer
      Get
        Return _codeRegisteredLevel
      End Get
      Set(ByVal Value As Integer)
        _codeRegisteredLevel = Value
      End Set
    End Property

    Public Property CreatedByUser() As String
      Get
        Return _createdByUser
      End Get
      Set(ByVal Value As String)
        _createdByUser = Value
      End Set
    End Property

    Public Property CreatedDate() As DateTime
      Get
        Return _createdDate
      End Get
      Set(ByVal Value As DateTime)
        _createdDate = Value
      End Set
    End Property
#End Region

  End Class


  Public Class AlugenCodeFullInfo

#Region "Private Members"
    Dim _codeID As Integer
    Dim _portalID As Integer
    Dim _moduleID As Integer
    Dim _customerID As Integer
    Dim _productID As Integer
    Dim _codeUserName As String
    Dim _codeInstalationCode As String
    Dim _codeActivationCode As String
    Dim _codeLicenseType As Integer
    Dim _codeExpireDate As DateTime
    Dim _codeDays As Integer
    Dim _codeRegisteredLevel As Integer
    Dim _createdByUser As String
    Dim _createdDate As DateTime
    Dim _customerName As String
    Dim _productName As String
    Dim _productVersion As String
    Dim _codeGenByCustomer As Boolean
#End Region

#Region "Constructors"
    Public Sub New()
    End Sub

    Public Sub New(ByVal codeID As Integer, ByVal portalID As Integer, ByVal moduleID As Integer, ByVal customerID As Integer, ByVal productID As Integer, ByVal codeUserName As String, ByVal codeInstalationCode As String, ByVal codeActivationCode As String, ByVal codeLicenseType As Integer, ByVal codeExpireDate As DateTime, ByVal codeRegisteredLevel As Integer, ByVal createdByUser As String, ByVal createdDate As DateTime, ByVal customerName As String, ByVal productName As String, ByVal productVersion As String)
      Me.CodeID = codeID
      Me.PortalID = portalID
      Me.ModuleID = moduleID
      Me.CustomerID = customerID
      Me.ProductID = productID
      Me.CodeUserName = codeUserName
      Me.CodeInstalationCode = codeInstalationCode
      Me.CodeActivationCode = codeActivationCode
      Me.CodeLicenseType = codeLicenseType
      Me.CodeExpireDate = codeExpireDate
      Me.CodeRegisteredLevel = codeRegisteredLevel
      Me.CreatedByUser = createdByUser
      Me.CreatedDate = createdDate
      Me.CustomerName = customerName
      Me.ProductName = productName
      Me.ProductVersion = productVersion
      Me.CodeDays = _codeDays
      Me.GenByCustomer = _codeGenByCustomer
    End Sub
#End Region

#Region "Public Properties"
    Public Property GenByCustomer() As Boolean
      Get
        Return _codeGenByCustomer
      End Get
      Set(ByVal Value As Boolean)
        _codeGenByCustomer = Value
      End Set
    End Property
    Public Property CodeDays() As Integer
      Get
        Return _codeDays
      End Get
      Set(ByVal Value As Integer)
        _codeDays = Value
      End Set
    End Property
    Public Property ProductVersion() As String
      Get
        Return _productVersion
      End Get
      Set(ByVal Value As String)
        _productVersion = Value
      End Set
    End Property
    Public Property ProductName() As String
      Get
        Return _productName
      End Get
      Set(ByVal Value As String)
        _productName = Value
      End Set
    End Property
    Public Property CustomerName() As String
      Get
        Return _customerName
      End Get
      Set(ByVal Value As String)
        _customerName = Value
      End Set
    End Property
    Public Property CodeID() As Integer
      Get
        Return _codeID
      End Get
      Set(ByVal Value As Integer)
        _codeID = Value
      End Set
    End Property

    Public Property PortalID() As Integer
      Get
        Return _portalID
      End Get
      Set(ByVal Value As Integer)
        _portalID = Value
      End Set
    End Property

    Public Property ModuleID() As Integer
      Get
        Return _moduleID
      End Get
      Set(ByVal Value As Integer)
        _moduleID = Value
      End Set
    End Property

    Public Property CustomerID() As Integer
      Get
        Return _customerID
      End Get
      Set(ByVal Value As Integer)
        _customerID = Value
      End Set
    End Property

    Public Property ProductID() As Integer
      Get
        Return _productID
      End Get
      Set(ByVal Value As Integer)
        _productID = Value
      End Set
    End Property

    Public Property CodeUserName() As String
      Get
        Return _codeUserName
      End Get
      Set(ByVal Value As String)
        _codeUserName = Value
      End Set
    End Property

    Public Property CodeInstalationCode() As String
      Get
        Return _codeInstalationCode
      End Get
      Set(ByVal Value As String)
        _codeInstalationCode = Value
      End Set
    End Property

    Public Property CodeActivationCode() As String
      Get
        Return _codeActivationCode
      End Get
      Set(ByVal Value As String)
        _codeActivationCode = Value
      End Set
    End Property

    Public Property CodeLicenseType() As Integer
      Get
        Return _codeLicenseType
      End Get
      Set(ByVal Value As Integer)
        _codeLicenseType = Value
      End Set
    End Property

    Public Property CodeExpireDate() As DateTime
      Get
        Return _codeExpireDate
      End Get
      Set(ByVal Value As DateTime)
        _codeExpireDate = Value
      End Set
    End Property

    Public Property CodeRegisteredLevel() As Integer
      Get
        Return _codeRegisteredLevel
      End Get
      Set(ByVal Value As Integer)
        _codeRegisteredLevel = Value
      End Set
    End Property

    Public Property CreatedByUser() As String
      Get
        Return _createdByUser
      End Get
      Set(ByVal Value As String)
        _createdByUser = Value
      End Set
    End Property

    Public Property CreatedDate() As DateTime
      Get
        Return _createdDate
      End Get
      Set(ByVal Value As DateTime)
        _createdDate = Value
      End Set
    End Property
#End Region

  End Class
End Namespace

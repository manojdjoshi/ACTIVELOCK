Imports System
Imports System.Web.Caching
Imports System.Reflection

Namespace FriendSoftware.DNN.Modules.Alugen3NET.Data

    Public MustInherit Class DataProvider

#Region "Shared/Static Methods"
        ' singleton reference to the instantiated object 
        Private Shared objProvider As DataProvider = Nothing

        ' constructor
        Shared Sub New()
            CreateProvider()
        End Sub

        ' dynamically create provider
        Private Shared Sub CreateProvider()
            objProvider = CType(Framework.Reflection.CreateObject("data", "FriendSoftware.DNN.Modules.Alugen3NET.Data", "FriendSoftware.DNN.Modules.Alugen3NET"), DataProvider)
        End Sub

        ' return the provider
        Public Shared Shadows Function Instance() As DataProvider
            Return objProvider
        End Function
#End Region

#Region "AlugenProducts Methods"
    Public MustOverride Function Get_AlugenProducts(ByVal productID As Integer) As IDataReader
    Public MustOverride Function Get_AlugenProducts(ByVal productName As String, ByVal productVersion As String, ByVal portalID As Integer, ByVal moduleID As Integer) As IDataReader
    Public MustOverride Function List_AlugenProducts(ByVal portalID As Integer, ByVal moduleID As Integer) As IDataReader
    Public MustOverride Function ListFull_AlugenProducts(ByVal portalID As Integer, ByVal moduleID As Integer) As IDataReader
    Public MustOverride Function Add_AlugenProducts(ByVal portalID As Integer, ByVal moduleID As Integer, ByVal UserName As String, ByVal productName As String, ByVal productVersion As String, ByVal productVCode As String, ByVal productGCode As String, ByVal productPrice As Decimal) As Integer
    Public MustOverride Sub Update_AlugenProducts(ByVal productID As Integer, ByVal portalID As Integer, ByVal moduleID As Integer, ByVal productName As String, ByVal productVersion As String, ByVal productVCode As String, ByVal productGCode As String, ByVal productPrice As Decimal, ByVal UserName As String)
    Public MustOverride Sub Delete_AlugenProducts(ByVal productID As Integer)
#End Region

#Region "AlugenCustomers Methods"
    Public MustOverride Function Getdnn_AlugenCustomers(ByVal customerID As Integer) As IDataReader
    Public MustOverride Function Getdnn_AlugenCustomers(ByVal customerName As String, ByVal portalID As Integer, ByVal moduleID As Integer) As IDataReader
    Public MustOverride Function Listdnn_AlugenCustomers(ByVal portalID As Integer, ByVal moduleID As Integer) As IDataReader
    Public MustOverride Function Listdnn_AlugenCustomers(ByVal moduleID As Integer) As IDataReader
    Public MustOverride Function Adddnn_AlugenCustomers(ByVal portalID As Integer, ByVal moduleID As Integer, ByVal customerName As String, ByVal customerAddress As String, ByVal customerContactPerson As String, ByVal customerPhone As String, ByVal customerEmail As String, ByVal associatedUser As String, ByVal useUserEmail As Boolean) As Integer
    Public MustOverride Sub Updatednn_AlugenCustomers(ByVal customerID As Integer, ByVal portalID As Integer, ByVal moduleID As Integer, ByVal customerName As String, ByVal customerAddress As String, ByVal customerContactPerson As String, ByVal customerPhone As String, ByVal customerEmail As String, ByVal associatedUser As String, ByVal useUserEmail As Boolean)
    Public MustOverride Sub Deletednn_AlugenCustomers(ByVal customerID As Integer)
#End Region

#Region "AlugenCode Methods"
    Public MustOverride Function Getdnn_AlugenCode(ByVal codeID As Integer) As IDataReader
    Public MustOverride Function Listdnn_AlugenCode() As IDataReader
    Public MustOverride Function Listdnn_AlugenCode(ByVal portalID As Integer, ByVal moduleID As Integer) As IDataReader
    Public MustOverride Function Listdnn_AlugenCodeByCustomerID(ByVal portalID As Integer, ByVal moduleID As Integer, ByVal customerID As Integer) As IDataReader
    Public MustOverride Function Listdnn_AlugenCodeByProductID(ByVal portalID As Integer, ByVal moduleID As Integer, ByVal productID As Integer) As IDataReader
    Public MustOverride Function ListFulldnn_AlugenCode(ByVal portalID As Integer, ByVal moduleID As Integer) As IDataReader
    Public MustOverride Function ListFullForUserdnn_AlugenCode(ByVal forUser As String, ByVal portalID As Integer, ByVal moduleID As Integer) As IDataReader
    Public MustOverride Function Adddnn_AlugenCode(ByVal portalID As Integer, ByVal moduleID As Integer, ByVal customerID As Integer, ByVal productID As Integer, ByVal codeUserName As String, ByVal codeInstalationCode As String, ByVal codeActivationCode As String, ByVal codeLicenseType As Integer, ByVal codeExpireDate As DateTime, ByVal codeDays As Integer, ByVal codeRegisteredLevel As Integer, ByVal codeGenByCustomer As Boolean, ByVal createdByUser As String, ByVal createdDate As DateTime) As Integer
    Public MustOverride Sub Updatednn_AlugenCode(ByVal codeID As Integer, ByVal portalID As Integer, ByVal moduleID As Integer, ByVal customerID As Integer, ByVal productID As Integer, ByVal codeUserName As String, ByVal codeInstalationCode As String, ByVal codeActivationCode As String, ByVal codeLicenseType As Integer, ByVal codeExpireDate As DateTime, ByVal codeDays As Integer, ByVal codeRegisteredLevel As Integer, ByVal codeGenByCustomer As Boolean, ByVal createdByUser As String, ByVal createdDate As DateTime)
    Public MustOverride Sub Deletednn_AlugenCode(ByVal codeID As Integer)
#End Region


  End Class

End Namespace

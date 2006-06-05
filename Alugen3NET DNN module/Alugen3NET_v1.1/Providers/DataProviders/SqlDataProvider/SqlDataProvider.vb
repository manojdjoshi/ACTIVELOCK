Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.ApplicationBlocks.Data
Imports DotNetNuke

Namespace FriendSoftware.DNN.Modules.Alugen3NET.Data

    Public Class SqlDataProvider

        Inherits DataProvider

#Region "Private Members"
		Private Const ProviderType As String = "data"
		Private _providerConfiguration As Framework.Providers.ProviderConfiguration = Framework.Providers.ProviderConfiguration.GetProviderConfiguration(ProviderType)
		Private _connectionString As String
		Private _providerPath As String
		Private _objectQualifier As String
		Private _databaseOwner As String
#End Region

#Region "Constructors"
		Public Sub New()

			' Read the configuration specific information for this provider
			Dim objProvider As Framework.Providers.Provider = CType(_providerConfiguration.Providers(_providerConfiguration.DefaultProvider), Framework.Providers.Provider)

			' Read the attributes for this provider
			If objProvider.Attributes("connectionStringName") <> "" AndAlso _
			System.Configuration.ConfigurationSettings.AppSettings(objProvider.Attributes("connectionStringName")) <> "" Then
				_connectionString = System.Configuration.ConfigurationSettings.AppSettings(objProvider.Attributes("connectionStringName"))
			Else
				_connectionString = objProvider.Attributes("connectionString")
			End If

			_providerPath = objProvider.Attributes("providerPath")

			_objectQualifier = objProvider.Attributes("objectQualifier")
			If _objectQualifier <> "" And _objectQualifier.EndsWith("_") = False Then
				_objectQualifier += "_"
			End If

			_databaseOwner = objProvider.Attributes("databaseOwner")
			If _databaseOwner <> "" And _databaseOwner.EndsWith(".") = False Then
				_databaseOwner += "."
			End If

		End Sub
#End Region

#Region "Properties"
		Public ReadOnly Property ConnectionString() As String
			Get
				Return _connectionString
			End Get
		End Property

		Public ReadOnly Property ProviderPath() As String
			Get
				Return _providerPath
			End Get
		End Property

		Public ReadOnly Property ObjectQualifier() As String
			Get
				Return _objectQualifier
			End Get
		End Property

		Public ReadOnly Property DatabaseOwner() As String
			Get
				Return _databaseOwner
			End Get
		End Property
#End Region

#Region "General Public Methods"
		Private Function GetNull(ByVal Field As Object) As Object
			Return DotNetNuke.Common.Utilities.Null.GetNull(Field, DBNull.Value)
		End Function
#End Region



#Region "AlugenProducts Methods"
    Public Overloads Overrides Function Get_AlugenProducts(ByVal productID As Integer) As IDataReader
      Return CType(SqlHelper.ExecuteReader(ConnectionString, DatabaseOwner & ObjectQualifier & "AlugenProductsGet", productID), IDataReader)
    End Function
    Public Overloads Overrides Function Get_AlugenProducts(ByVal productName As String, ByVal productVersion As String, ByVal portalID As Integer, ByVal moduleID As Integer) As IDataReader
      Return CType(SqlHelper.ExecuteReader(ConnectionString, DatabaseOwner & ObjectQualifier & "AlugenProductsGetByNameVersion", productName, productVersion, portalID, moduleID), IDataReader)
    End Function

    Public Overloads Overrides Function List_AlugenProducts(ByVal portalID As Integer, ByVal moduleID As Integer) As IDataReader
      Return CType(SqlHelper.ExecuteReader(ConnectionString, DatabaseOwner & ObjectQualifier & "AlugenProductsList", portalID, moduleID), IDataReader)
    End Function

    Public Overloads Overrides Function ListFull_AlugenProducts(ByVal portalID As Integer, ByVal moduleID As Integer) As IDataReader
      Return CType(SqlHelper.ExecuteReader(ConnectionString, DatabaseOwner & ObjectQualifier & "AlugenProductsListFull", portalID, moduleID), IDataReader)
    End Function

    Public Overrides Function Add_AlugenProducts(ByVal portalID As Integer, ByVal moduleID As Integer, ByVal UserName As String, ByVal productName As String, ByVal productVersion As String, ByVal productVCode As String, ByVal productGCode As String, ByVal productPrice As Decimal) As Integer
      Return CType(SqlHelper.ExecuteScalar(ConnectionString, DatabaseOwner & ObjectQualifier & "AlugenProductsAdd", portalID, moduleID, UserName, productName, productVersion, productVCode, productGCode, productPrice), Integer)
    End Function

    Public Overrides Sub Update_AlugenProducts(ByVal productID As Integer, ByVal portalID As Integer, ByVal moduleID As Integer, ByVal productName As String, ByVal productVersion As String, ByVal productVCode As String, ByVal productGCode As String, ByVal productPrice As Decimal, ByVal UserName As String)
      SqlHelper.ExecuteNonQuery(ConnectionString, DatabaseOwner & ObjectQualifier & "AlugenProductsUpdate", productID, portalID, moduleID, productName, productVersion, productVCode, productGCode, productPrice, UserName)
    End Sub

    Public Overrides Sub Delete_AlugenProducts(ByVal productID As Integer)
      SqlHelper.ExecuteNonQuery(ConnectionString, DatabaseOwner & ObjectQualifier & "AlugenProductsDelete", productID)
    End Sub
#End Region

#Region "AlugenCustomers Methods"
    Public Overloads Overrides Function Getdnn_AlugenCustomers(ByVal customerID As Integer) As IDataReader
      Return CType(SqlHelper.ExecuteReader(ConnectionString, DatabaseOwner & ObjectQualifier & "AlugenCustomersGet", customerID), IDataReader)
    End Function

    Public Overloads Overrides Function Getdnn_AlugenCustomers(ByVal customerName As String, ByVal portalID As Integer, ByVal moduleID As Integer) As IDataReader
      Return CType(SqlHelper.ExecuteReader(ConnectionString, DatabaseOwner & ObjectQualifier & "AlugenCustomersGetByName", customerName, portalID, moduleID), IDataReader)
    End Function

    Public Overloads Overrides Function Listdnn_AlugenCustomers(ByVal portalID As Integer, ByVal moduleID As Integer) As IDataReader
      Return CType(SqlHelper.ExecuteReader(ConnectionString, DatabaseOwner & ObjectQualifier & "AlugenCustomersList", portalID, moduleID), IDataReader)
    End Function

    Public Overloads Overrides Function Listdnn_AlugenCustomers(ByVal moduleID As Integer) As IDataReader
      Return CType(SqlHelper.ExecuteReader(ConnectionString, DatabaseOwner & ObjectQualifier & "AlugenCustomersList", moduleID), IDataReader)
    End Function

    Public Overrides Function Adddnn_AlugenCustomers(ByVal portalID As Integer, ByVal moduleID As Integer, ByVal customerName As String, ByVal customerAddress As String, ByVal customerContactPerson As String, ByVal customerPhone As String, ByVal customerEmail As String, ByVal associatedUser As String, ByVal useUserEmail As Boolean) As Integer
      Return CType(SqlHelper.ExecuteScalar(ConnectionString, DatabaseOwner & ObjectQualifier & "AlugenCustomersAdd", portalID, moduleID, customerName, customerAddress, customerContactPerson, customerPhone, customerEmail, associatedUser, useUserEmail), Integer)
    End Function

    Public Overrides Sub Updatednn_AlugenCustomers(ByVal customerID As Integer, ByVal portalID As Integer, ByVal moduleID As Integer, ByVal customerName As String, ByVal customerAddress As String, ByVal customerContactPerson As String, ByVal customerPhone As String, ByVal customerEmail As String, ByVal associatedUser As String, ByVal useUserEmail As Boolean)
      SqlHelper.ExecuteNonQuery(ConnectionString, DatabaseOwner & ObjectQualifier & "AlugenCustomersUpdate", customerID, portalID, moduleID, customerName, customerAddress, customerContactPerson, customerPhone, customerEmail, associatedUser, useUserEmail)
    End Sub

    Public Overrides Sub Deletednn_AlugenCustomers(ByVal customerID As Integer)
      SqlHelper.ExecuteNonQuery(ConnectionString, DatabaseOwner & ObjectQualifier & "AlugenCustomersDelete", customerID)
    End Sub
#End Region

#Region "AlugenCode Methods"
    Public Overrides Function Getdnn_AlugenCode(ByVal codeID As Integer) As IDataReader
      Return CType(SqlHelper.ExecuteReader(ConnectionString, DatabaseOwner & ObjectQualifier & "AlugenCodeGet", codeID), IDataReader)
    End Function

    Public Overloads Overrides Function Listdnn_AlugenCode() As IDataReader
      Return CType(SqlHelper.ExecuteReader(ConnectionString, DatabaseOwner & ObjectQualifier & "AlugenCodeList"), IDataReader)
    End Function

    Public Overloads Overrides Function Listdnn_AlugenCode(ByVal portalID As Integer, ByVal moduleID As Integer) As IDataReader
      Return CType(SqlHelper.ExecuteReader(ConnectionString, DatabaseOwner & ObjectQualifier & "AlugenCodeList", portalID, moduleID), IDataReader)
    End Function

    Public Overloads Overrides Function Listdnn_AlugenCodeByCustomerID(ByVal portalID As Integer, ByVal moduleID As Integer, ByVal customerID As Integer) As IDataReader
      Return CType(SqlHelper.ExecuteReader(ConnectionString, DatabaseOwner & ObjectQualifier & "AlugenCodeListByCustomer", portalID, moduleID, customerID), IDataReader)
    End Function

    Public Overloads Overrides Function Listdnn_AlugenCodeByProductID(ByVal portalID As Integer, ByVal moduleID As Integer, ByVal productID As Integer) As IDataReader
      Return CType(SqlHelper.ExecuteReader(ConnectionString, DatabaseOwner & ObjectQualifier & "AlugenCodeListByProduct", portalID, moduleID, productID), IDataReader)
    End Function


    Public Overloads Overrides Function ListFulldnn_AlugenCode(ByVal portalID As Integer, ByVal moduleID As Integer) As IDataReader
      Return CType(SqlHelper.ExecuteReader(ConnectionString, DatabaseOwner & ObjectQualifier & "AlugenCodeListFull", portalID, moduleID), IDataReader)
    End Function

    Public Overloads Overrides Function ListFullForUserdnn_AlugenCode(ByVal forUser As String, ByVal portalID As Integer, ByVal moduleID As Integer) As IDataReader
      Return CType(SqlHelper.ExecuteReader(ConnectionString, DatabaseOwner & ObjectQualifier & "AlugenCodeListFullForUser", forUser, portalID, moduleID), IDataReader)
    End Function

    Public Overrides Function Adddnn_AlugenCode(ByVal portalID As Integer, ByVal moduleID As Integer, ByVal customerID As Integer, ByVal productID As Integer, ByVal codeUserName As String, ByVal codeInstalationCode As String, ByVal codeActivationCode As String, ByVal codeLicenseType As Integer, ByVal codeExpireDate As DateTime, ByVal codeDays As Integer, ByVal codeRegisteredLevel As Integer, ByVal codeGenByCustomer As Boolean, ByVal createdByUser As String, ByVal createdDate As DateTime) As Integer
      Return CType(SqlHelper.ExecuteScalar(ConnectionString, DatabaseOwner & ObjectQualifier & "AlugenCodeAdd", portalID, moduleID, customerID, productID, codeUserName, codeInstalationCode, codeActivationCode, codeLicenseType, codeExpireDate, codeDays, codeRegisteredLevel, codeGenByCustomer, GetNull(createdByUser), GetNull(createdDate)), Integer)
    End Function

    Public Overrides Sub Updatednn_AlugenCode(ByVal codeID As Integer, ByVal portalID As Integer, ByVal moduleID As Integer, ByVal customerID As Integer, ByVal productID As Integer, ByVal codeUserName As String, ByVal codeInstalationCode As String, ByVal codeActivationCode As String, ByVal codeLicenseType As Integer, ByVal codeExpireDate As DateTime, ByVal codeDays As Integer, ByVal codeRegisteredLevel As Integer, ByVal codeGenByCustomer As Boolean, ByVal createdByUser As String, ByVal createdDate As DateTime)
      SqlHelper.ExecuteNonQuery(ConnectionString, DatabaseOwner & ObjectQualifier & "AlugenCodeUpdate", codeID, portalID, moduleID, customerID, productID, codeUserName, codeInstalationCode, codeActivationCode, codeLicenseType, codeExpireDate, codeDays, codeRegisteredLevel, codeGenByCustomer, GetNull(createdByUser), GetNull(createdDate))
    End Sub

    Public Overrides Sub Deletednn_AlugenCode(ByVal codeID As Integer)
      SqlHelper.ExecuteNonQuery(ConnectionString, DatabaseOwner & ObjectQualifier & "AlugenCodeDelete", codeID)
    End Sub
#End Region

  End Class

End Namespace
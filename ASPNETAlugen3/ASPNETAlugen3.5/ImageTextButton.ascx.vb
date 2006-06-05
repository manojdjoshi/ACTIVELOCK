Public Class ImageTextButton
    Inherits System.Web.UI.UserControl

  Private _imageUrl As String
  Private _text As String

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

  End Sub
  Protected WithEvents Image1 As New System.Web.UI.WebControls.Image
  Protected WithEvents Label1 As New System.Web.UI.WebControls.Label

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here
    End Sub

  Public Sub New()
    _text = "[Enter some text here]"
  End Sub

  Public Property ImageUrl() As String
    Get
      Return _imageUrl
    End Get
    Set(ByVal Value As String)
      _imageUrl = Value
      Image1.ImageUrl = Value
    End Set
  End Property

  Public Property Text() As String
    Get
      Return Label1.Text
    End Get
    Set(ByVal Value As String)
      Label1.Text = Value
    End Set
  End Property
End Class

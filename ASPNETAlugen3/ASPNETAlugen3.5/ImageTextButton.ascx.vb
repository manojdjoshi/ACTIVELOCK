Namespace ASPNETAlugen3

Partial Class ImageTextButton
    Inherits System.Web.UI.UserControl

  Private _imageUrl As String

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

  End Sub


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
    '
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

End Namespace

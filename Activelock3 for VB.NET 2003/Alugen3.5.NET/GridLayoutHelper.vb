Imports System
Imports System.Windows.Forms

Public Class GridLayoutHelper
  Private _grid As DataGrid
  Private _tableStyle As DataGridTableStyle
  Private _userOverride As Boolean = False
  Private _inManualSize As Boolean = False
  Private _percentages As Decimal() = Nothing
  Private _minimumWidths As Integer() = Nothing

  Public Sub New(ByVal grid As DataGrid, ByVal tableStyle As DataGridTableStyle, ByVal percentages As Decimal(), ByVal minimumWidths As Integer())
    Me._grid = grid
    Me._tableStyle = tableStyle
    Me._percentages = percentages
    Me._minimumWidths = minimumWidths
    For Each col As DataGridColumnStyle In tableStyle.GridColumnStyles
      AddHandler col.WidthChanged, AddressOf Me.Column_WidthChanged
    Next
    AddHandler grid.Resize, AddressOf Me.Grid_SizeChanged
  End Sub

  Public Sub SizeTheGrid()
    If _tableStyle Is Nothing Then
      Return
    End If
    Dim colStyles As GridColumnStylesCollection = _tableStyle.GridColumnStyles
    If colStyles.Count = 0 Then
      Return
    End If
    _inManualSize = True
    Dim width As Integer = _grid.Size.Width - _grid.GetCellBounds(0, 0).X - 4
    If IsScrollBarVisible(_grid) Then
      width -= SystemInformation.VerticalScrollBarWidth
    End If
    Dim nCols As Integer = _tableStyle.GridColumnStyles.Count
    Dim lastColIndex As Integer = nCols - 1
    colStyles(lastColIndex).Width = 0
    Dim dWidth As Decimal = CType(width, Decimal)
    Dim totalWidth As Integer = width
    Dim colWidth As Integer
    colWidth = width / nCols
    Dim i As Integer = 0
    While i < lastColIndex
      If Not _percentages Is Nothing Then
        colWidth = CType((dWidth * _percentages(i)), Integer)
      End If
      If Not _minimumWidths Is Nothing Then
        colStyles(i).Width = Math.Max(colWidth, _minimumWidths(i))
        totalWidth -= colStyles(i).Width
      Else
        colStyles(i).Width = colWidth
        totalWidth -= colStyles(i).Width
      End If
      System.Threading.Interlocked.Increment(i)
    End While
    If Not _minimumWidths Is Nothing Then
      colStyles(lastColIndex).Width = Math.Max(totalWidth, _minimumWidths(lastColIndex))
    Else
      colStyles(lastColIndex).Width = totalWidth
    End If
    _inManualSize = False
  End Sub

  Public Sub FixupLastColumn()
    If _tableStyle Is Nothing Then
      Return
    End If
    Dim colStyles As GridColumnStylesCollection = _tableStyle.GridColumnStyles
    If colStyles.Count = 0 Then
      Return
    End If
    _inManualSize = True
    Dim width As Integer = _grid.Size.Width - _grid.GetCellBounds(0, 0).X - 4
    If IsScrollBarVisible(_grid) Then
      width -= SystemInformation.VerticalScrollBarWidth
    End If
    Dim nCols As Integer = _tableStyle.GridColumnStyles.Count
    Dim lastColIndex As Integer = nCols - 1
    Dim totalWidth As Integer = width
    Dim i As Integer = 0
    While i < lastColIndex
      totalWidth -= colStyles(i).Width
      System.Threading.Interlocked.Increment(i)
    End While
    Dim minSizeLastCol As Integer = -1
    If Not _minimumWidths Is Nothing Then
      minSizeLastCol = _minimumWidths(lastColIndex)
    End If
    If totalWidth > minSizeLastCol Then
      colStyles(lastColIndex).Width = totalWidth
    End If
    _inManualSize = False
  End Sub

  Protected Function IsScrollBarVisible(ByVal aControl As Control) As Boolean
    For Each c As Control In aControl.Controls
      If c.GetType.Equals(GetType(VScrollBar)) Then
        Return c.Visible
      End If
    Next
    Return False
  End Function

  Public ReadOnly Property Percentages() As Decimal()
    Get
      Return _percentages
    End Get
  End Property

  Public ReadOnly Property MinimumColumnWidths() As Integer()
    Get
      Return _minimumWidths
    End Get
  End Property

  Public Sub Grid_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Me.Handle_SizeChanged(_grid, _tableStyle)
  End Sub

  Public Sub Column_WidthChanged(ByVal sender As Object, ByVal e As EventArgs)
    Me.Handle_WidthChanged(_grid, _tableStyle)
  End Sub

  Public Sub Handle_SizeChanged(ByVal grid As DataGrid, ByVal ts As DataGridTableStyle)
    If Not _userOverride Then
      Me.SizeTheGrid()
    Else
      Me.FixupLastColumn()
    End If
  End Sub

  Public Sub Handle_WidthChanged(ByVal grid As DataGrid, ByVal ts As DataGridTableStyle)
    If Not _inManualSize Then
      _userOverride = True
      Me.FixupLastColumn()
    End If
  End Sub
End Class

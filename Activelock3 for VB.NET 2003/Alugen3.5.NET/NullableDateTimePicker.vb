' Copyright (c) 2005 Claudio Grazioli, http://www.grazioli.ch
'
' This code is free software; you can redistribute it and/or modify it.
' However, this header must remain intact and unchanged.  Additional
' information may be appended after this header.  Publications based on
' this code must also include an appropriate reference.
' 
' This code is distributed in the hope that it will be useful, but 
' WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY 
' or FITNESS FOR A PARTICULAR PURPOSE.
'
Option Strict On
Option Explicit On 

Imports System
Imports System.ComponentModel
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows.Forms

<ComVisible(False)> _
  Public Class NullableDateTimePicker
  Inherits System.Windows.Forms.DateTimePicker
  ' true, when no date shall be displayed (empty DateTimePicker)
  Private _isNull As Boolean '

  ' If _isNull = true, this value is shown in the DTP
  Private _nullValue As String

  ' The format of the DateTimePicker control
  Private _format As DateTimePickerFormat = DateTimePickerFormat.Long

  ' The custom format of the DateTimePicker control
  Private _customFormat As String

  ' The format of the DateTimePicker control as string
  Private _formatAsString As String

  Public Sub New()
    MyBase.Format = DateTimePickerFormat.Custom
    NullValue = " "
    Format = DateTimePickerFormat.Long
  End Sub

  Public Shadows Property Value() As [Object]
    Get
      If _isNull Then
        Return Nothing
      Else
        Return MyBase.Value
      End If
    End Get
    Set(ByVal Value As [Object])
      If Value Is Nothing Or Value Is DBNull.Value Then
        SetToNullValue()
      Else
        SetToDateTimeValue()
        MyBase.Value = CType(Value, DateTime)
      End If
    End Set
  End Property

  <Browsable(True), DefaultValue(DateTimePickerFormat.Long), TypeConverter(GetType([Enum]))> _
  Public Shadows Property Format() As DateTimePickerFormat
    Get
      Return _format
    End Get
    Set(ByVal Value As DateTimePickerFormat)
      _format = Value
      If Not _isNull Then
        SetFormat()
      End If
      OnFormatChanged(EventArgs.Empty)
    End Set
  End Property

  Public Shadows Property CustomFormat() As [String]
    Get
      Return _customFormat
    End Get
    Set(ByVal Value As [String])
      _customFormat = Value
    End Set
  End Property

  <Browsable(True), Category("Behavior"), Description("The string used to display null values in the control"), DefaultValue(" ")> _
  Public Property NullValue() As [String]
    Get
      Return _nullValue
    End Get
    Set(ByVal Value As [String])
      _nullValue = Value
    End Set
  End Property

  Private Property FormatAsString() As String
    Get
      Return _formatAsString
    End Get
    Set(ByVal Value As String)
      _formatAsString = Value
      MyBase.CustomFormat = Value
    End Set
  End Property

  Private Sub SetFormat()
    Dim ci As CultureInfo = Thread.CurrentThread.CurrentCulture
    Dim dtf As DateTimeFormatInfo = ci.DateTimeFormat
    Select Case _format
      Case DateTimePickerFormat.Long
        FormatAsString = dtf.LongDatePattern
      Case DateTimePickerFormat.Short
        FormatAsString = dtf.ShortDatePattern
      Case DateTimePickerFormat.Time
        FormatAsString = dtf.ShortTimePattern
      Case DateTimePickerFormat.Custom
        FormatAsString = Me.CustomFormat
    End Select
  End Sub

  Private Sub SetToNullValue()
    _isNull = True
    If Not (_nullValue Is Nothing Or _nullValue = [String].Empty) Then
      MyBase.CustomFormat = "'" + _nullValue + "'"
    Else
      MyBase.CustomFormat = " "
    End If
  End Sub

  Private Sub SetToDateTimeValue()
    If _isNull Then
      SetFormat()
      _isNull = False
      MyBase.OnValueChanged(New EventArgs)
    End If
  End Sub

  Protected Overrides Sub WndProc(ByRef m As Message)
    If _isNull Then
      If m.Msg = &H4E Then ' WM_NOTIFY
        Dim nm As NMHDR = CType(m.GetLParam(GetType(NMHDR)), NMHDR)
        If nm.Code = -746 Or nm.Code = -722 Then   ' DTN_CLOSEUP || DTN_?
          SetToDateTimeValue()
        End If
      End If
    End If
    MyBase.WndProc(m)
  End Sub

  <StructLayout(LayoutKind.Sequential)> _
  Private Structure NMHDR
    Public HwndFrom As IntPtr
    Public IdFrom As Integer
    Public Code As Integer
  End Structure

  Protected Overrides Sub OnKeyUp(ByVal e As KeyEventArgs)
    If e.KeyCode = Keys.Delete Then
      Me.Value = Nothing
      OnValueChanged(EventArgs.Empty)
    End If
    MyBase.OnKeyUp(e)
  End Sub

  Protected Overrides Sub OnValueChanged(ByVal eventargs As EventArgs)
    MyBase.OnValueChanged(eventargs)
  End Sub
End Class

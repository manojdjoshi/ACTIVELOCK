VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLBHScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' --- required API declarations ---

Private Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" _
   (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const LB_SETHORIZONTALEXTENT = &H194
Private Const WM_VSCROLL = &H115
Private Const SB_BOTTOM = 7

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXVSCROLL = 2

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" _
   (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, _
   lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_SINGLELINE = &H20
Private Const DT_CALCRECT = &H400

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function SelectObject Lib "gdi32" _
   (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
   (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_VSCROLL = &H200000

' --- private class variables ---
Private mvarListBox As ListBox
Private m_lMaxItemWidth As Long
Private m_hItemFont As Long
Private m_ListBoxHwnd As Long

Public Sub Init(ByRef pListBox As ListBox)
   Dim FontInt As IFont

   Set mvarListBox = pListBox
   mvarListBox.Clear
   m_lMaxItemWidth = 0
   m_ListBoxHwnd = mvarListBox.hwnd
   
   ' Determining the handle of the font used in the specified listbox.
   ' Using the IFont interface we can retreive the handle of the font
   ' used in the specified listbox.
   ' We'll use this handle further when we'll calculate the width of
   ' listbox items
   Set FontInt = pListBox.Font
   m_hItemFont = FontInt.hFont
End Sub

' The following routine adds a string to a specified list box
' and displays the horizontal scroll bar in it if required
Public Sub AddItem(ByRef psItemText As String)
   Dim m As Long
   Dim hdc As Long
   Dim rc As RECT
   Dim hOldFont As Long
   Dim bHasVScrBar As Boolean
   
   mvarListBox.AddItem psItemText
   
   ' --- calculating the width of the currently added item ---
   hdc = GetDC(m_ListBoxHwnd) ' retrieving HDC for the listbox
   hOldFont = SelectObject(hdc, m_hItemFont) ' selecting the required font
   ' if you specify the DT_CALCRECT flag,
   ' DrawText only Determines the width and height of the rectangle
   ' required to display the text:
   DrawText hdc, psItemText, -1, rc, DT_SINGLELINE + DT_CALCRECT
   m = rc.Right - rc.Left
   ' restoring the previous state
   Call SelectObject(hdc, hOldFont)
   ReleaseDC m_ListBoxHwnd, hdc
   
   ' --- determining whether we need to display the horizontal scroll bar ---
   If m > m_lMaxItemWidth Then
      m_lMaxItemWidth = m
      bHasVScrBar = GetWindowLong(m_ListBoxHwnd, GWL_STYLE) And WS_VSCROLL
      SendMessageByLong m_ListBoxHwnd, LB_SETHORIZONTALEXTENT, _
         m + IIf(bHasVScrBar, GetSystemMetrics(SM_CXVSCROLL), 4), 0
   End If
   
   ' --- scrolling the listbox to be sure that the user see the last item ---
   SendMessageByLong m_ListBoxHwnd, WM_VSCROLL, SB_BOTTOM, 0
End Sub

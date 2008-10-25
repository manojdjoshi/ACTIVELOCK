Option Strict Off
Option Explicit On
Friend Class frmC
	Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try 
					'For the start-up form, the first instance created is the default instance.
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmC))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.Timer1 = New System.Windows.Forms.Timer(components)
		Me.Text = "Form1"
		Me.ClientSize = New System.Drawing.Size(226, 216)
		Me.Location = New System.Drawing.Point(4, 30)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmC"
		Me.Timer1.Interval = 5000
		Me.Timer1.Enabled = True
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmC
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmC
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmC()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	'*   ActiveLock
	'*   Copyright 1998-2002 Nelson Ferraz
	'*   Copyright 2003-2006 The ActiveLock Software Group (ASG)
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
	'===============================================================================
	' Name: frmC
	' Purpose: This is the form that detects debuggers, hex editors and
	' file monitoring applications
	' Remarks: PROTECTION TECHNIQUES:
	'------------------------------------------------------------
	'- "SoftICE" (via VxD, window name and process name)
	'------------------------------------------------------------
	'- "Win32Dasm"
	'- "Debuggy By Vanja Fuckar" - Debuggy.exe
	'- "OllyDBG" - OLLYDBG.exe
	'- "ProcDump by G-Rom, Lorian & Stone" - PROCDUMP.exe
	'- "SoftSnoop by Yoda/f2f" - SoftSnoop.exe
	'- "TimeFix by GodsJiva" - TimeFix.exe
	'- "TMR Ripper Studio" - "TMG Ripper Studio.exe"
	'------------------------------------------------------------
	'- Regmon
	'- Filemon
	'------------------------------------------------------------
	'- Protection against step-debugging. (TimeFix-kill used before!!)
	'------------------------------------------------------------
	'- Processor handling routines are based on the processor handling
	'  routines written by Deniz Mert Edincik.
	'- SoftICE detection using VxD written by JooX (and/or) David Eriksson
	'- RegMon & FileMon detection by Detonate
	'- String encoding, window title detection, process detection and
	'  process destruction force-crash routines by Method.
	' Functions:
	' Properties:
	' Methods:
	' Started: 21.04.2005
	' Modified: 08.05.2005
	'===============================================================================
	' @author: activelock-admins
    ' @version: 3.3.0
    ' @date: 03.24.2005

	'===============================================================================
	' Name: Sub Form_Load
	' Input: None
	' Output: None
	' Purpose: Load event of the form detects debuggers, hex editors and
	' file monitoring applications
	' Remarks: GetSystemTime is the main call
	'===============================================================================
	Private Sub frmC_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		HAD2HAMMER = False
		
		wX = 0 : wY = 0
		GetSystemTime()
	End Sub
    '===============================================================================
	' Name: Sub Form_Resize
	' Input: None
	' Output: None
	' Purpose: If the height of this form becomes zero: BINGO
	' Remarks: None
	'===============================================================================
	'UPGRADE_WARNING: Event frmC.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
	Public Sub frmC_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		If wY = 0 Then
			HAD2HAMMER = True
			Err.Raise(0)
		End If
	End Sub
    '===============================================================================
	' Name: Sub Timer1_Timer
	' Input: None
	' Output: None
	' Purpose: Keeps shrinking the size of the form until it dies
	' Remarks: None
	'===============================================================================
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
        Me.SetBounds(Me.Left, Me.Top, Microsoft.VisualBasic.Compatibility.VB6.TwipsToPixelsX(wX), Microsoft.VisualBasic.Compatibility.VB6.TwipsToPixelsY(wY))
    End Sub
End Class
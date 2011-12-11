/***********************************************************************
ActiveLock
Copyright 1998-2002 Nelson Ferraz
Copyright 2003-2006 The ActiveLock Software Group (ASG)
All material is the property of the contributing authors.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are
met:

  [o] Redistributions of source code must retain the above copyright
      notice, this list of conditions and the following disclaimer.

  [o] Redistributions in binary form must reproduce the above
      copyright notice, this list of conditions and the following
      disclaimer in the documentation and/or other materials provided
      with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
"AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

***********************************************************************/
using System;
using System.Windows.Forms;
using System.IO;
using System.Text;
using System.Threading;
using System.Runtime.InteropServices;
using ActiveLock37Net;

namespace ALTestApp36NET_CS
{

	/// <summary>
	/// Summary description for frmMain.
	/// </summary>
	internal class frmMain : System.Windows.Forms.Form
	{

        [DllImport("kernel32.dll")]
        static extern uint GetSystemDirectory([Out] StringBuilder lpBuffer, uint uSize);
        [DllImport("shell32.dll")]
        static extern bool SHGetSpecialFolderPath(IntPtr hwndOwner, [Out] StringBuilder lpszPath, int nFolder, bool fCreate);
        
        private System.Windows.Forms.TabControl SStab1;
		private System.Windows.Forms.TabPage _SSTab1_TabPage0;
		private System.Windows.Forms.TabPage _SSTab1_TabPage1;
		private System.Windows.Forms.GroupBox fraRegStatus;
		private System.Windows.Forms.Button cmdResetTrial;
		private System.Windows.Forms.Button cmdKillTrial;
		private System.Windows.Forms.PictureBox Picture2;
		private System.Windows.Forms.Label Label9;
		private System.Windows.Forms.Label Label16;
		private System.Windows.Forms.Label Label5;
		private System.Windows.Forms.Label Label3;
		private System.Windows.Forms.Label Label2;
		private System.Windows.Forms.Label Label1;
		private System.Windows.Forms.Label Label8;
		private System.Windows.Forms.Label Label7;
		private System.Windows.Forms.Label Label6;
		private System.Windows.Forms.GroupBox fraReg;
		private System.Windows.Forms.Button cmdPaste;
		private System.Windows.Forms.Button cmdCopy;
		private System.Windows.Forms.Button cmdKillLicense;
		private System.Windows.Forms.Button cmdReqGen;
		private System.Windows.Forms.Button cmdRegister;
		private System.Windows.Forms.Label Label13;
		private System.Windows.Forms.Label Label11;
		private System.Windows.Forms.Label Label4;
		private System.Windows.Forms.Label lblLockStatus;
		private System.Windows.Forms.Label lblLockStatus2;
		private System.Windows.Forms.Label lblTrialInfo;
		private System.Windows.Forms.Panel Frame1;
		private System.Windows.Forms.Label lblHost;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;


		#region "Private members"

		//wrong way
		private _IActiveLock MyActiveLock; 
		//right way
		//private ActiveLock MyActiveLock; 

		private ActiveLockEventNotifier ActiveLockEventSink; 

        // Trial mode variables
		private bool noTrialThisTime; // ialkan - needed for registration while form was loaded via trial
        //private bool expireTrialLicense;
        private string strKeyStorePath;
        private string strAutoRegisterKeyPath;

        // Application name used
        private const string LICENSE_ROOT = "TestApp";

        // Timer count to check the license
        // private long timerCount		
		
		private System.Windows.Forms.RadioButton optForm1;
		private System.Windows.Forms.RadioButton optForm0;
		private System.Windows.Forms.TextBox txtVersion;
		private System.Windows.Forms.TextBox txtName;
		private System.Windows.Forms.TextBox txtRegStatus;
		private System.Windows.Forms.TextBox txtUsedDays;
		private System.Windows.Forms.TextBox txtExpiration;
		private System.Windows.Forms.TextBox txtRegisteredLevel;
		private System.Windows.Forms.TextBox txtChecksum;
		private System.Windows.Forms.TextBox txtLicenseType;
		private System.Windows.Forms.TextBox txtUser;
		private System.Windows.Forms.TextBox txtReqCodeGen;
		private System.Windows.Forms.TextBox txtLibKeyIn;
		private System.Windows.Forms.CheckBox chkScroll;
		private System.Windows.Forms.CheckBox chkPause;
		private System.Windows.Forms.CheckBox chkFlash;
		private System.Windows.Forms.ComboBox cboSpeed;
		private System.Windows.Forms.Label lblSpeed; 		 
		
		#endregion

		public frmMain()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//

			ActiveLockEventSink = new ActiveLockEventNotifier();
			ActiveLockEventSink.ValidateValue += new ActiveLockEventNotifier.ValidateValueEventHandler(ActiveLockEventSink_ValidateValue);
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.SStab1 = new System.Windows.Forms.TabControl();
            this._SSTab1_TabPage0 = new System.Windows.Forms.TabPage();
            this.fraRegStatus = new System.Windows.Forms.GroupBox();
            this.txtLicenseType = new System.Windows.Forms.TextBox();
            this.cmdResetTrial = new System.Windows.Forms.Button();
            this.cmdKillTrial = new System.Windows.Forms.Button();
            this.Picture2 = new System.Windows.Forms.PictureBox();
            this.txtRegisteredLevel = new System.Windows.Forms.TextBox();
            this.txtChecksum = new System.Windows.Forms.TextBox();
            this.txtVersion = new System.Windows.Forms.TextBox();
            this.txtName = new System.Windows.Forms.TextBox();
            this.txtExpiration = new System.Windows.Forms.TextBox();
            this.txtUsedDays = new System.Windows.Forms.TextBox();
            this.txtRegStatus = new System.Windows.Forms.TextBox();
            this.Label9 = new System.Windows.Forms.Label();
            this.Label16 = new System.Windows.Forms.Label();
            this.Label5 = new System.Windows.Forms.Label();
            this.Label3 = new System.Windows.Forms.Label();
            this.Label2 = new System.Windows.Forms.Label();
            this.Label1 = new System.Windows.Forms.Label();
            this.Label8 = new System.Windows.Forms.Label();
            this.Label7 = new System.Windows.Forms.Label();
            this.Label6 = new System.Windows.Forms.Label();
            this.fraReg = new System.Windows.Forms.GroupBox();
            this.cmdPaste = new System.Windows.Forms.Button();
            this.cmdCopy = new System.Windows.Forms.Button();
            this.cmdKillLicense = new System.Windows.Forms.Button();
            this.txtUser = new System.Windows.Forms.TextBox();
            this.cmdReqGen = new System.Windows.Forms.Button();
            this.txtReqCodeGen = new System.Windows.Forms.TextBox();
            this.cmdRegister = new System.Windows.Forms.Button();
            this.txtLibKeyIn = new System.Windows.Forms.TextBox();
            this.Label13 = new System.Windows.Forms.Label();
            this.Label11 = new System.Windows.Forms.Label();
            this.Label4 = new System.Windows.Forms.Label();
            this._SSTab1_TabPage1 = new System.Windows.Forms.TabPage();
            this.lblLockStatus = new System.Windows.Forms.Label();
            this.lblLockStatus2 = new System.Windows.Forms.Label();
            this.lblTrialInfo = new System.Windows.Forms.Label();
            this.Frame1 = new System.Windows.Forms.Panel();
            this.optForm1 = new System.Windows.Forms.RadioButton();
            this.optForm0 = new System.Windows.Forms.RadioButton();
            this.cboSpeed = new System.Windows.Forms.ComboBox();
            this.chkPause = new System.Windows.Forms.CheckBox();
            this.chkFlash = new System.Windows.Forms.CheckBox();
            this.chkScroll = new System.Windows.Forms.CheckBox();
            this.lblHost = new System.Windows.Forms.Label();
            this.lblSpeed = new System.Windows.Forms.Label();
            this.SStab1.SuspendLayout();
            this._SSTab1_TabPage0.SuspendLayout();
            this.fraRegStatus.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Picture2)).BeginInit();
            this.fraReg.SuspendLayout();
            this._SSTab1_TabPage1.SuspendLayout();
            this.Frame1.SuspendLayout();
            this.SuspendLayout();
            // 
            // SStab1
            // 
            this.SStab1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.SStab1.Controls.Add(this._SSTab1_TabPage0);
            this.SStab1.Controls.Add(this._SSTab1_TabPage1);
            this.SStab1.Location = new System.Drawing.Point(8, 8);
            this.SStab1.Name = "SStab1";
            this.SStab1.SelectedIndex = 0;
            this.SStab1.Size = new System.Drawing.Size(648, 496);
            this.SStab1.TabIndex = 0;
            // 
            // _SSTab1_TabPage0
            // 
            this._SSTab1_TabPage0.Controls.Add(this.fraRegStatus);
            this._SSTab1_TabPage0.Controls.Add(this.fraReg);
            this._SSTab1_TabPage0.ForeColor = System.Drawing.Color.Blue;
            this._SSTab1_TabPage0.Location = new System.Drawing.Point(4, 22);
            this._SSTab1_TabPage0.Name = "_SSTab1_TabPage0";
            this._SSTab1_TabPage0.Size = new System.Drawing.Size(640, 470);
            this._SSTab1_TabPage0.TabIndex = 0;
            this._SSTab1_TabPage0.Text = "Registration";
            // 
            // fraRegStatus
            // 
            this.fraRegStatus.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.fraRegStatus.BackColor = System.Drawing.SystemColors.Control;
            this.fraRegStatus.Controls.Add(this.txtLicenseType);
            this.fraRegStatus.Controls.Add(this.cmdResetTrial);
            this.fraRegStatus.Controls.Add(this.cmdKillTrial);
            this.fraRegStatus.Controls.Add(this.Picture2);
            this.fraRegStatus.Controls.Add(this.txtRegisteredLevel);
            this.fraRegStatus.Controls.Add(this.txtChecksum);
            this.fraRegStatus.Controls.Add(this.txtVersion);
            this.fraRegStatus.Controls.Add(this.txtName);
            this.fraRegStatus.Controls.Add(this.txtExpiration);
            this.fraRegStatus.Controls.Add(this.txtUsedDays);
            this.fraRegStatus.Controls.Add(this.txtRegStatus);
            this.fraRegStatus.Controls.Add(this.Label9);
            this.fraRegStatus.Controls.Add(this.Label16);
            this.fraRegStatus.Controls.Add(this.Label5);
            this.fraRegStatus.Controls.Add(this.Label3);
            this.fraRegStatus.Controls.Add(this.Label2);
            this.fraRegStatus.Controls.Add(this.Label1);
            this.fraRegStatus.Controls.Add(this.Label8);
            this.fraRegStatus.Controls.Add(this.Label7);
            this.fraRegStatus.Controls.Add(this.Label6);
            this.fraRegStatus.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.fraRegStatus.ForeColor = System.Drawing.Color.Blue;
            this.fraRegStatus.Location = new System.Drawing.Point(0, 0);
            this.fraRegStatus.Name = "fraRegStatus";
            this.fraRegStatus.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.fraRegStatus.Size = new System.Drawing.Size(633, 177);
            this.fraRegStatus.TabIndex = 18;
            this.fraRegStatus.TabStop = false;
            this.fraRegStatus.Text = "Status";
            // 
            // txtLicenseType
            // 
            this.txtLicenseType.AcceptsReturn = true;
            this.txtLicenseType.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtLicenseType.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.txtLicenseType.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.txtLicenseType.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtLicenseType.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtLicenseType.Location = new System.Drawing.Point(276, 136);
            this.txtLicenseType.MaxLength = 0;
            this.txtLicenseType.Name = "txtLicenseType";
            this.txtLicenseType.ReadOnly = true;
            this.txtLicenseType.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtLicenseType.Size = new System.Drawing.Size(117, 20);
            this.txtLicenseType.TabIndex = 45;
            // 
            // cmdResetTrial
            // 
            this.cmdResetTrial.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cmdResetTrial.BackColor = System.Drawing.SystemColors.Control;
            this.cmdResetTrial.Cursor = System.Windows.Forms.Cursors.Default;
            this.cmdResetTrial.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdResetTrial.ForeColor = System.Drawing.SystemColors.ControlText;
            this.cmdResetTrial.Location = new System.Drawing.Point(552, 88);
            this.cmdResetTrial.Name = "cmdResetTrial";
            this.cmdResetTrial.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.cmdResetTrial.Size = new System.Drawing.Size(73, 21);
            this.cmdResetTrial.TabIndex = 42;
            this.cmdResetTrial.Text = "&Reset Trial";
            this.cmdResetTrial.UseVisualStyleBackColor = false;
            this.cmdResetTrial.Click += new System.EventHandler(this.cmdResetTrial_Click);
            // 
            // cmdKillTrial
            // 
            this.cmdKillTrial.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cmdKillTrial.BackColor = System.Drawing.SystemColors.Control;
            this.cmdKillTrial.Cursor = System.Windows.Forms.Cursors.Default;
            this.cmdKillTrial.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdKillTrial.ForeColor = System.Drawing.SystemColors.ControlText;
            this.cmdKillTrial.Location = new System.Drawing.Point(552, 116);
            this.cmdKillTrial.Name = "cmdKillTrial";
            this.cmdKillTrial.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.cmdKillTrial.Size = new System.Drawing.Size(73, 21);
            this.cmdKillTrial.TabIndex = 41;
            this.cmdKillTrial.Text = "&Kill Trial";
            this.cmdKillTrial.UseVisualStyleBackColor = false;
            this.cmdKillTrial.Click += new System.EventHandler(this.cmdKillTrial_Click);
            // 
            // Picture2
            // 
            this.Picture2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.Picture2.BackColor = System.Drawing.SystemColors.Window;
            this.Picture2.Cursor = System.Windows.Forms.Cursors.Default;
            this.Picture2.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Picture2.ForeColor = System.Drawing.SystemColors.WindowText;
            this.Picture2.Image = ((System.Drawing.Image)(resources.GetObject("Picture2.Image")));
            this.Picture2.Location = new System.Drawing.Point(558, 14);
            this.Picture2.Name = "Picture2";
            this.Picture2.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Picture2.Size = new System.Drawing.Size(55, 55);
            this.Picture2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.Picture2.TabIndex = 39;
            this.Picture2.TabStop = false;
            // 
            // txtRegisteredLevel
            // 
            this.txtRegisteredLevel.AcceptsReturn = true;
            this.txtRegisteredLevel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtRegisteredLevel.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.txtRegisteredLevel.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.txtRegisteredLevel.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRegisteredLevel.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtRegisteredLevel.Location = new System.Drawing.Point(104, 116);
            this.txtRegisteredLevel.MaxLength = 0;
            this.txtRegisteredLevel.Name = "txtRegisteredLevel";
            this.txtRegisteredLevel.ReadOnly = true;
            this.txtRegisteredLevel.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtRegisteredLevel.Size = new System.Drawing.Size(289, 20);
            this.txtRegisteredLevel.TabIndex = 36;
            // 
            // txtChecksum
            // 
            this.txtChecksum.AcceptsReturn = true;
            this.txtChecksum.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.txtChecksum.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.txtChecksum.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtChecksum.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtChecksum.Location = new System.Drawing.Point(104, 136);
            this.txtChecksum.MaxLength = 0;
            this.txtChecksum.Name = "txtChecksum";
            this.txtChecksum.ReadOnly = true;
            this.txtChecksum.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtChecksum.Size = new System.Drawing.Size(81, 20);
            this.txtChecksum.TabIndex = 6;
            // 
            // txtVersion
            // 
            this.txtVersion.AcceptsReturn = true;
            this.txtVersion.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtVersion.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.txtVersion.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.txtVersion.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtVersion.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtVersion.Location = new System.Drawing.Point(104, 36);
            this.txtVersion.MaxLength = 0;
            this.txtVersion.Name = "txtVersion";
            this.txtVersion.ReadOnly = true;
            this.txtVersion.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtVersion.Size = new System.Drawing.Size(289, 20);
            this.txtVersion.TabIndex = 2;
            this.txtVersion.Text = "3";
            // 
            // txtName
            // 
            this.txtName.AcceptsReturn = true;
            this.txtName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtName.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.txtName.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.txtName.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtName.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtName.Location = new System.Drawing.Point(104, 16);
            this.txtName.MaxLength = 0;
            this.txtName.Name = "txtName";
            this.txtName.ReadOnly = true;
            this.txtName.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtName.Size = new System.Drawing.Size(289, 20);
            this.txtName.TabIndex = 1;
            this.txtName.Text = "TestApp";
            // 
            // txtExpiration
            // 
            this.txtExpiration.AcceptsReturn = true;
            this.txtExpiration.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtExpiration.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.txtExpiration.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.txtExpiration.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtExpiration.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtExpiration.Location = new System.Drawing.Point(104, 96);
            this.txtExpiration.MaxLength = 0;
            this.txtExpiration.Name = "txtExpiration";
            this.txtExpiration.ReadOnly = true;
            this.txtExpiration.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtExpiration.Size = new System.Drawing.Size(289, 20);
            this.txtExpiration.TabIndex = 5;
            // 
            // txtUsedDays
            // 
            this.txtUsedDays.AcceptsReturn = true;
            this.txtUsedDays.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtUsedDays.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.txtUsedDays.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.txtUsedDays.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtUsedDays.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtUsedDays.Location = new System.Drawing.Point(104, 76);
            this.txtUsedDays.MaxLength = 0;
            this.txtUsedDays.Name = "txtUsedDays";
            this.txtUsedDays.ReadOnly = true;
            this.txtUsedDays.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtUsedDays.Size = new System.Drawing.Size(289, 20);
            this.txtUsedDays.TabIndex = 4;
            // 
            // txtRegStatus
            // 
            this.txtRegStatus.AcceptsReturn = true;
            this.txtRegStatus.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtRegStatus.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.txtRegStatus.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.txtRegStatus.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRegStatus.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtRegStatus.Location = new System.Drawing.Point(104, 56);
            this.txtRegStatus.MaxLength = 0;
            this.txtRegStatus.Name = "txtRegStatus";
            this.txtRegStatus.ReadOnly = true;
            this.txtRegStatus.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtRegStatus.Size = new System.Drawing.Size(289, 20);
            this.txtRegStatus.TabIndex = 3;
            // 
            // Label9
            // 
            this.Label9.BackColor = System.Drawing.SystemColors.Control;
            this.Label9.Cursor = System.Windows.Forms.Cursors.Default;
            this.Label9.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label9.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Label9.Location = new System.Drawing.Point(196, 138);
            this.Label9.Name = "Label9";
            this.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label9.Size = new System.Drawing.Size(77, 17);
            this.Label9.TabIndex = 44;
            this.Label9.Text = "License Type:";
            this.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // Label16
            // 
            this.Label16.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.Label16.BackColor = System.Drawing.SystemColors.Control;
            this.Label16.Cursor = System.Windows.Forms.Cursors.Default;
            this.Label16.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label16.ForeColor = System.Drawing.Color.Blue;
            this.Label16.Location = new System.Drawing.Point(550, 70);
            this.Label16.Name = "Label16";
            this.Label16.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label16.Size = new System.Drawing.Size(71, 11);
            this.Label16.TabIndex = 40;
            this.Label16.Text = "Activelock V3";
            this.Label16.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label5
            // 
            this.Label5.BackColor = System.Drawing.SystemColors.Control;
            this.Label5.Cursor = System.Windows.Forms.Cursors.Default;
            this.Label5.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label5.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Label5.Location = new System.Drawing.Point(8, 118);
            this.Label5.Name = "Label5";
            this.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label5.Size = new System.Drawing.Size(92, 17);
            this.Label5.TabIndex = 37;
            this.Label5.Text = "Registered Level:";
            // 
            // Label3
            // 
            this.Label3.BackColor = System.Drawing.SystemColors.Control;
            this.Label3.Cursor = System.Windows.Forms.Cursors.Default;
            this.Label3.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label3.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Label3.Location = new System.Drawing.Point(8, 138);
            this.Label3.Name = "Label3";
            this.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label3.Size = new System.Drawing.Size(89, 17);
            this.Label3.TabIndex = 35;
            this.Label3.Text = "DLL Checksum:";
            // 
            // Label2
            // 
            this.Label2.BackColor = System.Drawing.SystemColors.Control;
            this.Label2.Cursor = System.Windows.Forms.Cursors.Default;
            this.Label2.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label2.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Label2.Location = new System.Drawing.Point(8, 38);
            this.Label2.Name = "Label2";
            this.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label2.Size = new System.Drawing.Size(88, 17);
            this.Label2.TabIndex = 32;
            this.Label2.Text = "App Version:";
            // 
            // Label1
            // 
            this.Label1.BackColor = System.Drawing.SystemColors.Control;
            this.Label1.Cursor = System.Windows.Forms.Cursors.Default;
            this.Label1.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label1.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Label1.Location = new System.Drawing.Point(8, 18);
            this.Label1.Name = "Label1";
            this.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label1.Size = new System.Drawing.Size(65, 17);
            this.Label1.TabIndex = 31;
            this.Label1.Text = "App Name:";
            // 
            // Label8
            // 
            this.Label8.BackColor = System.Drawing.SystemColors.Control;
            this.Label8.Cursor = System.Windows.Forms.Cursors.Default;
            this.Label8.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label8.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Label8.Location = new System.Drawing.Point(8, 98);
            this.Label8.Name = "Label8";
            this.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label8.Size = new System.Drawing.Size(65, 17);
            this.Label8.TabIndex = 16;
            this.Label8.Text = "Expiry Date:";
            // 
            // Label7
            // 
            this.Label7.BackColor = System.Drawing.SystemColors.Control;
            this.Label7.Cursor = System.Windows.Forms.Cursors.Default;
            this.Label7.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label7.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Label7.Location = new System.Drawing.Point(8, 78);
            this.Label7.Name = "Label7";
            this.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label7.Size = new System.Drawing.Size(65, 17);
            this.Label7.TabIndex = 15;
            this.Label7.Text = "Days Used:";
            // 
            // Label6
            // 
            this.Label6.BackColor = System.Drawing.SystemColors.Control;
            this.Label6.Cursor = System.Windows.Forms.Cursors.Default;
            this.Label6.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label6.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Label6.Location = new System.Drawing.Point(8, 58);
            this.Label6.Name = "Label6";
            this.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label6.Size = new System.Drawing.Size(65, 17);
            this.Label6.TabIndex = 14;
            this.Label6.Text = "Registered:";
            // 
            // fraReg
            // 
            this.fraReg.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.fraReg.BackColor = System.Drawing.SystemColors.Control;
            this.fraReg.Controls.Add(this.cmdPaste);
            this.fraReg.Controls.Add(this.cmdCopy);
            this.fraReg.Controls.Add(this.cmdKillLicense);
            this.fraReg.Controls.Add(this.txtUser);
            this.fraReg.Controls.Add(this.cmdReqGen);
            this.fraReg.Controls.Add(this.txtReqCodeGen);
            this.fraReg.Controls.Add(this.cmdRegister);
            this.fraReg.Controls.Add(this.txtLibKeyIn);
            this.fraReg.Controls.Add(this.Label13);
            this.fraReg.Controls.Add(this.Label11);
            this.fraReg.Controls.Add(this.Label4);
            this.fraReg.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.fraReg.ForeColor = System.Drawing.Color.Blue;
            this.fraReg.Location = new System.Drawing.Point(0, 183);
            this.fraReg.Name = "fraReg";
            this.fraReg.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.fraReg.Size = new System.Drawing.Size(633, 277);
            this.fraReg.TabIndex = 19;
            this.fraReg.TabStop = false;
            this.fraReg.Text = "Register";
            // 
            // cmdPaste
            // 
            this.cmdPaste.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cmdPaste.BackColor = System.Drawing.SystemColors.Control;
            this.cmdPaste.Cursor = System.Windows.Forms.Cursors.Default;
            this.cmdPaste.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdPaste.ForeColor = System.Drawing.SystemColors.ControlText;
            this.cmdPaste.Image = ((System.Drawing.Image)(resources.GetObject("cmdPaste.Image")));
            this.cmdPaste.Location = new System.Drawing.Point(548, 112);
            this.cmdPaste.Name = "cmdPaste";
            this.cmdPaste.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.cmdPaste.Size = new System.Drawing.Size(23, 23);
            this.cmdPaste.TabIndex = 47;
            this.cmdPaste.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.cmdPaste.UseVisualStyleBackColor = false;
            this.cmdPaste.Click += new System.EventHandler(this.cmdPaste_Click);
            // 
            // cmdCopy
            // 
            this.cmdCopy.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cmdCopy.BackColor = System.Drawing.SystemColors.Control;
            this.cmdCopy.Cursor = System.Windows.Forms.Cursors.Default;
            this.cmdCopy.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdCopy.ForeColor = System.Drawing.SystemColors.ControlText;
            this.cmdCopy.Image = ((System.Drawing.Image)(resources.GetObject("cmdCopy.Image")));
            this.cmdCopy.Location = new System.Drawing.Point(548, 60);
            this.cmdCopy.Name = "cmdCopy";
            this.cmdCopy.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.cmdCopy.Size = new System.Drawing.Size(23, 23);
            this.cmdCopy.TabIndex = 46;
            this.cmdCopy.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.cmdCopy.UseVisualStyleBackColor = false;
            this.cmdCopy.Click += new System.EventHandler(this.cmdCopy_Click);
            // 
            // cmdKillLicense
            // 
            this.cmdKillLicense.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cmdKillLicense.BackColor = System.Drawing.SystemColors.Control;
            this.cmdKillLicense.Cursor = System.Windows.Forms.Cursors.Default;
            this.cmdKillLicense.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdKillLicense.ForeColor = System.Drawing.SystemColors.ControlText;
            this.cmdKillLicense.Location = new System.Drawing.Point(548, 168);
            this.cmdKillLicense.Name = "cmdKillLicense";
            this.cmdKillLicense.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.cmdKillLicense.Size = new System.Drawing.Size(73, 21);
            this.cmdKillLicense.TabIndex = 43;
            this.cmdKillLicense.Text = "&Kill License";
            this.cmdKillLicense.UseVisualStyleBackColor = false;
            this.cmdKillLicense.Click += new System.EventHandler(this.cmdKillLicense_Click);
            // 
            // txtUser
            // 
            this.txtUser.AcceptsReturn = true;
            this.txtUser.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtUser.BackColor = System.Drawing.SystemColors.Window;
            this.txtUser.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.txtUser.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtUser.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtUser.Location = new System.Drawing.Point(96, 20);
            this.txtUser.MaxLength = 0;
            this.txtUser.Name = "txtUser";
            this.txtUser.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtUser.Size = new System.Drawing.Size(445, 20);
            this.txtUser.TabIndex = 7;
            this.txtUser.Text = "Evaluation User";
            // 
            // cmdReqGen
            // 
            this.cmdReqGen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cmdReqGen.BackColor = System.Drawing.SystemColors.Control;
            this.cmdReqGen.Cursor = System.Windows.Forms.Cursors.Default;
            this.cmdReqGen.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdReqGen.ForeColor = System.Drawing.SystemColors.ControlText;
            this.cmdReqGen.Location = new System.Drawing.Point(548, 36);
            this.cmdReqGen.Name = "cmdReqGen";
            this.cmdReqGen.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.cmdReqGen.Size = new System.Drawing.Size(73, 21);
            this.cmdReqGen.TabIndex = 9;
            this.cmdReqGen.Text = "&Generate";
            this.cmdReqGen.UseVisualStyleBackColor = false;
            this.cmdReqGen.Click += new System.EventHandler(this.cmdReqGen_Click);
            // 
            // txtReqCodeGen
            // 
            this.txtReqCodeGen.AcceptsReturn = true;
            this.txtReqCodeGen.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtReqCodeGen.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.txtReqCodeGen.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.txtReqCodeGen.Font = new System.Drawing.Font("Courier New", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtReqCodeGen.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtReqCodeGen.Location = new System.Drawing.Point(96, 40);
            this.txtReqCodeGen.MaxLength = 0;
            this.txtReqCodeGen.Multiline = true;
            this.txtReqCodeGen.Name = "txtReqCodeGen";
            this.txtReqCodeGen.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtReqCodeGen.Size = new System.Drawing.Size(445, 51);
            this.txtReqCodeGen.TabIndex = 8;
            // 
            // cmdRegister
            // 
            this.cmdRegister.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cmdRegister.BackColor = System.Drawing.SystemColors.Control;
            this.cmdRegister.Cursor = System.Windows.Forms.Cursors.Default;
            this.cmdRegister.Enabled = false;
            this.cmdRegister.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdRegister.ForeColor = System.Drawing.SystemColors.ControlText;
            this.cmdRegister.Location = new System.Drawing.Point(548, 140);
            this.cmdRegister.Name = "cmdRegister";
            this.cmdRegister.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.cmdRegister.Size = new System.Drawing.Size(73, 21);
            this.cmdRegister.TabIndex = 11;
            this.cmdRegister.Text = "&Register";
            this.cmdRegister.UseVisualStyleBackColor = false;
            this.cmdRegister.Click += new System.EventHandler(this.cmdRegister_Click);
            // 
            // txtLibKeyIn
            // 
            this.txtLibKeyIn.AcceptsReturn = true;
            this.txtLibKeyIn.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtLibKeyIn.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.txtLibKeyIn.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.txtLibKeyIn.Font = new System.Drawing.Font("Courier New", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtLibKeyIn.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtLibKeyIn.Location = new System.Drawing.Point(96, 92);
            this.txtLibKeyIn.MaxLength = 0;
            this.txtLibKeyIn.Multiline = true;
            this.txtLibKeyIn.Name = "txtLibKeyIn";
            this.txtLibKeyIn.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtLibKeyIn.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtLibKeyIn.Size = new System.Drawing.Size(445, 171);
            this.txtLibKeyIn.TabIndex = 10;
            this.txtLibKeyIn.TextChanged += new System.EventHandler(this.txtLibKeyIn_TextChanged);
            // 
            // Label13
            // 
            this.Label13.BackColor = System.Drawing.SystemColors.Control;
            this.Label13.Cursor = System.Windows.Forms.Cursors.Default;
            this.Label13.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label13.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Label13.Location = new System.Drawing.Point(8, 20);
            this.Label13.Name = "Label13";
            this.Label13.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label13.Size = new System.Drawing.Size(89, 17);
            this.Label13.TabIndex = 20;
            this.Label13.Text = "User Name:";
            // 
            // Label11
            // 
            this.Label11.BackColor = System.Drawing.SystemColors.Control;
            this.Label11.Cursor = System.Windows.Forms.Cursors.Default;
            this.Label11.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label11.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Label11.Location = new System.Drawing.Point(8, 40);
            this.Label11.Name = "Label11";
            this.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label11.Size = new System.Drawing.Size(89, 17);
            this.Label11.TabIndex = 19;
            this.Label11.Text = "Installation Code:";
            // 
            // Label4
            // 
            this.Label4.BackColor = System.Drawing.SystemColors.Control;
            this.Label4.Cursor = System.Windows.Forms.Cursors.Default;
            this.Label4.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label4.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Label4.Location = new System.Drawing.Point(8, 92);
            this.Label4.Name = "Label4";
            this.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label4.Size = new System.Drawing.Size(89, 17);
            this.Label4.TabIndex = 18;
            this.Label4.Text = "License Key:";
            // 
            // _SSTab1_TabPage1
            // 
            this._SSTab1_TabPage1.Controls.Add(this.lblLockStatus);
            this._SSTab1_TabPage1.Controls.Add(this.lblLockStatus2);
            this._SSTab1_TabPage1.Controls.Add(this.lblTrialInfo);
            this._SSTab1_TabPage1.Controls.Add(this.Frame1);
            this._SSTab1_TabPage1.Location = new System.Drawing.Point(4, 22);
            this._SSTab1_TabPage1.Name = "_SSTab1_TabPage1";
            this._SSTab1_TabPage1.Size = new System.Drawing.Size(640, 470);
            this._SSTab1_TabPage1.TabIndex = 1;
            this._SSTab1_TabPage1.Text = "Sample App";
            // 
            // lblLockStatus
            // 
            this.lblLockStatus.BackColor = System.Drawing.SystemColors.Control;
            this.lblLockStatus.Cursor = System.Windows.Forms.Cursors.Default;
            this.lblLockStatus.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblLockStatus.ForeColor = System.Drawing.SystemColors.ControlText;
            this.lblLockStatus.Location = new System.Drawing.Point(8, 30);
            this.lblLockStatus.Name = "lblLockStatus";
            this.lblLockStatus.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblLockStatus.Size = new System.Drawing.Size(214, 25);
            this.lblLockStatus.TabIndex = 41;
            this.lblLockStatus.Text = "Application Functionalities Are Currently: ";
            // 
            // lblLockStatus2
            // 
            this.lblLockStatus2.BackColor = System.Drawing.SystemColors.Control;
            this.lblLockStatus2.Cursor = System.Windows.Forms.Cursors.Default;
            this.lblLockStatus2.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblLockStatus2.ForeColor = System.Drawing.SystemColors.ControlText;
            this.lblLockStatus2.Location = new System.Drawing.Point(224, 30);
            this.lblLockStatus2.Name = "lblLockStatus2";
            this.lblLockStatus2.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblLockStatus2.Size = new System.Drawing.Size(301, 17);
            this.lblLockStatus2.TabIndex = 42;
            this.lblLockStatus2.Text = "Disabled";
            // 
            // lblTrialInfo
            // 
            this.lblTrialInfo.BackColor = System.Drawing.SystemColors.Control;
            this.lblTrialInfo.Cursor = System.Windows.Forms.Cursors.Default;
            this.lblTrialInfo.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTrialInfo.ForeColor = System.Drawing.SystemColors.ControlText;
            this.lblTrialInfo.Location = new System.Drawing.Point(8, 54);
            this.lblTrialInfo.Name = "lblTrialInfo";
            this.lblTrialInfo.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblTrialInfo.Size = new System.Drawing.Size(318, 25);
            this.lblTrialInfo.TabIndex = 43;
            this.lblTrialInfo.Text = "NOTE: All application functionalities are available in Trial Mode.";
            // 
            // Frame1
            // 
            this.Frame1.BackColor = System.Drawing.SystemColors.Control;
            this.Frame1.Controls.Add(this.optForm1);
            this.Frame1.Controls.Add(this.optForm0);
            this.Frame1.Controls.Add(this.cboSpeed);
            this.Frame1.Controls.Add(this.chkPause);
            this.Frame1.Controls.Add(this.chkFlash);
            this.Frame1.Controls.Add(this.chkScroll);
            this.Frame1.Controls.Add(this.lblHost);
            this.Frame1.Controls.Add(this.lblSpeed);
            this.Frame1.Cursor = System.Windows.Forms.Cursors.Default;
            this.Frame1.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Frame1.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Frame1.Location = new System.Drawing.Point(8, 96);
            this.Frame1.Name = "Frame1";
            this.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Frame1.Size = new System.Drawing.Size(503, 125);
            this.Frame1.TabIndex = 39;
            // 
            // optForm1
            // 
            this.optForm1.BackColor = System.Drawing.SystemColors.Control;
            this.optForm1.Cursor = System.Windows.Forms.Cursors.Default;
            this.optForm1.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.optForm1.ForeColor = System.Drawing.SystemColors.ControlText;
            this.optForm1.Location = new System.Drawing.Point(202, 83);
            this.optForm1.Name = "optForm1";
            this.optForm1.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.optForm1.Size = new System.Drawing.Size(93, 15);
            this.optForm1.TabIndex = 27;
            this.optForm1.TabStop = true;
            this.optForm1.Text = "Option 2";
            this.optForm1.UseVisualStyleBackColor = false;
            // 
            // optForm0
            // 
            this.optForm0.BackColor = System.Drawing.SystemColors.Control;
            this.optForm0.Checked = true;
            this.optForm0.Cursor = System.Windows.Forms.Cursors.Default;
            this.optForm0.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.optForm0.ForeColor = System.Drawing.SystemColors.ControlText;
            this.optForm0.Location = new System.Drawing.Point(202, 66);
            this.optForm0.Name = "optForm0";
            this.optForm0.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.optForm0.Size = new System.Drawing.Size(93, 15);
            this.optForm0.TabIndex = 26;
            this.optForm0.TabStop = true;
            this.optForm0.Text = "Option 1";
            this.optForm0.UseVisualStyleBackColor = false;
            // 
            // cboSpeed
            // 
            this.cboSpeed.BackColor = System.Drawing.SystemColors.Window;
            this.cboSpeed.Cursor = System.Windows.Forms.Cursors.Default;
            this.cboSpeed.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboSpeed.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cboSpeed.ForeColor = System.Drawing.SystemColors.WindowText;
            this.cboSpeed.Items.AddRange(new object[] {
            "Slowest",
            "Slow",
            "Normal",
            "Fast",
            "Fastest"});
            this.cboSpeed.Location = new System.Drawing.Point(356, 4);
            this.cboSpeed.Name = "cboSpeed";
            this.cboSpeed.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.cboSpeed.Size = new System.Drawing.Size(109, 22);
            this.cboSpeed.TabIndex = 25;
            // 
            // chkPause
            // 
            this.chkPause.BackColor = System.Drawing.SystemColors.Control;
            this.chkPause.Cursor = System.Windows.Forms.Cursors.Default;
            this.chkPause.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkPause.ForeColor = System.Drawing.SystemColors.ControlText;
            this.chkPause.Location = new System.Drawing.Point(0, 26);
            this.chkPause.Name = "chkPause";
            this.chkPause.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.chkPause.Size = new System.Drawing.Size(171, 21);
            this.chkPause.TabIndex = 24;
            this.chkPause.Text = "Checkbox for Level 3 only";
            this.chkPause.UseVisualStyleBackColor = false;
            // 
            // chkFlash
            // 
            this.chkFlash.BackColor = System.Drawing.SystemColors.Control;
            this.chkFlash.Cursor = System.Windows.Forms.Cursors.Default;
            this.chkFlash.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkFlash.ForeColor = System.Drawing.SystemColors.ControlText;
            this.chkFlash.Location = new System.Drawing.Point(0, 46);
            this.chkFlash.Name = "chkFlash";
            this.chkFlash.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.chkFlash.Size = new System.Drawing.Size(174, 33);
            this.chkFlash.TabIndex = 23;
            this.chkFlash.Text = "Checkbox for ALL Levels";
            this.chkFlash.UseVisualStyleBackColor = false;
            // 
            // chkScroll
            // 
            this.chkScroll.BackColor = System.Drawing.SystemColors.Control;
            this.chkScroll.Checked = true;
            this.chkScroll.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkScroll.Cursor = System.Windows.Forms.Cursors.Default;
            this.chkScroll.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkScroll.ForeColor = System.Drawing.SystemColors.ControlText;
            this.chkScroll.Location = new System.Drawing.Point(0, 6);
            this.chkScroll.Name = "chkScroll";
            this.chkScroll.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.chkScroll.Size = new System.Drawing.Size(178, 15);
            this.chkScroll.TabIndex = 22;
            this.chkScroll.Text = "Checkbox for ALL Levels";
            this.chkScroll.UseVisualStyleBackColor = false;
            // 
            // lblHost
            // 
            this.lblHost.BackColor = System.Drawing.SystemColors.Control;
            this.lblHost.Cursor = System.Windows.Forms.Cursors.Default;
            this.lblHost.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblHost.ForeColor = System.Drawing.SystemColors.ControlText;
            this.lblHost.Location = new System.Drawing.Point(204, 46);
            this.lblHost.Name = "lblHost";
            this.lblHost.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblHost.Size = new System.Drawing.Size(184, 17);
            this.lblHost.TabIndex = 29;
            this.lblHost.Text = "Option Buttons for ALL Levels:";
            // 
            // lblSpeed
            // 
            this.lblSpeed.BackColor = System.Drawing.SystemColors.Control;
            this.lblSpeed.Cursor = System.Windows.Forms.Cursors.Default;
            this.lblSpeed.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSpeed.ForeColor = System.Drawing.SystemColors.ControlText;
            this.lblSpeed.Location = new System.Drawing.Point(202, 6);
            this.lblSpeed.Name = "lblSpeed";
            this.lblSpeed.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblSpeed.Size = new System.Drawing.Size(148, 17);
            this.lblSpeed.TabIndex = 28;
            this.lblSpeed.Text = "Activated with Level 4 Only";
            // 
            // frmMain
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(664, 509);
            this.Controls.Add(this.SStab1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ALTestApp - ActiveLock Test Application for VC# 2008 v3.6";
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.SStab1.ResumeLayout(false);
            this._SSTab1_TabPage0.ResumeLayout(false);
            this.fraRegStatus.ResumeLayout(false);
            this.fraRegStatus.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Picture2)).EndInit();
            this.fraReg.ResumeLayout(false);
            this.fraReg.PerformLayout();
            this._SSTab1_TabPage1.ResumeLayout(false);
            this.Frame1.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion


		#region "Methods"
		
		private static string Encrypt(ref string strdata) 
		{ 
			int n; 
			n = strdata.Length - 1; 
			StringBuilder mStringBuilder = new StringBuilder(null);
			for (int Index = 0; Index <= n; Index++) 
			{ 
				mStringBuilder.Append(modMain.Asc(strdata.Substring(Index,1)) * 7);
			} 
			return mStringBuilder.ToString(); 
		}

		private void SetFunctionalities(bool isEnabled)
		{ 
			chkScroll.Enabled = isEnabled; 
			chkFlash.Enabled = isEnabled; 
			lblHost.Enabled = isEnabled; 
			optForm0.Enabled = isEnabled; 
			optForm1.Enabled = isEnabled; 
			chkPause.Enabled = isEnabled; 
			lblSpeed.Enabled = isEnabled; 
			cboSpeed.Enabled = isEnabled; 
			if (isEnabled) 
			{ 
				if (txtRegisteredLevel.Text.Length>0) 
				{ 
					lblLockStatus2.Text = "Enabled with " + txtRegisteredLevel.Text; 
					chkPause.Enabled = (txtRegisteredLevel.Text.IndexOf("Level 3") >= 0); 
					lblSpeed.Enabled = (txtRegisteredLevel.Text.IndexOf("Level 4") >= 0); 
					cboSpeed.Enabled = (txtRegisteredLevel.Text.IndexOf("Level 4") >= 0); 
				} 
				else 
				{ 
					lblLockStatus2.Text = "Enabled with " + txtUsedDays.Text; 
				} 
			} 
			else 
			{ 
				lblLockStatus2.Text = "Disabled (Registration Required)"; 
				chkPause.Enabled = false; 
				lblSpeed.Enabled = false; 
				cboSpeed.Enabled = false; 
			} 
		}

        private bool CheckForResources(params object[] MyArray) 
        {
            /*MyArray is a list of things to check
            'These can be DLLs or OCXs

            'Files, by default, are searched for in the Windows System Directory
            'Exceptions;
            '   Begins with a # means it should be in the same directory with the executable
            '   Contains the full path (anything with a "\")

            'Typical names would be "#aaa.dll", "mydll.dll", "myocx.ocx", "comdlg32.ocx", "mscomctl.ocx", "msflxgrd.ocx"

            'If the file has no extension, we;
            '     assume it's a DLL, and if it still can't be found
            '     assume it's an OCX            */

            try
            {
                bool foundIt;                
                int j;
                string systemDir, s, pathName;

                // WhereIsDLL("") 'initialize

                systemDir = WindowsSystemDirectory(); //'Get the Windows system directory
                foreach( Object y in MyArray )
                {                
                    foundIt = false;
                    s = (string)y;                   

                    if( s.StartsWith("#") )
                    {                       
                        pathName = Application.StartupPath;                        
                        s = s.Substring(1);
                    }
                    else if( s.Contains(@"\") )
                    {
                        j = s.IndexOf(@"\");
                        pathName = s.Substring(0, j - 1);
                        s = s.Substring(j + 1);
                    } else
                        pathName = systemDir;
                    
                    if( s.Contains(".") )
                    {
                        if( File.Exists(pathName + @"\" + s) ) 
                            foundIt = true;
                    }
                    else if ( File.Exists(pathName + @"\" + s + ".DLL") ) 
                        foundIt = true;
                    else if ( File.Exists(pathName + @"\" + s + ".OCX") )
                    {
                        foundIt = true;
                        s = s + ".OCX";
                    }

                    if( !foundIt )
                    {
                        MessageBox.Show(s + " could not be found in " + pathName + "." + "\n\r" + 
                            System.Reflection.Assembly.GetExecutingAssembly().GetName().Name + " cannot run without this library file!" +
                            "\r\n\r\nExiting!");
                        Application.Exit();
                    }
                }

                return true;                
            }
            catch
            {
                MessageBox.Show("CheckForResources error");
                // 'an error kills the program
                Application.Exit();
                return false;
            }
        }
        private string WindowsSystemDirectory()
        {
            //int cnt;
            string s;
            uint dl;

            //cnt = 254;

            StringBuilder sbSystemDir = new StringBuilder(256);
            dl = GetSystemDirectory(sbSystemDir,256);

            s = sbSystemDir.ToString();

            return s;
        }

        /*public string LooseSpace(string invoer)
        {
            //This routine terminates a string if it detects char 0.

            int P;

            P = invoer.IndexOf('\0');
            if (P >= 0)
                return invoer.Substring(0, P);
            else 
                return invoer;
        }*/

		#endregion



		#region "Events"

		private void cmdCopy_Click(object sender, System.EventArgs e)
		{
			  DataObject aDataObject = new DataObject();
        aDataObject.SetData(DataFormats.Text, txtReqCodeGen.Text);
        Clipboard.SetDataObject(aDataObject);
		}

		private void cmdPaste_Click(object sender, System.EventArgs e)
		{
			if( Clipboard.GetDataObject().GetDataPresent(DataFormats.Text))
			{
                if (Clipboard.GetDataObject().GetData(DataFormats.Text).ToString() == txtReqCodeGen.Text ) 
				{
                    MessageBox.Show("You cannot paste the Installation Code into the Liberation Key field.", modMain.MSGBOXCAPTION,MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
					return;
				}
				txtLibKeyIn.Text = Clipboard.GetDataObject().GetData(DataFormats.Text).ToString();
			}
		}

		private void ActiveLockEventSink_ValidateValue(ref string Value)
		{
		    /*'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Key Validation Functionalities
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' ActiveLock raises this event typically when it needs a value to be encrypted.
            ' We can use any kind of encryption we'd like here, as long as it's deterministic.
            ' i.e. there's a one-to-one correspondence between unencrypted value and encrypted value.
            ' NOTE: BlowFish is NOT an example of deterministic encryption so you can't use it here.*/
			Value = Encrypt(ref Value);
		}

		private void frmMain_Load(object sender, System.EventArgs e)
		{
			string autoRegisterKey=null; 
			bool boolAutoRegisterKeyPath = false; 
			string[] A; 
			
			string strMsg=null;
            string strRemainingTrialDays = null, strRemainingTrialRuns = null, strTrialLength = null, strUsedDays = null, strExpirationDate = null,
                strRegisteredUser = null, strRegisteredLevel = null, strLicenseClass = null, strMaxCount = null, strLicenseFileType = null,
                strLicenseType = null;
			try
			{
                // Form's caption
                this.Text = "ALTestApp3NET - ActiveLock Test Application for VC#2010 - v3.7"; //& Application.ProductVersion
                
                // Check the existence of necessary files to run this application
                // This is not necessary if you're not using these controls in your app.
                CheckForResources( "#ActiveLock37Net.dll", "comctl32.ocx", "tabctl32.ocx");

                // Set the path to the license file (LIC) and ALL file (if it exists)
                StringBuilder sbAppfilePath = new StringBuilder(260);
                SHGetSpecialFolderPath(IntPtr.Zero, sbAppfilePath, 46, false); // 46 is for ...\All Users\Documents folder.
                string AppfilePath = sbAppfilePath.ToString();

				Globals MyAL = new Globals(); 
				MyActiveLock = MyAL.NewInstance(); 
				
				MyActiveLock.SoftwareName = modMain.SOFTWARENAME; 
				txtName.Text = MyActiveLock.SoftwareName; 
				
                // Note: Do not use (App.Major & "." & App.Minor & "." & App.Revision)
                // since the license will fail with version incremented exe builds
				MyActiveLock.SoftwareVersion = "3.7"; // WARNING *** WARNING *** DO NOT USE App.Major & "." & App.Minor & "." & App.Revision
				txtVersion.Text = MyActiveLock.SoftwareVersion; 
				// This should be set to protect yourself against ResetTrial abuse
				MyActiveLock.SoftwarePassword = Convert.ToChar(99).ToString() + Convert.ToChar(111).ToString() + Convert.ToChar(111).ToString() + Convert.ToChar(108).ToString(); 
                
                // Set whether the software/application will use a short key or RSA method
                // alsRSA covers both ALCrypto and RSA native classes approach.
                // RSA classes in .NET allows you to pick from several cipher strengths
                // however ALCrypto uses 1024 bit strength key only.
                // alsShortKeyMD5 is for short key protection only
                // WARNING: Short key licenses use the lockFingerprint by default
                //.LicenseKeyType = ActiveLock37Net.IActiveLock.ALLicenseKeyTypes.alsShortKeyMD5
                MyActiveLock.LicenseKeyType = ActiveLock37Net.IActiveLock.ALLicenseKeyTypes.alsRSA;

                // Set the Trial Feature properties
                // If you don't want to use the trial feature in your app, set the TrialType
                // property to trialNone.
				
                // Set the trial type property
                // this is either trialDays, or trialRuns or trialNone.
				MyActiveLock.TrialType = IActiveLock.ALTrialTypes.trialDays; 
                // Set the Trial Length property.
                // This number represents the number of days or the number of runs (whichever is applicable).
				MyActiveLock.TrialLength = 15; 
				if (MyActiveLock.TrialType != IActiveLock.ALTrialTypes.trialNone & MyActiveLock.TrialLength == 0) 
				{ 
					/*
					 *  Do Nothing
                     *  In such cases Activelock automatically generates errors -11001100 or -11001101
                     *  to indicate that you're using the trial feature but, trial length was not specified 
					 */
				} 
                // Comment the following statement to use a certain trial data hiding technique
                // Use OR to combine one or more trial hiding techniques
                // or don't use this property to use ALL techniques
                // WARNING: trialRegistryPerUser is "Per User"; this means each user trial feature 
                // is controlled that user's own registry hive.
                // This means initiating a trial with one user does not initiate a trial for another user.
                // trialHiddenFolder and trialSteganography are for "All Users"
				MyActiveLock.TrialHideType = IActiveLock.ALTrialHideTypes.trialHiddenFolder | IActiveLock.ALTrialHideTypes.trialRegistryPerUser | IActiveLock.ALTrialHideTypes.trialSteganography; 

                // Use the following if you'd like to make the trial warning message persistent 
                // until a license is registered.
                // Expiration of the trial is not so clear to some users without any warnings.
                // .trialWarningTemporary will show the trial expired warning only once.
                MyActiveLock.TrialWarning = IActiveLock.ALTrialWarningTypes.trialWarningPersistent;
				
                // Set the Software code
                // This is the same thing as VCode
                // Run Alugen first and create a VCode and GCode 
                // for the software name and version number you used above
                // Then copy and use the VCode as the PUB_KEY here.
                // It's up to you to encrypt it; just makes it more secure
                // Enc encodes, Dec decodes the public key (VCode)
                // Change Enc() and Dec(0 the way you want.
                //MyActiveLock.SoftwareCode = modMain.Dec(modMain.PUB_KEY);
                MyActiveLock.SoftwareCode = modMain.PUB_KEY; 

                // uncomment the following when unmanaged Activelock3NET.dll is used
                //.LockType = ActiveLock3.ALLockTypes.lockNone
                                 
                // Set the Hardware keys
                // In order to pick the keys that you want to lock to in Alugen, use lockNone only
                // Example: lockWindows Or lockComp
                // You can combine any lockType(s) using OR as above
                // WARNING: Short key licenses use the lockFingerprint by default
                // WARNING: This has no effect for short key licenses
                //.LockType = ActiveLock37Net.IActiveLock.ALLockTypes.lockNone
				MyActiveLock.LockType = IActiveLock.ALLockTypes.lockNone; 

                //.LockType = ActiveLock37Net.IActiveLock.ALLockTypes.lockIP Or _
                //ActiveLock37Net.IActiveLock.ALLockTypes.lockComp

                // If you want to lock to any keys explicitly, combine them using OR
                // But you won't be able to uncheck/check any of them while in Alugen (too late at that point).
                //.LockType = _
                // ActiveLock37Net.IActiveLock.ALLockTypes.lockBIOS Or _
                // ActiveLock37Net.IActiveLock.ALLockTypes.lockComp Or _
                // ActiveLock37Net.IActiveLock.ALLockTypes.lockHD Or _
                // ActiveLock37Net.IActiveLock.ALLockTypes.lockHDFirmware Or _
                // ActiveLock37Net.IActiveLock.ALLockTypes.lockIP Or _
                // ActiveLock37Net.IActiveLock.ALLockTypes.lockMotherboard Or _
                // ActiveLock37Net.IActiveLock.ALLockTypes.lockWindows Or _
                // ActiveLock37Net.IActiveLock.ALLockTypes.lockMAC Or _
                // ActiveLock37Net.IActiveLock.ALLockTypes.lockExternalIP Or _
                // ActiveLock37Net.IActiveLock.ALLockTypes.lockFingerprint Or _
                // ActiveLock37Net.IActiveLock.ALLockTypes.lockSID Or _
                // ActiveLock37Net.IActiveLock.ALLockTypes.lockCPUID Or _
                // ActiveLock37Net.IActiveLock.ALLockTypes.lockBaseboardID Or _
                // ActiveLock37Net.IActiveLock.ALLockTypes.lockVideoID

                // Set the .ALL file path if you're using an ALL file.
                // .ALL is an auto registration file.
                // You generate .ALL files via Alugen and then send to the users
                // They put the .ALL file in the directory you specify below
                // .ALL simply contains the license key
                // WARNING: .ALL files are deleted after they are used.
                // It is recommended that you use both SoftwareName and Version in the .ALL filename
                // since multiple .ALL files might exist in the same directory
                // If you don't want to use the software name and version number explicitly, use an .ALL
                // filename that is specific to this application
                            
                if( !Directory.Exists(AppfilePath + @"\" + MyActiveLock.SoftwareName + MyActiveLock.SoftwareVersion) )
                    Directory.CreateDirectory( AppfilePath + @"\" + MyActiveLock.SoftwareName + MyActiveLock.SoftwareVersion );

                strAutoRegisterKeyPath = AppfilePath + @"\" + MyActiveLock.SoftwareName + MyActiveLock.SoftwareVersion +
                    @"\" + MyActiveLock.SoftwareName + MyActiveLock.SoftwareVersion + ".all";              
				MyActiveLock.AutoRegisterKeyPath = strAutoRegisterKeyPath; 
                if (File.Exists(strAutoRegisterKeyPath)) 
					boolAutoRegisterKeyPath = true; 

                // Set if auto registration will be used.
                // Auto registration uses the ALL file for license registration.
                MyActiveLock.AutoRegister = ActiveLock37Net.IActiveLock.ALAutoRegisterTypes.alsEnableAutoRegistration;

                // Set the Time Server check for Clock Tampering
                // This is optional but highly recommended.
                // Although Activelock makes every effort to check if the system clock was tampered,
                // checking a time server is the guaranteed way of knowing the correct UTC time/day.
                // This feature might add some delay to your apps start-up time.
                MyActiveLock.CheckTimeServerForClockTampering = ActiveLock37Net.IActiveLock.ALTimeServerTypes.alsDontCheckTimeServer; // use alsCheckTimeServer to enforce time server checks for clock tampering check
                //.CheckTimeServerForClockTampering = ActiveLock37Net.IActiveLock.ALTimeServerTypes.alsCheckTimeServer

                // Set the system files clock tampering check
                // This feature might add some delay to your apps start-up time.
                MyActiveLock.CheckSystemFilesForClockTampering = ActiveLock37Net.IActiveLock.ALSystemFilesTypes.alsDontCheckSystemFiles; // use alsCheckSystemFiles to enforce system files scanning for clock tampering check
                //.CheckSystemFilesForClockTampering = ActiveLock37Net.IActiveLock.ALSystemFilesTypes.alsCheckSystemFiles
                
                // Set the license file format; this could be encrypted or plain
                // Even in a plain file format, certain keys and dates are still encrypted.
                MyActiveLock.LicenseFileType = ActiveLock37Net.IActiveLock.ALLicenseFileTypes.alsLicenseFilePlain;

                // Verify AL's authenticity
				txtChecksum.Text = modMain.VerifyActiveLockNETdll();
                if (txtChecksum.Text == "-1")
                {
                    Environment.Exit(0);
                }
				
                // Initialize the keystore. We use a File keystore in this case.
                // The other type alsRegistry is NOT supported.
                MyActiveLock.KeyStoreType = ActiveLock37Net.IActiveLock.LicStoreType.alsFile;
                // uncomment the following when unmanaged Activelock3NET.dll is used
                //MyActiveLock.KeyStoreType = ActiveLock3.LicStoreType.alsFile

                // Initialize the keystore. We use a File keystore in this case.
				// MyActiveLock.KeyStoreType = IActiveLock.LicStoreType.alsFile; 
				// Path to the license file
				strKeyStorePath = modMain.AppPath() + "\\" + modMain.SOFTWARENAME + ".lic"; 
				System.Diagnostics.Debug.WriteLine("License path is " + strKeyStorePath); 
				MyActiveLock.KeyStorePath = strKeyStorePath; 
				
                if( !Directory.Exists(AppfilePath + @"\" + MyActiveLock.SoftwareName + MyActiveLock.SoftwareVersion) )
                    Directory.CreateDirectory( AppfilePath + @"\" + MyActiveLock.SoftwareName + MyActiveLock.SoftwareVersion );
                strKeyStorePath = AppfilePath + @"\" + MyActiveLock.SoftwareName + MyActiveLock.SoftwareVersion +
                        @"\" + MyActiveLock.SoftwareName + MyActiveLock.SoftwareVersion + ".lic";
                
                // Obtain the EventNotifier so that we can receive notifications from AL.
				ActiveLockEventSink = MyActiveLock.EventNotifier; 
				
                // Initialize AL
                // Important: If you're not going to put Alcrypto3NET.dll under
                // the system32 directory, you should pass the path of the exe
                // to the Init() method otherwise this call will fail
                // Putting Alcrypto3NET.dll under the system32 is a problem with ASP.NET apps
                // since Activelock3NET is shared between .NET apps.
                // Use the following with ASP.NET applications
                // MyActiveLock.Init(Application.StartupPath & "\bin");
                // Use the following with C# applications
				MyActiveLock.Init(Application.StartupPath,ref strKeyStorePath); 
				if (File.Exists(strKeyStorePath) && boolAutoRegisterKeyPath == true && autoRegisterKey.Length>0) 
				{ 
				  // This means, an ALL file existed and was used to create a LIC file
                  // Init() method successfully registered the ALL file
                  // and returned the license key
                  // You can process that key here to see if there is any abuse, etc.
                  // ie. whether the key was used before, etc.
				} 
				//set dummy test combo
				cboSpeed.Text = cboSpeed.Items[2].ToString(); 
				
                // Check registration status
                // Acquire() method does both trial and regular licensing
                // If it generates an error, that means there NO trial, NO license
                // If no error and returns a string, there's a trial but No license. Parse the string to display a trial message.
                // If no error and no string returned, you've got a valid license.

                // In case the Acquire method generates an error, so no license and no trial:
                // If InStr(1, Err.Description, "No valid license") > 0 Or InStr(1, Err.Description, "license invalid") > 0 Then '-2147221503 & -2147221502
                MyActiveLock.Acquire(ref strMsg, ref strRemainingTrialDays, ref strRemainingTrialRuns, ref strTrialLength,
                    ref strUsedDays, ref strExpirationDate, ref strRegisteredUser, ref strRegisteredLevel, ref strLicenseClass,
                    ref strMaxCount, ref strLicenseFileType, ref strLicenseType); 
                // strMsg is to get the trial status
                // All other parameters are Optional and you can actually get all of them
                // using MyActivelock.Property usage, but keep in mind that 
                // doing so will check the license every time making this a time consuming 
                // way of reading those properties
                // The fastest approach is to use the arguments from Acquire() method.
				if (strMsg!=null && strMsg.Length>0) //There's a trial
				{ 
					A = strMsg.Split(new char[] {Convert.ToChar(13)}); 
					txtRegStatus.Text = A[0]; 
					txtUsedDays.Text = A[1].Replace("\n",""); 
					SetFunctionalities(true); 
					frmSplash mfrmsplash = new frmSplash(); 
					mfrmsplash.lblInfo.Text = "\r\n" + strMsg; 
					mfrmsplash.Visible = true; 
					mfrmsplash.Refresh(); 
					Thread.Sleep(3000); //wait about 3 seconds
					mfrmsplash.Close(); 
					cmdKillTrial.Visible = true; 
					cmdResetTrial.Visible = true; 
					txtLicenseType.Text = "Free Trial"; 
					this.Refresh(); 
					return; 
				} 
				else 
				{ 
					cmdKillTrial.Visible = false; 
					cmdResetTrial.Visible = false; 
				} 
                // If you are here already, that means you have a valid license.
                // Set the textboxes in your app accordingly.
				txtRegStatus.Text = "Registered";
                txtUsedDays.Text = strUsedDays;
                txtExpiration.Text = strExpirationDate;
				if (txtExpiration.Text.Length == 0) 				
					txtExpiration.Text = "Permanent"; // App has a permanent license


                txtUser.Text = strRegisteredUser.Substring(0, strRegisteredUser.IndexOf("&&&"));
				txtRegisteredLevel.Text = strRegisteredLevel; 

				//Read the license file into a string to determine the license type
				string strBuff; 
				strBuff = modMain.ReadFile(strKeyStorePath);
				if (strBuff.IndexOf("LicenseType=3")>=0) 
				{ 
					txtLicenseType.Text = "Time Limited"; 
				} 
				else if (strBuff.IndexOf("LicenseType=1")>=0) 
				{ 
					txtLicenseType.Text = "Periodic"; 
				} 
				else if (strBuff.IndexOf("LicenseType=2")>=0) 
				{ 
					txtLicenseType.Text = "Permanent"; 
				} 
				SetFunctionalities(true); 
				return; 
			}
			catch (Exception ex)
			{
				//not registered
				SetFunctionalities(false); 
				if ((ex.Message.ToLower().IndexOf("no valid license")>=0) == false && noTrialThisTime == false) 
				{ 
					MessageBox.Show(ex.Message, modMain.MSGBOXCAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error); 
				} 
				txtRegStatus.Text = ex.Message; 
				txtLicenseType.Text = "None";
                txtUser.Text = System.Environment.UserName;
				if (strMsg!=null && strMsg.Length>0) 
				{ 
					MessageBox.Show(strMsg, modMain.MSGBOXCAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information); 
				} 
			}
			return; 
		}

		private void cmdReqGen_Click(object sender, System.EventArgs e)
		{
			if (txtUser.Text.Length == 0)
			{
				MessageBox.Show("User Name field is blank."  ,modMain.MSGBOXCAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
			}
//Generate Request code to Lock
			if (MyActiveLock == null) 
			{ 
				noTrialThisTime = true; 
				frmMain_Load(this, new System.EventArgs()); 
			} 
			if (txtRegStatus.Text != "Registered") 
			{ 
				txtRegStatus.Text = ""; 
			} 
			if (!(modMain.isNumeric(txtUsedDays.Text))) 
			{ 
				txtUsedDays.Text = ""; 
			} 
			txtReqCodeGen.Text = MyActiveLock.get_InstallationCode(txtUser.Text,null);
		}

		private void cmdRegister_Click(object sender, System.EventArgs e)
		{
			//Register this key
			try
			{
                string LibKey = txtLibKeyIn.Text;
                string sUser = txtUser.Text;
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                if( LibKey.Substring(5, 1) == "-" & LibKey.Substring(10,1) == "-" & LibKey.Substring(15, 1) == "-" &
                    LibKey.Substring(20,1) == "-")
                    MyActiveLock.Register(LibKey, ref sUser); //YOU MUST SPECIFY THE USER NAME WITH SHORT KEYS !!!
                else //    ' ALCRYPTO RSA
                {
                    sUser = "";                    
                    MyActiveLock.Register(LibKey, ref sUser);
                }
        
				string userStr;
				userStr = txtUser.Text;
				//MyActiveLock.Register(txtLibKeyIn.Text, ref userStr); 
				MessageBox.Show(modMain.Dec("386.457.46D.483.4F1.4FC.4E6.42B.4FC.483.4C5.4BA.160.4F1.507.441.441.457.4F1.4F1.462.507.4A4.16B"), modMain.MSGBOXCAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information); 
				frmMain_Load(this, new System.EventArgs()); 
				this.Visible = true; 
			}
			catch (SystemException ex)
			{
				MessageBox.Show(ex.Message + ": " + ex.StackTrace, modMain.MSGBOXCAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void cmdKillLicense_Click(object sender, System.EventArgs e)
		{
            // Kill the License File
            // Let Activelock handle this
            // This is intended for developer use only
            MyActiveLock.KillLicense(MyActiveLock.SoftwareName + MyActiveLock.SoftwareVersion, strKeyStorePath);
			MessageBox.Show("Your license has been killed." + "\r\n" + "You need to get a new license for this application if you want to use it.", modMain.MSGBOXCAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information); 
			txtUsedDays.Text = ""; 
			txtExpiration.Text = ""; 
			txtRegisteredLevel.Text = ""; 
            //txtNetworkedLicense.Text = "";
            //txtMaxCount.Text = "";

            /*
			string licFile; 
			licFile = strKeyStorePath; 
			if (File.Exists(licFile)) 
			{ 
				FileInfo fi = new FileInfo(licFile);
				if (fi!=null) 
				{ 
					File.Delete(licFile);
					MessageBox.Show("Your license has been killed." + "\r\n" + "You need to get a new license for this application if you want to use it.", modMain.MSGBOXCAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information); 
					txtUsedDays.Text = ""; 
					txtExpiration.Text = ""; 
					txtRegisteredLevel.Text = ""; 
				} 
				else 
				{ 
					MessageBox.Show("There's no license to kill.", modMain.MSGBOXCAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information); 
				} 
			} 
			else 
			{ 
				MessageBox.Show("There's no license to kill.", modMain.MSGBOXCAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information); 
			} */
			frmMain_Load(this, new System.EventArgs()); 
			cmdResetTrial.Visible = true;
		}

		private void cmdResetTrial_Click(object sender, System.EventArgs e)
		{
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor; 
			MyActiveLock.ResetTrial(); 
			MyActiveLock.ResetTrial(); // DO NOT REMOVE, NEED TO CALL TWICE
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default; 
			MessageBox.Show("Free Trial has been Reset." + "\r\n" + "You'll need to restart the application for a new Free Trial.", modMain.MSGBOXCAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information); 
			txtRegStatus.Text = "Free Trial has been Reset"; 
			txtUsedDays.Text = ""; 
			txtExpiration.Text = ""; 
			txtRegisteredLevel.Text = ""; 
			txtLicenseType.Text = "None";
		}

		private void cmdKillTrial_Click(object sender, System.EventArgs e)
		{
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor; 
			MyActiveLock.KillTrial(); 
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default; 
			MessageBox.Show("Free Trial has been Killed." + "\r\n" + "There will be no more Free Trial next time you start this application." + "\r\n\r\n" + "You must register this application for further use.", modMain.MSGBOXCAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information); 
			txtRegStatus.Text = "Free Trial has been Killed"; 
			txtUsedDays.Text = ""; 
			txtExpiration.Text = ""; 
			txtRegisteredLevel.Text = ""; 
			txtLicenseType.Text = "None";
		}

		private void txtLibKeyIn_TextChanged(object sender, System.EventArgs e)
		{
			cmdRegister.Enabled = (txtLibKeyIn.Text.Trim().Length>0);
		}

		#endregion


	}
}

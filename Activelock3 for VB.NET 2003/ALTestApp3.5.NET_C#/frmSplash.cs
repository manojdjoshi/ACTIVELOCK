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
using System.ComponentModel;
using System.Windows.Forms;

namespace Altestapp35net
{

	internal class frmSplash : Form
	{
		private System.ComponentModel.Container components = null;
		internal System.Windows.Forms.Label Label16;
		internal System.Windows.Forms.PictureBox Image1;
		internal System.Windows.Forms.Label lblInfo;

		public frmSplash()
		{
			this.InitializeComponent();
		}

		protected override void Dispose(bool Disposing)
		{
			if (Disposing && (this.components != null))
			{
				this.components.Dispose();
			}
			base.Dispose(Disposing);
		}

		private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmSplash));
			this.Label16 = new System.Windows.Forms.Label();
			this.Image1 = new System.Windows.Forms.PictureBox();
			this.lblInfo = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// Label16
			// 
			this.Label16.AutoSize = true;
			this.Label16.BackColor = System.Drawing.Color.Transparent;
			this.Label16.Cursor = System.Windows.Forms.Cursors.Default;
			this.Label16.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Label16.ForeColor = System.Drawing.Color.Blue;
			this.Label16.Location = new System.Drawing.Point(83, 132);
			this.Label16.Name = "Label16";
			this.Label16.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.Label16.Size = new System.Drawing.Size(82, 16);
			this.Label16.TabIndex = 4;
			this.Label16.Text = "Activelock3NET";
			this.Label16.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			// 
			// Image1
			// 
			this.Image1.Cursor = System.Windows.Forms.Cursors.Default;
			this.Image1.Image = ((System.Drawing.Image)(resources.GetObject("Image1.Image")));
			this.Image1.Location = new System.Drawing.Point(96, 72);
			this.Image1.Name = "Image1";
			this.Image1.Size = new System.Drawing.Size(55, 55);
			this.Image1.TabIndex = 5;
			this.Image1.TabStop = false;
			// 
			// lblInfo
			// 
			this.lblInfo.BackColor = System.Drawing.Color.Transparent;
			this.lblInfo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblInfo.Cursor = System.Windows.Forms.Cursors.Default;
			this.lblInfo.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblInfo.ForeColor = System.Drawing.SystemColors.WindowText;
			this.lblInfo.Location = new System.Drawing.Point(0, 0);
			this.lblInfo.Name = "lblInfo";
			this.lblInfo.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.lblInfo.Size = new System.Drawing.Size(249, 150);
			this.lblInfo.TabIndex = 3;
			this.lblInfo.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			// 
			// frmSplash
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(192)), ((System.Byte)(255)));
			this.ClientSize = new System.Drawing.Size(249, 150);
			this.ControlBox = false;
			this.Controls.Add(this.Label16);
			this.Controls.Add(this.Image1);
			this.Controls.Add(this.lblInfo);
			this.Cursor = System.Windows.Forms.Cursors.Default;
			this.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
			this.Location = new System.Drawing.Point(83, 132);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmSplash";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.ResumeLayout(false);

		}

	}
}


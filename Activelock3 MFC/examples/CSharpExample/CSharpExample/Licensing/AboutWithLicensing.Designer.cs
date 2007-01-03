namespace Licensing
{
  partial class AboutWithLicensing
  {
    /// <summary>
    /// Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    /// Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing)
    {
      if (disposing && (components != null))
      {
        components.Dispose();
      }
      base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
      this.groupBox1 = new System.Windows.Forms.GroupBox();
      this.label2 = new System.Windows.Forms.Label();
      this.label1 = new System.Windows.Forms.Label();
      this.userName = new System.Windows.Forms.TextBox();
      this.installationCode = new System.Windows.Forms.TextBox();
      this.generate = new System.Windows.Forms.Button();
      this.explain = new System.Windows.Forms.TextBox();
      this.licenseInformation = new System.Windows.Forms.GroupBox();
      this.label6 = new System.Windows.Forms.Label();
      this.label5 = new System.Windows.Forms.Label();
      this.label4 = new System.Windows.Forms.Label();
      this.label3 = new System.Windows.Forms.Label();
      this.expiryDateTextBox = new System.Windows.Forms.TextBox();
      this.daysUsedTextBox = new System.Windows.Forms.TextBox();
      this.registeredDateTextBox = new System.Windows.Forms.TextBox();
      this.statusTextBox = new System.Windows.Forms.TextBox();
      this.Ok = new System.Windows.Forms.Button();
      this.Cancel = new System.Windows.Forms.Button();
      this.label7 = new System.Windows.Forms.Label();
      this.groupBox1.SuspendLayout();
      this.licenseInformation.SuspendLayout();
      this.SuspendLayout();
      // 
      // groupBox1
      // 
      this.groupBox1.Controls.Add(this.label2);
      this.groupBox1.Controls.Add(this.label1);
      this.groupBox1.Controls.Add(this.userName);
      this.groupBox1.Controls.Add(this.installationCode);
      this.groupBox1.Controls.Add(this.generate);
      this.groupBox1.Location = new System.Drawing.Point(13, 60);
      this.groupBox1.Name = "groupBox1";
      this.groupBox1.Size = new System.Drawing.Size(733, 105);
      this.groupBox1.TabIndex = 0;
      this.groupBox1.TabStop = false;
      this.groupBox1.Text = "User Data";
      // 
      // label2
      // 
      this.label2.AutoSize = true;
      this.label2.Location = new System.Drawing.Point(10, 72);
      this.label2.Name = "label2";
      this.label2.Size = new System.Drawing.Size(60, 13);
      this.label2.TabIndex = 4;
      this.label2.Text = "User Name";
      // 
      // label1
      // 
      this.label1.AutoSize = true;
      this.label1.Location = new System.Drawing.Point(7, 45);
      this.label1.Name = "label1";
      this.label1.Size = new System.Drawing.Size(85, 13);
      this.label1.TabIndex = 3;
      this.label1.Text = "Installation Code";
      // 
      // userName
      // 
      this.userName.Location = new System.Drawing.Point(110, 66);
      this.userName.Name = "userName";
      this.userName.Size = new System.Drawing.Size(609, 20);
      this.userName.TabIndex = 2;
      this.userName.Tag = "";
      // 
      // installationCode
      // 
      this.installationCode.Location = new System.Drawing.Point(110, 39);
      this.installationCode.Name = "installationCode";
      this.installationCode.ReadOnly = true;
      this.installationCode.Size = new System.Drawing.Size(609, 20);
      this.installationCode.TabIndex = 1;
      // 
      // generate
      // 
      this.generate.Location = new System.Drawing.Point(110, 10);
      this.generate.Name = "generate";
      this.generate.Size = new System.Drawing.Size(75, 23);
      this.generate.TabIndex = 0;
      this.generate.Text = "Generate";
      this.generate.UseVisualStyleBackColor = true;
      this.generate.Click += new System.EventHandler(this.generate_Click);
      // 
      // explain
      // 
      this.explain.Location = new System.Drawing.Point(13, 185);
      this.explain.Multiline = true;
      this.explain.Name = "explain";
      this.explain.ReadOnly = true;
      this.explain.Size = new System.Drawing.Size(733, 98);
      this.explain.TabIndex = 1;
      // 
      // licenseInformation
      // 
      this.licenseInformation.Controls.Add(this.label6);
      this.licenseInformation.Controls.Add(this.label5);
      this.licenseInformation.Controls.Add(this.label4);
      this.licenseInformation.Controls.Add(this.label3);
      this.licenseInformation.Controls.Add(this.expiryDateTextBox);
      this.licenseInformation.Controls.Add(this.daysUsedTextBox);
      this.licenseInformation.Controls.Add(this.registeredDateTextBox);
      this.licenseInformation.Controls.Add(this.statusTextBox);
      this.licenseInformation.Location = new System.Drawing.Point(123, 301);
      this.licenseInformation.Name = "licenseInformation";
      this.licenseInformation.Size = new System.Drawing.Size(533, 104);
      this.licenseInformation.TabIndex = 2;
      this.licenseInformation.TabStop = false;
      this.licenseInformation.Text = "License Information";
      // 
      // label6
      // 
      this.label6.AutoSize = true;
      this.label6.Location = new System.Drawing.Point(287, 61);
      this.label6.Name = "label6";
      this.label6.Size = new System.Drawing.Size(61, 13);
      this.label6.TabIndex = 7;
      this.label6.Text = "Expiry Date";
      // 
      // label5
      // 
      this.label5.AutoSize = true;
      this.label5.Location = new System.Drawing.Point(287, 34);
      this.label5.Name = "label5";
      this.label5.Size = new System.Drawing.Size(84, 13);
      this.label5.TabIndex = 6;
      this.label5.Text = "Registered Date";
      // 
      // label4
      // 
      this.label4.AutoSize = true;
      this.label4.Location = new System.Drawing.Point(39, 58);
      this.label4.Name = "label4";
      this.label4.Size = new System.Drawing.Size(59, 13);
      this.label4.TabIndex = 5;
      this.label4.Text = "Days Used";
      // 
      // label3
      // 
      this.label3.AutoSize = true;
      this.label3.Location = new System.Drawing.Point(36, 27);
      this.label3.Name = "label3";
      this.label3.Size = new System.Drawing.Size(37, 13);
      this.label3.TabIndex = 4;
      this.label3.Text = "Status";
      // 
      // expiryDateTextBox
      // 
      this.expiryDateTextBox.Location = new System.Drawing.Point(390, 58);
      this.expiryDateTextBox.Name = "expiryDateTextBox";
      this.expiryDateTextBox.ReadOnly = true;
      this.expiryDateTextBox.Size = new System.Drawing.Size(100, 20);
      this.expiryDateTextBox.TabIndex = 3;
      // 
      // daysUsedTextBox
      // 
      this.daysUsedTextBox.Location = new System.Drawing.Point(129, 58);
      this.daysUsedTextBox.Name = "daysUsedTextBox";
      this.daysUsedTextBox.ReadOnly = true;
      this.daysUsedTextBox.Size = new System.Drawing.Size(100, 20);
      this.daysUsedTextBox.TabIndex = 2;
      // 
      // registeredDateTextBox
      // 
      this.registeredDateTextBox.Location = new System.Drawing.Point(390, 27);
      this.registeredDateTextBox.Name = "registeredDateTextBox";
      this.registeredDateTextBox.ReadOnly = true;
      this.registeredDateTextBox.Size = new System.Drawing.Size(100, 20);
      this.registeredDateTextBox.TabIndex = 1;
      // 
      // statusTextBox
      // 
      this.statusTextBox.Location = new System.Drawing.Point(129, 27);
      this.statusTextBox.Name = "statusTextBox";
      this.statusTextBox.ReadOnly = true;
      this.statusTextBox.Size = new System.Drawing.Size(100, 20);
      this.statusTextBox.TabIndex = 0;
      // 
      // Ok
      // 
      this.Ok.Location = new System.Drawing.Point(671, 12);
      this.Ok.Name = "Ok";
      this.Ok.Size = new System.Drawing.Size(75, 23);
      this.Ok.TabIndex = 3;
      this.Ok.Text = "Ok";
      this.Ok.UseVisualStyleBackColor = true;
      this.Ok.Click += new System.EventHandler(this.Ok_Click);
      // 
      // Cancel
      // 
      this.Cancel.Location = new System.Drawing.Point(563, 12);
      this.Cancel.Name = "Cancel";
      this.Cancel.Size = new System.Drawing.Size(75, 23);
      this.Cancel.TabIndex = 4;
      this.Cancel.Text = "Cancel";
      this.Cancel.UseVisualStyleBackColor = true;
      this.Cancel.Click += new System.EventHandler(this.Cancel_Click);
      // 
      // label7
      // 
      this.label7.AutoSize = true;
      this.label7.Location = new System.Drawing.Point(26, 5);
      this.label7.Name = "label7";
      this.label7.Size = new System.Drawing.Size(85, 13);
      this.label7.TabIndex = 5;
      this.label7.Text = "CSharp Example";
      // 
      // AboutWithLicensing
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(772, 435);
      this.ControlBox = false;
      this.Controls.Add(this.label7);
      this.Controls.Add(this.Cancel);
      this.Controls.Add(this.Ok);
      this.Controls.Add(this.licenseInformation);
      this.Controls.Add(this.explain);
      this.Controls.Add(this.groupBox1);
      this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
      this.MaximizeBox = false;
      this.MinimizeBox = false;
      this.Name = "AboutWithLicensing";
      this.Text = "AboutWithLicensing";
      this.Load += new System.EventHandler(this.Awl_Load);
      this.groupBox1.ResumeLayout(false);
      this.groupBox1.PerformLayout();
      this.licenseInformation.ResumeLayout(false);
      this.licenseInformation.PerformLayout();
      this.ResumeLayout(false);
      this.PerformLayout();

    }

    #endregion

    private System.Windows.Forms.GroupBox groupBox1;
    private System.Windows.Forms.Label label1;
    private System.Windows.Forms.TextBox userName;
    private System.Windows.Forms.TextBox installationCode;
    private System.Windows.Forms.Button generate;
    private System.Windows.Forms.Label label2;
    private System.Windows.Forms.TextBox explain;
    private System.Windows.Forms.GroupBox licenseInformation;
    private System.Windows.Forms.Label label3;
    private System.Windows.Forms.TextBox expiryDateTextBox;
    private System.Windows.Forms.TextBox daysUsedTextBox;
    private System.Windows.Forms.TextBox registeredDateTextBox;
    private System.Windows.Forms.TextBox statusTextBox;
    private System.Windows.Forms.Label label6;
    private System.Windows.Forms.Label label5;
    private System.Windows.Forms.Label label4;
    private System.Windows.Forms.Button Ok;
    private System.Windows.Forms.Button Cancel;
    private System.Windows.Forms.Label label7;

  }
}
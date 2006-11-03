namespace CSharpExample
{
  partial class Form1
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
      System.Windows.Forms.ListViewItem listViewItem1 = new System.Windows.Forms.ListViewItem(new string[] {
            "ListView",
            "Enabled",
            "For Registered or In Trial Mode"}, -1);
      System.Windows.Forms.ListViewItem listViewItem2 = new System.Windows.Forms.ListViewItem(new string[] {
            "ListView",
            "Diasabled",
            "Not Registered"}, -1);
      this.listView1 = new System.Windows.Forms.ListView();
      this.columnHeader1 = new System.Windows.Forms.ColumnHeader();
      this.columnHeader2 = new System.Windows.Forms.ColumnHeader();
      this.columnHeader3 = new System.Windows.Forms.ColumnHeader();
      this.menuStrip1 = new System.Windows.Forms.MenuStrip();
      this.licensingToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
      this.licensingToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
      this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
      this.label1 = new System.Windows.Forms.Label();
      this.importantBtn = new System.Windows.Forms.Button();
      this.normalBtn = new System.Windows.Forms.Button();
      this.panel1 = new System.Windows.Forms.Panel();
      this.checkBox2 = new System.Windows.Forms.CheckBox();
      this.checkBox1 = new System.Windows.Forms.CheckBox();
      this.radioButton2 = new System.Windows.Forms.RadioButton();
      this.radioButton1 = new System.Windows.Forms.RadioButton();
      this.label2 = new System.Windows.Forms.Label();
      this.label3 = new System.Windows.Forms.Label();
      this.label4 = new System.Windows.Forms.Label();
      this.label5 = new System.Windows.Forms.Label();
      this.menuStrip1.SuspendLayout();
      this.panel1.SuspendLayout();
      this.SuspendLayout();
      // 
      // listView1
      // 
      this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader2,
            this.columnHeader3});
      this.listView1.FullRowSelect = true;
      this.listView1.HideSelection = false;
      this.listView1.Items.AddRange(new System.Windows.Forms.ListViewItem[] {
            listViewItem1,
            listViewItem2});
      this.listView1.Location = new System.Drawing.Point(33, 148);
      this.listView1.Name = "listView1";
      this.listView1.Size = new System.Drawing.Size(448, 181);
      this.listView1.TabIndex = 0;
      this.listView1.UseCompatibleStateImageBehavior = false;
      this.listView1.View = System.Windows.Forms.View.Details;
      // 
      // columnHeader1
      // 
      this.columnHeader1.Text = "Important";
      this.columnHeader1.Width = 94;
      // 
      // columnHeader2
      // 
      this.columnHeader2.Text = "Interesting";
      this.columnHeader2.Width = 120;
      // 
      // columnHeader3
      // 
      this.columnHeader3.Text = "Boring";
      this.columnHeader3.Width = 194;
      // 
      // menuStrip1
      // 
      this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.licensingToolStripMenuItem});
      this.menuStrip1.Location = new System.Drawing.Point(0, 0);
      this.menuStrip1.Name = "menuStrip1";
      this.menuStrip1.Size = new System.Drawing.Size(516, 24);
      this.menuStrip1.TabIndex = 1;
      this.menuStrip1.Text = "menuStrip1";
      // 
      // licensingToolStripMenuItem
      // 
      this.licensingToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.licensingToolStripMenuItem1,
            this.exitToolStripMenuItem});
      this.licensingToolStripMenuItem.Name = "licensingToolStripMenuItem";
      this.licensingToolStripMenuItem.Size = new System.Drawing.Size(72, 20);
      this.licensingToolStripMenuItem.Text = "Licensing";
      // 
      // licensingToolStripMenuItem1
      // 
      this.licensingToolStripMenuItem1.Name = "licensingToolStripMenuItem1";
      this.licensingToolStripMenuItem1.Size = new System.Drawing.Size(152, 22);
      this.licensingToolStripMenuItem1.Text = "Licensing";
      this.licensingToolStripMenuItem1.Click += new System.EventHandler(this.licensingToolStripMenuItem1_Click);
      // 
      // exitToolStripMenuItem
      // 
      this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
      this.exitToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
      this.exitToolStripMenuItem.Text = "Exit";
      this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
      // 
      // label1
      // 
      this.label1.AutoSize = true;
      this.label1.Location = new System.Drawing.Point(181, 132);
      this.label1.Name = "label1";
      this.label1.Size = new System.Drawing.Size(127, 13);
      this.label1.TabIndex = 2;
      this.label1.Text = "Extremely Stupid ListView";
      // 
      // importantBtn
      // 
      this.importantBtn.Location = new System.Drawing.Point(100, 375);
      this.importantBtn.Name = "importantBtn";
      this.importantBtn.Size = new System.Drawing.Size(75, 23);
      this.importantBtn.TabIndex = 3;
      this.importantBtn.Text = "Important";
      this.importantBtn.UseVisualStyleBackColor = true;
      // 
      // normalBtn
      // 
      this.normalBtn.Location = new System.Drawing.Point(338, 375);
      this.normalBtn.Name = "normalBtn";
      this.normalBtn.Size = new System.Drawing.Size(75, 23);
      this.normalBtn.TabIndex = 4;
      this.normalBtn.Text = "Normal";
      this.normalBtn.UseVisualStyleBackColor = true;
      // 
      // panel1
      // 
      this.panel1.Controls.Add(this.checkBox2);
      this.panel1.Controls.Add(this.checkBox1);
      this.panel1.Controls.Add(this.radioButton2);
      this.panel1.Controls.Add(this.radioButton1);
      this.panel1.Location = new System.Drawing.Point(33, 51);
      this.panel1.Name = "panel1";
      this.panel1.Size = new System.Drawing.Size(448, 66);
      this.panel1.TabIndex = 5;
      // 
      // checkBox2
      // 
      this.checkBox2.AutoSize = true;
      this.checkBox2.Location = new System.Drawing.Point(298, 37);
      this.checkBox2.Name = "checkBox2";
      this.checkBox2.Size = new System.Drawing.Size(80, 17);
      this.checkBox2.TabIndex = 3;
      this.checkBox2.Text = "checkBox2";
      this.checkBox2.UseVisualStyleBackColor = true;
      // 
      // checkBox1
      // 
      this.checkBox1.AutoSize = true;
      this.checkBox1.Location = new System.Drawing.Point(298, 13);
      this.checkBox1.Name = "checkBox1";
      this.checkBox1.Size = new System.Drawing.Size(80, 17);
      this.checkBox1.TabIndex = 2;
      this.checkBox1.Text = "checkBox1";
      this.checkBox1.UseVisualStyleBackColor = true;
      // 
      // radioButton2
      // 
      this.radioButton2.AutoSize = true;
      this.radioButton2.Location = new System.Drawing.Point(18, 37);
      this.radioButton2.Name = "radioButton2";
      this.radioButton2.Size = new System.Drawing.Size(85, 17);
      this.radioButton2.TabIndex = 1;
      this.radioButton2.TabStop = true;
      this.radioButton2.Text = "radioButton2";
      this.radioButton2.UseVisualStyleBackColor = true;
      // 
      // radioButton1
      // 
      this.radioButton1.AutoSize = true;
      this.radioButton1.Location = new System.Drawing.Point(18, 13);
      this.radioButton1.Name = "radioButton1";
      this.radioButton1.Size = new System.Drawing.Size(85, 17);
      this.radioButton1.TabIndex = 0;
      this.radioButton1.TabStop = true;
      this.radioButton1.Text = "radioButton1";
      this.radioButton1.UseVisualStyleBackColor = true;
      // 
      // label2
      // 
      this.label2.AutoSize = true;
      this.label2.Location = new System.Drawing.Point(181, 35);
      this.label2.Name = "label2";
      this.label2.Size = new System.Drawing.Size(164, 13);
      this.label2.TabIndex = 6;
      this.label2.Text = "Some Buttons Which Do Nothing";
      // 
      // label3
      // 
      this.label3.AutoSize = true;
      this.label3.Location = new System.Drawing.Point(78, 346);
      this.label3.Name = "label3";
      this.label3.Size = new System.Drawing.Size(150, 13);
      this.label3.TabIndex = 7;
      this.label3.Text = "Enabled Fo Registered or Trial";
      // 
      // label4
      // 
      this.label4.AutoSize = true;
      this.label4.Location = new System.Drawing.Point(78, 359);
      this.label4.Name = "label4";
      this.label4.Size = new System.Drawing.Size(134, 13);
      this.label4.TabIndex = 8;
      this.label4.Text = "Disabled For UnRegistered";
      // 
      // label5
      // 
      this.label5.AutoSize = true;
      this.label5.Location = new System.Drawing.Point(301, 359);
      this.label5.Name = "label5";
      this.label5.Size = new System.Drawing.Size(180, 13);
      this.label5.TabIndex = 9;
      this.label5.Text = "Unimportant So Enabled In All Cases";
      // 
      // Form1
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(516, 411);
      this.Controls.Add(this.label5);
      this.Controls.Add(this.label4);
      this.Controls.Add(this.label3);
      this.Controls.Add(this.label2);
      this.Controls.Add(this.panel1);
      this.Controls.Add(this.normalBtn);
      this.Controls.Add(this.importantBtn);
      this.Controls.Add(this.label1);
      this.Controls.Add(this.menuStrip1);
      this.Controls.Add(this.listView1);
      this.MainMenuStrip = this.menuStrip1;
      this.Name = "Form1";
      this.Text = "Form1";
      this.menuStrip1.ResumeLayout(false);
      this.menuStrip1.PerformLayout();
      this.panel1.ResumeLayout(false);
      this.panel1.PerformLayout();
      this.ResumeLayout(false);
      this.PerformLayout();

    }

    #endregion

    private System.Windows.Forms.ListView listView1;
    private System.Windows.Forms.ColumnHeader columnHeader1;
    private System.Windows.Forms.ColumnHeader columnHeader2;
    private System.Windows.Forms.ColumnHeader columnHeader3;
    private System.Windows.Forms.MenuStrip menuStrip1;
    private System.Windows.Forms.ToolStripMenuItem licensingToolStripMenuItem;
    private System.Windows.Forms.ToolStripMenuItem licensingToolStripMenuItem1;
    private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
    private System.Windows.Forms.Label label1;
    private System.Windows.Forms.Button importantBtn;
    private System.Windows.Forms.Button normalBtn;
    private System.Windows.Forms.Panel panel1;
    private System.Windows.Forms.CheckBox checkBox2;
    private System.Windows.Forms.CheckBox checkBox1;
    private System.Windows.Forms.RadioButton radioButton2;
    private System.Windows.Forms.RadioButton radioButton1;
    private System.Windows.Forms.Label label2;
    private System.Windows.Forms.Label label3;
    private System.Windows.Forms.Label label4;
    private System.Windows.Forms.Label label5;
  }
}


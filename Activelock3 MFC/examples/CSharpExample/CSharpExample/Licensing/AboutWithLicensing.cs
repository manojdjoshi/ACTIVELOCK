using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using ActiveLock3_5NET;

namespace Licensing
{
  public partial class AboutWithLicensing : Form
  {
    
    private  _IActiveLock activeLock;
    public AboutWithLicensing(_IActiveLock _activeLock)
    {
      InitializeComponent();
      activeLock = _activeLock;
    }

    public string UserInput
    {
      get
      {
        return this.userName.Text;
      }
      set
      {
        this.userName.Text = value;
      }
    }

    public string Status
    {
      get
      {
        return this.statusTextBox.Text;
      }
      set
      {
        this.statusTextBox.Text = value;
      }
    }

    public string DaysUsed
    {
      get
      {
        return this.daysUsedTextBox.Text;
      }
      set
      {
        this.daysUsedTextBox.Text = value;
      }
    }

    public string RegisteredDate
    {
      get
      {
        return this.registeredDateTextBox.Text;
      }
      set
      {
        this.registeredDateTextBox.Text = value;
      }
    }

    public string ExpiryDate
    {
      get
      {
        return this.expiryDateTextBox.Text;
      }
      set
      {
        this.expiryDateTextBox.Text = value;
      }
    }

    public string Explain
    {
      get
      {
        return this.explain.Text;
      }
      set
      {
        this.explain.Text = value;
      }
    }

    public string InstallationCode
    {
      get
      {
        return this.installationCode.Text;
      }
      set
      {
        this.installationCode.Text = value;
      }
    }

    private void Cancel_Click(object sender, EventArgs e)
    {
      DialogResult = DialogResult.Cancel;
    }

    private void Ok_Click(object sender, EventArgs e)
    {
      DialogResult = DialogResult.OK;
    }

    private void generate_Click(object sender, EventArgs e)
    {
      InstallationCode = activeLock.get_InstallationCode(UserInput, null);
    }

    private void Awl_Load(object sender, EventArgs e)
    {
      if (Status == "Registered")
      {
        this.generate.Enabled = false;
        this.userName.Enabled = false;
      }
      else
      {
      }
    }

  }
}
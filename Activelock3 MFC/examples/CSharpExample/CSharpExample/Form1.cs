using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using ActiveLock3_5NET;
using Licensing;

namespace CSharpExample
{
  public partial class Form1 : Form
  {
    _IActiveLock activeLock;
    private string installationCode;

    public string InstallationCode
    {
      get { return installationCode; }
      set { installationCode = value; }
    }
    private string expiry;

    public string Expiry
    {
      get { return expiry; }
      set { expiry = value; }
    }
    private int usedDays;

    public int UsedDays
    {
      get { return usedDays; }
      set { usedDays = value; }
    }
    private string registeredUser;

    public string RegisteredUser
    {
      get { return registeredUser; }
      set { registeredUser = value; }
    }
    private string registeredDate;

    public string RegisteredDate
    {
      get { return registeredDate; }
      set { registeredDate = value; }
    }
    private IActiveLock.ALLockTypes lockType;

    public IActiveLock.ALLockTypes LockType
    {
      get { return lockType; }
      set { lockType = value; }
    }

    private string versionNr;

    public string VersionNr
    {
      get { return versionNr; }
      set { versionNr = value; }
    }
    private bool regStatus;

    public bool RegStatus
    {
      get { return regStatus; }
      set { regStatus = value; }
    }
    private string licenseStatus;

    public string LicenseStatus
    {
      get { return licenseStatus; }
      set { licenseStatus = value; }
    }
    private bool trial;

    public bool Trial
    {
      get { return trial; }
      set { trial = value; }
    }

    
    public Form1()
    {
      InitializeComponent();
      ActiveLockAccess();
    }

    private void exitToolStripMenuItem_Click(object sender, EventArgs e)
    {

    }

    private void licensingToolStripMenuItem1_Click(object sender, EventArgs e)
    {
      AboutWithLicensing awl = new AboutWithLicensing(activeLock);                // create dialog
      awl.InstallationCode = InstallationCode;
      awl.UserInput = RegisteredUser;

      awl.RegisteredDate = RegisteredDate;
      awl.Status = LicenseStatus;
      awl.ExpiryDate = Expiry;
      awl.DaysUsed = UsedDays.ToString();

      if (awl.ShowDialog() == DialogResult.Yes)
      {
      }

    }


    private void ActiveLockAccess()
    {
      string result = "see if it changes";
      try
      {
        ActiveLock3_5NET.Globals_Renamed glob = new Globals_Renamed();
        activeLock = glob.NewInstance();
        activeLock.SoftwareName = "TestApp";
        activeLock.SoftwareVersion = "1.0";
        activeLock.LockType = IActiveLock.ALLockTypes.lockHDFirmware;
        activeLock.SoftwareCode = "AAAAB3NzaC1yc2EAAAABJQAAAIB8/B2KWoai2WSGTRPcgmMoczeXpd8nv0Y4r1sJ1wV3vH21q4rTpEYuBiD4HFOpkbNBSRdpBHJGWec7jUi8ISV0pM6i2KznjhCms5CEtYHRybbiYvRXleGzFsAAP817PLN3JYo3WkErT2ofR5RCkfhmx060BT8waPoqnn3AB7sZ0Q==";
        activeLock.KeyStoreType = IActiveLock.LicStoreType.alsFile;
        activeLock.KeyStorePath = "c:\\S3.lic";
        activeLock.AutoRegisterKeyPath = "c:\\TestApp.all";
        activeLock.SoftwarePassword = "something better that this";

        activeLock.Init("", ref result);

        activeLock.TrialType = IActiveLock.ALTrialTypes.trialDays;
        activeLock.TrialHideType = IActiveLock.ALTrialHideTypes.trialSteganography;
        activeLock.TrialLength = 31;
        //activeLock.ResetTrial();
        //activeLock.ResetTrial();

        activeLock.Acquire(ref result);

        try
        {
          UsedDays = activeLock.UsedDays;
          VersionNr = activeLock.SoftwareVersion;
          RegisteredDate = activeLock.RegisteredDate;
          RegisteredUser = activeLock.RegisteredUser;
          InstallationCode = activeLock.get_InstallationCode(RegisteredUser, null);
          LockType = activeLock.LockType;                       // Activelock index 
          Expiry = activeLock.ExpirationDate;
          // The above could be a combination of codes - this case is not handled here - future job
          // in fact a combination can not be a legal enum type-my fault should report to forum
          // activeLock.UsedLockType = LockType;                         // make them all agree

          RegStatus = true;
          LicenseStatus = "Registered";
        }
        catch
        {
          // must (I hope) be a trial license 
          RegStatus = true;
          Trial = true;
          LicenseStatus = "Trial License";
        }
      }
      catch
      {
        RegStatus = false;
        Trial = false;
        LicenseStatus = "Not Registered";
      }
    }


  }
}
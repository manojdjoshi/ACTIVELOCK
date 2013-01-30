using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
public class FingerPrint
{

	private string m_CpuID = "";
	private string m_BiosID = "";
	private string m_DiskID = "";
	private string m_BaseID = "";
	private string m_VideoID = "";

	private string m_MacID = "";
	private bool m_UseCpuID = true;
	private bool m_UseBiosID = true;
	private bool m_UseDiskID = true;
	private bool m_UseBaseID = true;
	private bool m_UseVideoID = true;

	private bool m_UseMacID = true;
	private long m_ReturnLength = 8;

	private long m_TotalLength = 8;
	public event StartingWithEventHandler StartingWith;
	public delegate void StartingWithEventHandler(string Text);
	public event DoneWithEventHandler DoneWith;
	public delegate void DoneWithEventHandler(string Text);

	public long TotalLength {
		get { return m_TotalLength; }
	}

	public long ReturnLength {
		get { return m_ReturnLength; }
		set {
			if (m_ReturnLength < 0) m_ReturnLength = 0; 
			m_ReturnLength = value;
		}
	}

	public bool UseCpuID {
		get { return m_UseCpuID; }
		set { m_UseCpuID = value; }
	}

	public bool UseBiosID {
		get { return m_UseBiosID; }
		set { m_UseBiosID = value; }
	}

	public bool UseDiskID {
		get { return m_UseDiskID; }
		set { m_UseDiskID = value; }
	}

	public bool UseBaseID {
		get { return m_UseBaseID; }
		set { m_UseBaseID = value; }
	}

	public bool UseVideoID {
		get { return m_UseVideoID; }
		set { m_UseVideoID = value; }
	}

	public bool UseMacID {
		get { return m_UseMacID; }
		set { m_UseMacID = value; }
	}

	public string Value {
		get {
			if (StartingWith != null) {
				StartingWith("All");
			}

			m_CpuID = "";
			if (m_UseCpuID) m_CpuID = CpuID(); 
			m_BiosID = "";
			if (m_UseBiosID) m_BiosID = BiosID(); 
			m_DiskID = "";
			if (m_UseDiskID) m_DiskID = DiskID(); 
			m_BaseID = "";
			if (m_UseBaseID) m_BaseID = BaseID(); 
			m_VideoID = "";
			if (m_UseVideoID) m_VideoID = VideoID(); 
			//m_MacID = ""
			//If m_UseMacID Then m_MacID = MacID()

			return Pack(m_CpuID + m_BiosID + m_DiskID + m_BaseID + m_VideoID);
			//& m_MacID)

			if (DoneWith != null) {
				DoneWith("All");
			}
		}
	}

	private string Identifier(string wmiClass, string wmiProperty, string wmiMustBeTrue)
	{
		//Return a hardware identifier

		string Result = "";
		System.Management.ManagementClass mc = new System.Management.ManagementClass(wmiClass);
		System.Management.ManagementObjectCollection moc = mc.GetInstances();
		System.Management.ManagementObject mo = null;

		foreach (System.Management.ManagementObject mo_loopVariable in moc) {
			mo = mo_loopVariable;
			if (mo[wmiMustBeTrue].ToString() == "True") {
				//Only get the first one
				if (string.IsNullOrEmpty(Result)) {
					try {
						Result = mo[wmiProperty].ToString();
						break; // TODO: might not be correct. Was : Exit For
					}
					catch (Exception ex) {
						//Ignore error
					}
				}
			}
		}

		return Result;
	}

	private string Identifier(string wmiClass, string wmiProperty)
	{
		//Return a hardware identifier

		string Result = "";
		System.Management.ManagementClass mc = new System.Management.ManagementClass(wmiClass);
		System.Management.ManagementObjectCollection moc = mc.GetInstances();
		System.Management.ManagementObject mo = null;

		foreach (System.Management.ManagementObject mo_loopVariable in moc) {
			mo = mo_loopVariable;
			//Only get the first one
			if (string.IsNullOrEmpty(Result)) {
				try {
					Result = mo[wmiProperty].ToString();
					break; // TODO: might not be correct. Was : Exit For
				}
				catch (Exception ex) {
					//Ignore error
				}
			}
		}

		return Result;
	}

	private string CpuID()
	{
		if (StartingWith != null) {
			StartingWith("CpuID");
		}

		//Uses first CPU identifier available in order of preference
		//Don't get all identifiers as very time consuming
		// Do not get the following because it's mostly unavailable
		//Dim RetVal As String = Identifier("Win32_Processor", "UniqueId")

		string RetVal = string.Empty;
		//If no UniqueId, use ProcessorID
		if (string.IsNullOrEmpty(RetVal)) {
			RetVal = Identifier("Win32_Processor", "ProcessorId");

			//If no ProcessorID, use Name
			if (string.IsNullOrEmpty(RetVal)) {
				RetVal = Identifier("Win32_Processor", "Name");

				//If no Name, use Manufacturer
				if (string.IsNullOrEmpty(RetVal)) {
					RetVal = Identifier("Win32_Processor", "Manufacturer");
				}

				//Add clock speed for extra security
				RetVal += Identifier("Win32_Processor", "MaxClockSpeed");
			}
		}

		return RetVal;

		if (DoneWith != null) {
			DoneWith("CpuID");
		}
	}

	private string BiosID()
	{
		if (StartingWith != null) {
			StartingWith("BiosID");
		}

		//BIOS Identifier

		return Identifier("Win32_BIOS", "Manufacturer") + Identifier("Win32_BIOS", "SMBIOSBIOSVersion") + Identifier("Win32_BIOS", "SerialNumber") + Identifier("Win32_BIOS", "ReleaseDate") + Identifier("Win32_BIOS", "Version");
		//          & Identifier("Win32_BIOS", "IdentificationCode") _

		if (DoneWith != null) {
			DoneWith("BiosID");
		}
	}

	private string DiskID()
	{
		if (StartingWith != null) {
			StartingWith("DiskID");
		}

		//Main physical hard drive ID

		return Identifier("Win32_DiskDrive", "Manufacturer") + Identifier("Win32_DiskDrive", "Signature") + Identifier("Win32_DiskDrive", "TotalHeads");
		//Identifier("Win32_DiskDrive", "Model") _

		if (DoneWith != null) {
			DoneWith("CpuID");
		}
	}

	private string BaseID()
	{
		if (StartingWith != null) {
			StartingWith("BaseID");
		}

		//Motherboard ID

		return Identifier("Win32_BaseBoard", "Model") + Identifier("Win32_BaseBoard", "Manufacturer") + Identifier("Win32_BaseBoard", "Name") + Identifier("Win32_BaseBoard", "SerialNumber");

		if (DoneWith != null) {
			DoneWith("BaseID");
		}
	}

	private string VideoID()
	{
		if (StartingWith != null) {
			StartingWith("VideoID");
		}

		//Primary video controller ID

		return Identifier("Win32_VideoController", "DriverVersion") + Identifier("Win32_VideoController", "Name");

		if (DoneWith != null) {
			DoneWith("VideoID");
		}
	}

	private string MacID()
	{
		if (StartingWith != null) {
			StartingWith("MacID");
		}

		//First enabled network card ID

		return Identifier("Win32_NetworkAdapterConfiguration", "MACAddress", "IPEnabled");

		if (StartingWith != null) {
			StartingWith("MacID");
		}
	}

	private string Pack(string Text)
	{
		if (StartingWith != null) {
			StartingWith("Packing");
		}

		//Packs the string to m_ReturnLength digits
		//If m_ReturnLength=-1 : Return complete string

		string RetVal = null;
		long X = 0;
		long Y = 0;
		char N = '\0';

		foreach (char N_loopVariable in Text) {
			N = N_loopVariable;
			Y += 1;
			X += (Strings.Asc(N) * Y);
		}

		if (m_ReturnLength > 0) {
			RetVal = X.ToString().PadRight((int)m_ReturnLength, Strings.Chr(48));
		}
		else {
			RetVal = X.ToString();
		}

		if (m_ReturnLength == 0) {
			m_TotalLength = RetVal.Length;
			return RetVal;
		}
		else {
			m_TotalLength = RetVal.Length;
			return RetVal.Substring(0, (int)m_ReturnLength);
		}

		if (DoneWith != null) {
			DoneWith("Packing");
		}
	}
}





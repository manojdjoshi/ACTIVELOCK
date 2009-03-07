using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;

/// <summary>
/// clsNetworkStats - Needs updates to comments!
/// </summary>
/// <remarks></remarks>
class clsNetworkStats
{

	#region " DECLARES "
	private const long ERROR_SUCCESS = 0;
	private const long MAX_INTERFACE_NAME_LEN = 256;
	private const long MAXLEN_IFDESCR = 256;

	private const long MAXLEN_PHYSADDR = 8;
	/// <summary>
	/// MIB_IFROW - The MIB_IFROW structure stores information about a particular interface.
	/// </summary>
	/// <remarks>See http://msdn.microsoft.com/en-us/library/aa366836(VS.85).aspx for Full Documentation!</remarks>
	[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
	public struct MIB_IFROW
	{
		/// <summary>
		/// wszName - A pointer to a Unicode string that contains the name of the interface.
		/// </summary>
		/// <remarks></remarks>
		[MarshalAs(UnmanagedType.ByValTStr, sizeconst = clsNetworkStats.MAX_INTERFACE_NAME_LEN)]
		public string wszName;
		/// <summary>
		/// dwIndex - The index that identifies the interface. This index value may change when a network adapter is disabled and then enabled, and should not be considered persistent.
		/// </summary>
		/// <remarks></remarks>
		public UInt32 dwIndex;
		/// <summary>
		/// dwType = The interface type as defined by the Internet Assigned Names Authority (IANA). For more information, see <a href="http://www.iana.org/assignments/ianaiftype-mib">http://www.iana.org/assignments/ianaiftype-mib</a>. Possible values for the interface type are listed in the Ipifcons.h header file. 
		/// </summary>
		/// <remarks>See http://msdn.microsoft.com/en-us/library/aa366836(VS.85).aspx for more info!</remarks>
		public UInt32 dwType;
		/// <summary>
		/// dwMtu - The Maximum Transmission Unit (MTU) size in bytes.
		/// </summary>
		/// <remarks></remarks>
		public UInt32 dwMtu;
		/// <summary>
		/// dwSpeed - The speed of the interface in bits per second.
		/// </summary>
		/// <remarks></remarks>
		public UInt32 dwSpeed;
		/// <summary>
		/// dwPhysAddrLen - The length, in bytes, of the physical address specified by the bPhysAddr member.
		/// </summary>
		/// <remarks></remarks>
		public UInt32 dwPhysAddrLen;
		/// <summary>
		/// bPhysAddr - The physical address of the adapter for this interface.
		/// </summary>
		/// <remarks></remarks>
		[MarshalAs(UnmanagedType.ByValArray, sizeconst = clsNetworkStats.MAXLEN_PHYSADDR)]
		public byte[] bPhysAddr;
		/// <summary>
		/// dwAdminStatus - The interface is administratively enabled or disabled.
		/// </summary>
		/// <remarks></remarks>
		public UInt32 dwAdminStatus;
		/// <summary>
		/// dwOperStatus - The operational status of the interface. This member can be one of the following values defined in the INTERNAL_IF_OPER_STATUS enumeration defined in the Ipifcons.h header file.
		/// </summary>
		/// <remarks></remarks>
		public UInt32 dwOperStatus;
		/// <summary>
		/// dwLastChange - The length of time, in hundredths of seconds (10^-2 sec), starting from the last computer restart, when the interface entered its current operational state. This value rolls over after 2^32 hundredths of a second.
		/// </summary>
		/// <remarks></remarks>
		public UInt32 dwLastChange;
		/// <summary>
		/// dwInOctets - The number of octets of data received through this interface.
		/// </summary>
		/// <remarks></remarks>
		public UInt32 dwInOctets;
		/// <summary>
		/// dwInUcastPkts - The number of unicast packets received through this interface.
		/// </summary>
		/// <remarks></remarks>
		public UInt32 dwInUcastPkts;
		/// <summary>
		/// dwInNUcastPkts - The number of non-unicast packets received through this interface. Broadcast and multicast packets are included.
		/// </summary>
		/// <remarks></remarks>
		public UInt32 dwInNUcastPkts;
		/// <summary>
		/// dwInDiscards - The number of incoming packets that were discarded even though they did not have errors.
		/// </summary>
		/// <remarks></remarks>
		public UInt32 dwInDiscards;
		/// <summary>
		/// dwInErrors - The number of incoming packets that were discarded because of errors.
		/// </summary>
		/// <remarks></remarks>
		public UInt32 dwInErrors;
		/// <summary>
		/// dwInUnknownProtos - The number of incoming packets that were discarded because the protocol was unknown.
		/// </summary>
		/// <remarks></remarks>
		public UInt32 dwInUnknownProtos;
		/// <summary>
		/// dwOutOctets - The number of octets of data sent through this interface.
		/// </summary>
		/// <remarks></remarks>
		public UInt32 dwOutOctets;
		/// <summary>
		/// dwOutUcastPkts - The number of unicast packets sent through this interface.
		/// </summary>
		/// <remarks></remarks>
		public UInt32 dwOutUcastPkts;
		/// <summary>
		/// dwOutNUcastPkts - The number of non-unicast packets sent through this interface. Broadcast and multicast packets are included.
		/// </summary>
		/// <remarks></remarks>
		public UInt32 dwOutNUcastPkts;
		/// <summary>
		/// dwOutDiscards - The number of outgoing packets that were discarded even though they did not have errors.
		/// </summary>
		/// <remarks></remarks>
		public UInt32 dwOutDiscards;
		/// <summary>
		/// dwOutErrors - The number of outgoing packets that were discarded because of errors.
		/// </summary>
		/// <remarks></remarks>
		public UInt32 dwOutErrors;
		/// <summary>
		/// dwOutQLen - The transmit queue length. This field is not currently used.
		/// </summary>
		/// <remarks></remarks>
		public UInt32 dwOutQLen;
		/// <summary>
		/// dwDescrLen - The length, in bytes, of the bDescr member.
		/// </summary>
		/// <remarks></remarks>
		public UInt32 dwDescrLen;
		/// <summary>
		/// bDescr - A description of the interface.
		/// </summary>
		/// <remarks></remarks>
		[MarshalAs(UnmanagedType.ByValArray, sizeconst = clsNetworkStats.MAXLEN_IFDESCR)]
		public byte[] bDescr;
	}

	/// <summary>
	/// IFROW_HELPER - Undocumented!
	/// </summary>
	/// <remarks></remarks>
	public struct IFROW_HELPER
	{
		/// <summary>
		/// Name - Undocumented!
		/// </summary>
		/// <remarks></remarks>
		public string Name;
		/// <summary>
		/// Index - Undocumented!
		/// </summary>
		/// <remarks></remarks>
		public int Index;
		/// <summary>
		/// Type - Undocumented!
		/// </summary>
		/// <remarks></remarks>
		public int Type;
		/// <summary>
		/// Mtu - Undocumented!
		/// </summary>
		/// <remarks></remarks>
		public int Mtu;
		/// <summary>
		/// Speed - Undocumented!
		/// </summary>
		/// <remarks>Changed from Integer to Long for VISTA</remarks>
		public long Speed;
		/// <summary>
		/// PhysAddrLen - Undocumented!
		/// </summary>
		/// <remarks></remarks>
		public int PhysAddrLen;
		/// <summary>
		/// PhysAddr - Undocumented!
		/// </summary>
		/// <remarks></remarks>
		public string PhysAddr;
		/// <summary>
		/// AdminStatus - Undocumented!
		/// </summary>
		/// <remarks></remarks>
		public int AdminStatus;
		/// <summary>
		/// OperStatus - Undocumented!
		/// </summary>
		/// <remarks></remarks>
		public int OperStatus;
		/// <summary>
		/// LastChange - Undocumented!
		/// </summary>
		/// <remarks>Changed from Integer to Long to make it work</remarks>
		public long LastChange;
		/// <summary>
		/// InOctets - Undocumented!
		/// </summary>
		/// <remarks></remarks>
		public int InOctets;
		/// <summary>
		/// InUcastPkts - Undocumented!
		/// </summary>
		/// <remarks></remarks>
		public int InUcastPkts;
		/// <summary>
		/// InNUcastPkts - Undocumented!
		/// </summary>
		/// <remarks></remarks>
		public int InNUcastPkts;
		/// <summary>
		/// InDiscards - Undocumented!
		/// </summary>
		/// <remarks></remarks>
		public int InDiscards;
		/// <summary>
		/// InErrors - Undocumented!
		/// </summary>
		/// <remarks></remarks>
		public int InErrors;
		/// <summary>
		/// InUnknownProtos - Undocumented!
		/// </summary>
		/// <remarks></remarks>
		public int InUnknownProtos;
		/// <summary>
		/// OutOctets - Undocumented!
		/// </summary>
		/// <remarks></remarks>
		public int OutOctets;
		/// <summary>
		/// OutUcastPkts - Undocumented!
		/// </summary>
		/// <remarks></remarks>
		public int OutUcastPkts;
		/// <summary>
		/// OutNUcastPkts - Undocumented!
		/// </summary>
		/// <remarks></remarks>
		public int OutNUcastPkts;
		/// <summary>
		/// OutDiscards - Undocumented!
		/// </summary>
		/// <remarks></remarks>
		public int OutDiscards;
		/// <summary>
		/// OutErrors - Undocumented!
		/// </summary>
		/// <remarks></remarks>
		public int OutErrors;
		/// <summary>
		/// OutQLen - Undocumented!
		/// </summary>
		/// <remarks></remarks>
		public int OutQLen;
		/// <summary>
		/// Description - Undocumented!
		/// </summary>
		/// <remarks></remarks>
		public string Description;
		/// <summary>
		/// InMegs - Undocumented!
		/// </summary>
		/// <remarks></remarks>
		public string InMegs;
		/// <summary>
		/// OutMegs - Undocumented!
		/// </summary>
		/// <remarks></remarks>
		public string OutMegs;
	}

	//typedef struct _MIB_IFTABLE {
	//  DWORD dwNumEntries;
	//  MIB_IFROW table[ANY_SIZE];
	//} MIB_IFTABLE, *PMIB_IFTABLE;
	//<StructLayout(LayoutKind.Sequential)> Public Class MIB_IFTABLE
	//  Public dwNumEntries As Integer
	//  <MarshalAs(UnmanagedType.SafeArray)> Public table() As MIB_IFROW
	//End Class
	/// <summary>
	/// The MIB_IFTABLE structure contains a table of interface entries.
	/// </summary>
	/// <remarks></remarks>
	[StructLayout(LayoutKind.Sequential)]
	public struct MIB_IFTABLE
	{
		public UInt32 dwNumEntries;
		//<MarshalAs(UnmanagedType.SafeArray)> Dim table() As MIB_IFROW
		public IntPtr table;
	}

	//DWORD GetIfTable(
	//  PMIB_IFTABLE pIfTable,
	//  PULONG pdwSize,
	//  BOOL bOrder
	//);
	//<DllImport("iphlpapi")> Private Shared Function GetIfTable(ByRef pIfRowTable As Byte, ByRef pdwSize As Int32, ByVal bOrder As Int32) As Int32
	//End Function
	/// <summary>
	/// The GetIfTable function retrieves the MIB-II interface table.
	/// </summary>
	/// <param name="pIfRowTable">A pointer to a buffer that receives the interface table as a MIB_IFTABLE structure.</param>
	/// <param name="pdwSize">On input, specifies the size in bytes of the buffer pointed to by the pIfTable parameter.</param>
	/// <param name="bOrder">A Boolean value that specifies whether the returned interface table should be sorted in ascending order by interface index. If this parameter is TRUE, the table is sorted.</param>
	/// <returns>If the function succeeds, the return value is NO_ERROR.</returns>
	/// <remarks>See http://msdn.microsoft.com/en-us/library/aa365943(VS.85).aspx for more info!</remarks>
	[DllImport("iphlpapi")]
	private static extern int GetIfTable(IntPtr pIfRowTable, ref int pdwSize, bool bOrder);

	/// <summary>
	/// The GetIfEntry function retrieves information for the specified interface on the local computer.
	/// </summary>
	/// <param name="pIfRow">A pointer to a <a href="http://msdn.microsoft.com/en-us/library/aa366836(VS.85).aspx">MIB_IFROW</a> structure that, on successful return, receives information for an interface on the local computer. On input, set the dwIndex member of MIB_IFROW to the index of the interface for which to retrieve information. The value for the dwIndex must be retrieved by a previous call to the GetIfTable, GetIfTable2, or GetIfTable2Ex function.</param>
	/// <returns></returns>
	/// <remarks>See http://msdn.microsoft.com/en-us/library/aa365939(VS.85).aspx for more info!</remarks>
	[DllImport("iphlpapi")]
	private static extern Int32 GetIfEntry(ref MIB_IFROW pIfRow);

	#endregion


	private ArrayList m_Adapters;
	/// <summary>
	/// New - Undocumented!
	/// </summary>
	/// <param name="IgnoreLoopBack"></param>
	/// <remarks></remarks>
	public clsNetworkStats([System.Runtime.InteropServices.OptionalAttribute, System.Runtime.InteropServices.DefaultParameterValueAttribute(true)]  // ERROR: Optional parameters aren't supported in C#
bool IgnoreLoopBack)
	{
		long ret = 0;
		MIB_IFROW ifrow = new MIB_IFROW();
		MIB_IFTABLE IfTable = new MIB_IFTABLE();
		int tablesize = 0;
		IntPtr iBuf = IntPtr.Zero;

		//get tablesize
		ret = GetIfTable(iBuf, ref tablesize, true);

		//resize buffer on tablesize
		iBuf = Marshal.AllocHGlobal(tablesize);
		//load buffer
		ret = GetIfTable(iBuf, ref tablesize, true);

		//marshar from buffer into MIB_IFTABLE structure
		IfTable = (MIB_IFTABLE)Marshal.PtrToStructure(iBuf, typeof(MIB_IFTABLE));

		//initialize adapter list
		m_Adapters = new ArrayList();

		int noInterfaces = Convert.ToInt32(IfTable.dwNumEntries);
		//NIC Interfaces number
		int IFROWSize = Marshal.SizeOf(ifrow);

		MIB_IFROW[] mrows = new MIB_IFROW[(int)noInterfaces + 1];

		byte[] mDest = null;
		mDest = new byte[IFROWSize * noInterfaces + 1];

		for (byte i = 1; i <= Convert.ToByte(noInterfaces); i++) {
			mrows[i - 1] = (MIB_IFROW)Marshal.PtrToStructure(new IntPtr(iBuf.ToInt32() + 4 + (i - 1) * IFROWSize), typeof(MIB_IFROW));
			IFROW_HELPER ifhelp = PrivToPub(mrows[i - 1]);
			if (IgnoreLoopBack == true) {
				if (ifhelp.Description.IndexOf("Loopback") < 0) {
					m_Adapters.Add(ifhelp);
				}
			}
			else {
				m_Adapters.Add(ifhelp);
			}
		}
		Marshal.FreeHGlobal(iBuf);

	}

	/// <summary>
	/// GetAdapter - Undocumented!
	/// </summary>
	/// <returns></returns>
	/// <remarks></remarks>
	public IFROW_HELPER GetAdapter()
	{
		IFROW_HELPER functionReturnValue = default(IFROW_HELPER);
		short i = 0;
		for (i = 0; i <= Convert.ToInt16(m_Adapters.Count); i++) {
			functionReturnValue = (IFROW_HELPER)m_Adapters(i);
			if (functionReturnValue.PhysAddr.ToString != "00-00-00-00-00-00") return;
 
		}
		return null;
		return functionReturnValue;
	}

	/// <summary>
	/// PrivToPub - Undocumented!
	/// </summary>
	/// <param name="pri"></param>
	/// <returns></returns>
	/// <remarks></remarks>
	[DebuggerStepThrough()]

	private IFROW_HELPER PrivToPub(MIB_IFROW pri)
	{
		IFROW_HELPER ifhelp = new IFROW_HELPER();

		ifhelp.Name = pri.wszName.Trim();
		ifhelp.Index = Convert.ToInt32(pri.dwIndex);
		ifhelp.Type = Convert.ToInt32(pri.dwType);
		ifhelp.Mtu = Convert.ToInt32(pri.dwMtu);
		ifhelp.Speed = Convert.ToInt64(pri.dwSpeed);
		// Changed from ToInt32 to ToInt64 for VISTA
		ifhelp.PhysAddrLen = Convert.ToInt32(pri.dwPhysAddrLen);
		ifhelp.PhysAddr = MAC2String(pri.bPhysAddr);
		ifhelp.AdminStatus = Convert.ToInt32(pri.dwAdminStatus);
		ifhelp.OperStatus = Convert.ToInt32(pri.dwOperStatus);
		ifhelp.LastChange = Convert.ToInt64(pri.dwLastChange);
		// Changed from ToInt32 to ToInt64 to make it work and to uncomment
		ifhelp.InOctets = Convert.ToInt32(pri.dwInOctets);
		ifhelp.InUcastPkts = Convert.ToInt32(pri.dwInUcastPkts);
		ifhelp.InNUcastPkts = Convert.ToInt32(pri.dwInNUcastPkts);
		ifhelp.InDiscards = Convert.ToInt32(pri.dwInDiscards);
		ifhelp.InErrors = Convert.ToInt32(pri.dwInErrors);
		ifhelp.InUnknownProtos = Convert.ToInt32(pri.dwInUnknownProtos);
		ifhelp.OutOctets = Convert.ToInt32(pri.dwOutOctets);
		ifhelp.OutUcastPkts = Convert.ToInt32(pri.dwOutUcastPkts);
		ifhelp.OutNUcastPkts = Convert.ToInt32(pri.dwOutNUcastPkts);
		ifhelp.OutDiscards = Convert.ToInt32(pri.dwOutDiscards);
		ifhelp.OutErrors = Convert.ToInt32(pri.dwOutErrors);
		ifhelp.OutQLen = Convert.ToInt32(pri.dwOutQLen);
		ifhelp.Description = System.Text.Encoding.ASCII.GetString(pri.bDescr, 0, Convert.ToInt32(pri.dwDescrLen));
		ifhelp.InMegs = ToMegs(ifhelp.InOctets);
		ifhelp.OutMegs = ToMegs(ifhelp.OutOctets);

		return ifhelp;

	}

	/// <summary>
	/// ToMegs - Undocumented!
	/// </summary>
	/// <param name="lSize"></param>
	/// <returns></returns>
	/// <remarks></remarks>
	[DebuggerStepThrough()]

	private string ToMegs(long lSize)
	{
		string sDenominator = " B";
		//If lSize > 1024 Then lSize = (lSize / 1024) * 1000 'Windows styleee filesizing : ) 
		if (lSize > 1000) {
			sDenominator = " KB";
			lSize = Convert.ToInt64(lSize / 1000);
		}
		else if (lSize <= 1000) {
			sDenominator = " B";
			lSize = lSize;
		}

		return Strings.Format(lSize, "###,###0") + sDenominator;

	}

	/// <summary>
	/// Convert a byte array containing a MAC address to a hex string
	/// </summary>
	/// <param name="AdrArray"></param>
	/// <returns></returns>
	/// <remarks></remarks>
	private string MAC2String(byte[] AdrArray)
	{
		string aStr = null;
		string hexStr = null;
		int i = 0;

		for (i = 0; i <= 5; i++) {
			if ((i > Information.UBound(AdrArray))) {
				hexStr = "00";
			}
			else {
				hexStr = Conversion.Hex(AdrArray[i]);
			}

			if ((hexStr.Length < 2)) hexStr = "0" + hexStr; 
			aStr = aStr + hexStr;
			if ((i < 5)) aStr = aStr + "-"; 
		}

		return aStr;

	}

}

Imports System.Runtime.InteropServices

''' <summary>
''' clsNetworkStats - Needs updates to comments!
''' </summary>
''' <remarks></remarks>
Class clsNetworkStats

#Region " DECLARES "
    Private Const ERROR_SUCCESS As Long = 0
    Private Const MAX_INTERFACE_NAME_LEN As Long = 256
    Private Const MAXLEN_IFDESCR As Long = 256
    Private Const MAXLEN_PHYSADDR As Long = 8

    ''' <summary>
    ''' MIB_IFROW - The MIB_IFROW structure stores information about a particular interface.
    ''' </summary>
    ''' <remarks>See http://msdn.microsoft.com/en-us/library/aa366836(VS.85).aspx for Full Documentation!</remarks>
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> Public Structure MIB_IFROW
        ''' <summary>
        ''' wszName - A pointer to a Unicode string that contains the name of the interface.
        ''' </summary>
        ''' <remarks></remarks>
        <MarshalAs(UnmanagedType.ByValTStr, sizeconst:=MAX_INTERFACE_NAME_LEN)> Public wszName As String
        ''' <summary>
        ''' dwIndex - The index that identifies the interface. This index value may change when a network adapter is disabled and then enabled, and should not be considered persistent.
        ''' </summary>
        ''' <remarks></remarks>
        Public dwIndex As UInt32
        ''' <summary>
        ''' dwType = The interface type as defined by the Internet Assigned Names Authority (IANA). For more information, see <a href="http://www.iana.org/assignments/ianaiftype-mib">http://www.iana.org/assignments/ianaiftype-mib</a>. Possible values for the interface type are listed in the Ipifcons.h header file. 
        ''' </summary>
        ''' <remarks>See http://msdn.microsoft.com/en-us/library/aa366836(VS.85).aspx for more info!</remarks>
        Public dwType As UInt32
        ''' <summary>
        ''' dwMtu - The Maximum Transmission Unit (MTU) size in bytes.
        ''' </summary>
        ''' <remarks></remarks>
        Public dwMtu As UInt32
        ''' <summary>
        ''' dwSpeed - The speed of the interface in bits per second.
        ''' </summary>
        ''' <remarks></remarks>
        Public dwSpeed As UInt32
        ''' <summary>
        ''' dwPhysAddrLen - The length, in bytes, of the physical address specified by the bPhysAddr member.
        ''' </summary>
        ''' <remarks></remarks>
        Public dwPhysAddrLen As UInt32
        ''' <summary>
        ''' bPhysAddr - The physical address of the adapter for this interface.
        ''' </summary>
        ''' <remarks></remarks>
        <MarshalAs(UnmanagedType.ByValArray, sizeconst:=MAXLEN_PHYSADDR)> Public bPhysAddr() As Byte
        ''' <summary>
        ''' dwAdminStatus - The interface is administratively enabled or disabled.
        ''' </summary>
        ''' <remarks></remarks>
        Public dwAdminStatus As UInt32
        ''' <summary>
        ''' dwOperStatus - The operational status of the interface. This member can be one of the following values defined in the INTERNAL_IF_OPER_STATUS enumeration defined in the Ipifcons.h header file.
        ''' </summary>
        ''' <remarks></remarks>
        Public dwOperStatus As UInt32
        ''' <summary>
        ''' dwLastChange - The length of time, in hundredths of seconds (10^-2 sec), starting from the last computer restart, when the interface entered its current operational state. This value rolls over after 2^32 hundredths of a second.
        ''' </summary>
        ''' <remarks></remarks>
        Public dwLastChange As UInt32
        ''' <summary>
        ''' dwInOctets - The number of octets of data received through this interface.
        ''' </summary>
        ''' <remarks></remarks>
        Public dwInOctets As UInt32
        ''' <summary>
        ''' dwInUcastPkts - The number of unicast packets received through this interface.
        ''' </summary>
        ''' <remarks></remarks>
        Public dwInUcastPkts As UInt32
        ''' <summary>
        ''' dwInNUcastPkts - The number of non-unicast packets received through this interface. Broadcast and multicast packets are included.
        ''' </summary>
        ''' <remarks></remarks>
        Public dwInNUcastPkts As UInt32
        ''' <summary>
        ''' dwInDiscards - The number of incoming packets that were discarded even though they did not have errors.
        ''' </summary>
        ''' <remarks></remarks>
        Public dwInDiscards As UInt32
        ''' <summary>
        ''' dwInErrors - The number of incoming packets that were discarded because of errors.
        ''' </summary>
        ''' <remarks></remarks>
        Public dwInErrors As UInt32
        ''' <summary>
        ''' dwInUnknownProtos - The number of incoming packets that were discarded because the protocol was unknown.
        ''' </summary>
        ''' <remarks></remarks>
        Public dwInUnknownProtos As UInt32
        ''' <summary>
        ''' dwOutOctets - The number of octets of data sent through this interface.
        ''' </summary>
        ''' <remarks></remarks>
        Public dwOutOctets As UInt32
        ''' <summary>
        ''' dwOutUcastPkts - The number of unicast packets sent through this interface.
        ''' </summary>
        ''' <remarks></remarks>
        Public dwOutUcastPkts As UInt32
        ''' <summary>
        ''' dwOutNUcastPkts - The number of non-unicast packets sent through this interface. Broadcast and multicast packets are included.
        ''' </summary>
        ''' <remarks></remarks>
        Public dwOutNUcastPkts As UInt32
        ''' <summary>
        ''' dwOutDiscards - The number of outgoing packets that were discarded even though they did not have errors.
        ''' </summary>
        ''' <remarks></remarks>
        Public dwOutDiscards As UInt32
        ''' <summary>
        ''' dwOutErrors - The number of outgoing packets that were discarded because of errors.
        ''' </summary>
        ''' <remarks></remarks>
        Public dwOutErrors As UInt32
        ''' <summary>
        ''' dwOutQLen - The transmit queue length. This field is not currently used.
        ''' </summary>
        ''' <remarks></remarks>
        Public dwOutQLen As UInt32
        ''' <summary>
        ''' dwDescrLen - The length, in bytes, of the bDescr member.
        ''' </summary>
        ''' <remarks></remarks>
        Public dwDescrLen As UInt32
        ''' <summary>
        ''' bDescr - A description of the interface.
        ''' </summary>
        ''' <remarks></remarks>
        <MarshalAs(UnmanagedType.ByValArray, sizeconst:=MAXLEN_IFDESCR)> Public bDescr() As Byte
    End Structure

    ''' <summary>
    ''' IFROW_HELPER - Undocumented!
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure IFROW_HELPER
        ''' <summary>
        ''' Name - Undocumented!
        ''' </summary>
        ''' <remarks></remarks>
        Public Name As String
        ''' <summary>
        ''' Index - Undocumented!
        ''' </summary>
        ''' <remarks></remarks>
        Public Index As Integer
        ''' <summary>
        ''' Type - Undocumented!
        ''' </summary>
        ''' <remarks></remarks>
        Public Type As Integer
        ''' <summary>
        ''' Mtu - Undocumented!
        ''' </summary>
        ''' <remarks></remarks>
        Public Mtu As Integer
        ''' <summary>
        ''' Speed - Undocumented!
        ''' </summary>
        ''' <remarks>Changed from Integer to Long for VISTA</remarks>
        Public Speed As Long
        ''' <summary>
        ''' PhysAddrLen - Undocumented!
        ''' </summary>
        ''' <remarks></remarks>
        Public PhysAddrLen As Integer
        ''' <summary>
        ''' PhysAddr - Undocumented!
        ''' </summary>
        ''' <remarks></remarks>
        Public PhysAddr As String
        ''' <summary>
        ''' AdminStatus - Undocumented!
        ''' </summary>
        ''' <remarks></remarks>
        Public AdminStatus As Integer
        ''' <summary>
        ''' OperStatus - Undocumented!
        ''' </summary>
        ''' <remarks></remarks>
        Public OperStatus As Integer
        ''' <summary>
        ''' LastChange - Undocumented!
        ''' </summary>
        ''' <remarks>Changed from Integer to Long to make it work</remarks>
        Public LastChange As Long
        ''' <summary>
        ''' InOctets - Undocumented!
        ''' </summary>
        ''' <remarks></remarks>
        Public InOctets As Integer
        ''' <summary>
        ''' InUcastPkts - Undocumented!
        ''' </summary>
        ''' <remarks></remarks>
        Public InUcastPkts As Integer
        ''' <summary>
        ''' InNUcastPkts - Undocumented!
        ''' </summary>
        ''' <remarks></remarks>
        Public InNUcastPkts As Integer
        ''' <summary>
        ''' InDiscards - Undocumented!
        ''' </summary>
        ''' <remarks></remarks>
        Public InDiscards As Integer
        ''' <summary>
        ''' InErrors - Undocumented!
        ''' </summary>
        ''' <remarks></remarks>
        Public InErrors As Integer
        ''' <summary>
        ''' InUnknownProtos - Undocumented!
        ''' </summary>
        ''' <remarks></remarks>
        Public InUnknownProtos As Integer
        ''' <summary>
        ''' OutOctets - Undocumented!
        ''' </summary>
        ''' <remarks></remarks>
        Public OutOctets As Integer
        ''' <summary>
        ''' OutUcastPkts - Undocumented!
        ''' </summary>
        ''' <remarks></remarks>
        Public OutUcastPkts As Integer
        ''' <summary>
        ''' OutNUcastPkts - Undocumented!
        ''' </summary>
        ''' <remarks></remarks>
        Public OutNUcastPkts As Integer
        ''' <summary>
        ''' OutDiscards - Undocumented!
        ''' </summary>
        ''' <remarks></remarks>
        Public OutDiscards As Integer
        ''' <summary>
        ''' OutErrors - Undocumented!
        ''' </summary>
        ''' <remarks></remarks>
        Public OutErrors As Integer
        ''' <summary>
        ''' OutQLen - Undocumented!
        ''' </summary>
        ''' <remarks></remarks>
        Public OutQLen As Integer
        ''' <summary>
        ''' Description - Undocumented!
        ''' </summary>
        ''' <remarks></remarks>
        Public Description As String
        ''' <summary>
        ''' InMegs - Undocumented!
        ''' </summary>
        ''' <remarks></remarks>
        Public InMegs As String
        ''' <summary>
        ''' OutMegs - Undocumented!
        ''' </summary>
        ''' <remarks></remarks>
        Public OutMegs As String
    End Structure

    'typedef struct _MIB_IFTABLE {
    '  DWORD dwNumEntries;
    '  MIB_IFROW table[ANY_SIZE];
    '} MIB_IFTABLE, *PMIB_IFTABLE;
    '<StructLayout(LayoutKind.Sequential)> Public Class MIB_IFTABLE
    '  Public dwNumEntries As Integer
    '  <MarshalAs(UnmanagedType.SafeArray)> Public table() As MIB_IFROW
    'End Class
    ''' <summary>
    ''' The MIB_IFTABLE structure contains a table of interface entries.
    ''' </summary>
    ''' <remarks></remarks>
    <StructLayout(LayoutKind.Sequential)> _
    Structure MIB_IFTABLE
        Dim dwNumEntries As UInt32
        '<MarshalAs(UnmanagedType.SafeArray)> Dim table() As MIB_IFROW
        Dim table As IntPtr
    End Structure

    'DWORD GetIfTable(
    '  PMIB_IFTABLE pIfTable,
    '  PULONG pdwSize,
    '  BOOL bOrder
    ');
    '<DllImport("iphlpapi")> Private Shared Function GetIfTable(ByRef pIfRowTable As Byte, ByRef pdwSize As Int32, ByVal bOrder As Int32) As Int32
    'End Function
    ''' <summary>
    ''' The GetIfTable function retrieves the MIB-II interface table.
    ''' </summary>
    ''' <param name="pIfRowTable">A pointer to a buffer that receives the interface table as a MIB_IFTABLE structure.</param>
    ''' <param name="pdwSize">On input, specifies the size in bytes of the buffer pointed to by the pIfTable parameter.</param>
    ''' <param name="bOrder">A Boolean value that specifies whether the returned interface table should be sorted in ascending order by interface index. If this parameter is TRUE, the table is sorted.</param>
    ''' <returns>If the function succeeds, the return value is NO_ERROR.</returns>
    ''' <remarks>See http://msdn.microsoft.com/en-us/library/aa365943(VS.85).aspx for more info!</remarks>
    <DllImport("iphlpapi")> Private Shared Function GetIfTable( _
        ByVal pIfRowTable As IntPtr, _
        ByRef pdwSize As Integer, _
        ByVal bOrder As Boolean _
    ) As Integer
    End Function

    ''' <summary>
    ''' The GetIfEntry function retrieves information for the specified interface on the local computer.
    ''' </summary>
    ''' <param name="pIfRow">A pointer to a <a href="http://msdn.microsoft.com/en-us/library/aa366836(VS.85).aspx">MIB_IFROW</a> structure that, on successful return, receives information for an interface on the local computer. On input, set the dwIndex member of MIB_IFROW to the index of the interface for which to retrieve information. The value for the dwIndex must be retrieved by a previous call to the GetIfTable, GetIfTable2, or GetIfTable2Ex function.</param>
    ''' <returns></returns>
    ''' <remarks>See http://msdn.microsoft.com/en-us/library/aa365939(VS.85).aspx for more info!</remarks>
    <DllImport("iphlpapi")> Private Shared Function GetIfEntry(ByRef pIfRow As MIB_IFROW) As Int32
    End Function

#End Region

    Private m_Adapters As ArrayList

    ''' <summary>
    ''' New - Undocumented!
    ''' </summary>
    ''' <param name="IgnoreLoopBack"></param>
    ''' <remarks></remarks>
    Public Sub New(Optional ByVal IgnoreLoopBack As Boolean = True)
        Dim ret As Long
        Dim ifrow As New MIB_IFROW
        Dim IfTable As New MIB_IFTABLE
        Dim tablesize As Integer
        Dim iBuf As IntPtr = IntPtr.Zero

        'get tablesize
        ret = GetIfTable(iBuf, tablesize, True)

        'resize buffer on tablesize
        iBuf = Marshal.AllocHGlobal(tablesize)
        'load buffer
        ret = GetIfTable(iBuf, tablesize, True)

        'marshar from buffer into MIB_IFTABLE structure
        IfTable = CType(Marshal.PtrToStructure(iBuf, GetType(MIB_IFTABLE)), MIB_IFTABLE)

        'initialize adapter list
        m_Adapters = New ArrayList

        Dim noInterfaces As Integer = Convert.ToInt32(IfTable.dwNumEntries) 'NIC Interfaces number
        Dim IFROWSize As Integer = Marshal.SizeOf(ifrow)

        Dim mrows As MIB_IFROW() = New MIB_IFROW(CType(noInterfaces, Integer)) {}

        Dim mDest As Byte()
        ReDim mDest(IFROWSize * noInterfaces)

        For i As Byte = 1 To Convert.ToByte(noInterfaces)
            mrows(i - 1) = CType(Marshal.PtrToStructure(New IntPtr(iBuf.ToInt32 + 4 + (i - 1) * IFROWSize), GetType(MIB_IFROW)), MIB_IFROW)
            Dim ifhelp As IFROW_HELPER = PrivToPub(mrows(i - 1))
            If IgnoreLoopBack = True Then
                If ifhelp.Description.IndexOf("Loopback") < 0 Then
                    m_Adapters.Add(ifhelp)
                End If
            Else
                m_Adapters.Add(ifhelp)
            End If
        Next
        Marshal.FreeHGlobal(iBuf)

    End Sub

    ''' <summary>
    ''' GetAdapter - Undocumented!
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetAdapter() As IFROW_HELPER
        Dim i As Short
        For i = 0 To Convert.ToInt16(m_Adapters.Count)
            GetAdapter = CType(m_Adapters(i), IFROW_HELPER)
            If GetAdapter.PhysAddr.ToString <> "00-00-00-00-00-00" Then Exit Function
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' PrivToPub - Undocumented!
    ''' </summary>
    ''' <param name="pri"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> Private Function PrivToPub(ByVal pri As MIB_IFROW) As IFROW_HELPER

        Dim ifhelp As New IFROW_HELPER

        ifhelp.Name = pri.wszName.Trim
        ifhelp.Index = Convert.ToInt32(pri.dwIndex)
        ifhelp.Type = Convert.ToInt32(pri.dwType)
        ifhelp.Mtu = Convert.ToInt32(pri.dwMtu)
        ifhelp.Speed = Convert.ToInt64(pri.dwSpeed) ' Changed from ToInt32 to ToInt64 for VISTA
        ifhelp.PhysAddrLen = Convert.ToInt32(pri.dwPhysAddrLen)
        ifhelp.PhysAddr = MAC2String(pri.bPhysAddr)
        ifhelp.AdminStatus = Convert.ToInt32(pri.dwAdminStatus)
        ifhelp.OperStatus = Convert.ToInt32(pri.dwOperStatus)
        ifhelp.LastChange = Convert.ToInt64(pri.dwLastChange) ' Changed from ToInt32 to ToInt64 to make it work and to uncomment
        ifhelp.InOctets = Convert.ToInt32(pri.dwInOctets)
        ifhelp.InUcastPkts = Convert.ToInt32(pri.dwInUcastPkts)
        ifhelp.InNUcastPkts = Convert.ToInt32(pri.dwInNUcastPkts)
        ifhelp.InDiscards = Convert.ToInt32(pri.dwInDiscards)
        ifhelp.InErrors = Convert.ToInt32(pri.dwInErrors)
        ifhelp.InUnknownProtos = Convert.ToInt32(pri.dwInUnknownProtos)
        ifhelp.OutOctets = Convert.ToInt32(pri.dwOutOctets)
        ifhelp.OutUcastPkts = Convert.ToInt32(pri.dwOutUcastPkts)
        ifhelp.OutNUcastPkts = Convert.ToInt32(pri.dwOutNUcastPkts)
        ifhelp.OutDiscards = Convert.ToInt32(pri.dwOutDiscards)
        ifhelp.OutErrors = Convert.ToInt32(pri.dwOutErrors)
        ifhelp.OutQLen = Convert.ToInt32(pri.dwOutQLen)
        ifhelp.Description = System.Text.Encoding.ASCII.GetString(pri.bDescr, 0, Convert.ToInt32(pri.dwDescrLen))
        ifhelp.InMegs = ToMegs(ifhelp.InOctets)
        ifhelp.OutMegs = ToMegs(ifhelp.OutOctets)

        Return ifhelp

    End Function

    ''' <summary>
    ''' ToMegs - Undocumented!
    ''' </summary>
    ''' <param name="lSize"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> Private Function ToMegs(ByVal lSize As Long) As String

        Dim sDenominator As String = " B"
        'If lSize > 1024 Then lSize = (lSize / 1024) * 1000 'Windows styleee filesizing : ) 
        If lSize > 1000 Then
            sDenominator = " KB"
            lSize = Convert.ToInt64(lSize / 1000)
        ElseIf lSize <= 1000 Then
            sDenominator = " B"
            lSize = lSize
        End If

        ToMegs = Format(lSize, "###,###0") & sDenominator

    End Function

    ''' <summary>
    ''' Convert a byte array containing a MAC address to a hex string
    ''' </summary>
    ''' <param name="AdrArray"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function MAC2String(ByVal AdrArray() As Byte) As String
        Dim aStr As String = Nothing
        Dim hexStr As String
        Dim i As Integer

        For i = 0 To 5
            If (i > UBound(AdrArray)) Then
                hexStr = "00"
            Else
                hexStr = Hex$(AdrArray(i))
            End If

            If (hexStr.Length < 2) Then hexStr = "0" & hexStr
            aStr = aStr & hexStr
            If (i < 5) Then aStr = aStr & "-"
        Next i

        MAC2String = aStr

    End Function

End Class

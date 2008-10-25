Imports System.Runtime.InteropServices

Class clsNetworkStats

#Region " DECLARES "
    Private Const ERROR_SUCCESS As Long = 0
    Private Const MAX_INTERFACE_NAME_LEN As Long = 256
    Private Const MAXLEN_IFDESCR As Long = 256
    Private Const MAXLEN_PHYSADDR As Long = 8

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> Public Structure MIB_IFROW
        <MarshalAs(UnmanagedType.ByValTStr, sizeconst:=MAX_INTERFACE_NAME_LEN)> Public wszName As String
        Public dwIndex As UInt32
        Public dwType As UInt32
        Public dwMtu As UInt32
        Public dwSpeed As UInt32
        Public dwPhysAddrLen As UInt32
        <MarshalAs(UnmanagedType.ByValArray, sizeconst:=MAXLEN_PHYSADDR)> Public bPhysAddr() As Byte
        Public dwAdminStatus As UInt32
        Public dwOperStatus As UInt32
        Public dwLastChange As UInt32
        Public dwInOctets As UInt32
        Public dwInUcastPkts As UInt32
        Public dwInNUcastPkts As UInt32
        Public dwInDiscards As UInt32
        Public dwInErrors As UInt32
        Public dwInUnknownProtos As UInt32
        Public dwOutOctets As UInt32
        Public dwOutUcastPkts As UInt32
        Public dwOutNUcastPkts As UInt32
        Public dwOutDiscards As UInt32
        Public dwOutErrors As UInt32
        Public dwOutQLen As UInt32
        Public dwDescrLen As UInt32
        <MarshalAs(UnmanagedType.ByValArray, sizeconst:=MAXLEN_IFDESCR)> Public bDescr() As Byte
    End Structure

    Public Structure IFROW_HELPER
        Public Name As String
        Public Index As Integer
        Public Type As Integer
        Public Mtu As Integer
        Public Speed As Long ' Change from int to long for VISTA
        Public PhysAddrLen As Integer
        Public PhysAddr As String
        Public AdminStatus As Integer
        Public OperStatus As Integer
        Public LastChange As Integer
        Public InOctets As Integer
        Public InUcastPkts As Integer
        Public InNUcastPkts As Integer
        Public InDiscards As Integer
        Public InErrors As Integer
        Public InUnknownProtos As Integer
        Public OutOctets As Integer
        Public OutUcastPkts As Integer
        Public OutNUcastPkts As Integer
        Public OutDiscards As Integer
        Public OutErrors As Integer
        Public OutQLen As Integer
        Public Description As String
        Public InMegs As String
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
    <DllImport("iphlpapi")> Private Shared Function GetIfTable( _
        ByVal pIfRowTable As IntPtr, _
        ByRef pdwSize As Integer, _
        ByVal bOrder As Boolean _
    ) As Integer
    End Function


    <DllImport("iphlpapi")> Private Shared Function GetIfEntry(ByRef pIfRow As MIB_IFROW) As Int32
    End Function

#End Region

    Private m_Adapters As ArrayList

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

    Public Function GetAdapter() As IFROW_HELPER
        Dim i As Short
        For i = 0 To Convert.ToInt16(m_Adapters.Count)
            Return CType(m_Adapters(i), IFROW_HELPER)
            If GetAdapter.PhysAddr.ToString <> "00-00-00-00-00-00" Then Exit Function
        Next
        Return Nothing
    End Function

    <DebuggerStepThrough()> Private Function PrivToPub(ByVal pri As MIB_IFROW) As IFROW_HELPER

        Dim ifhelp As New IFROW_HELPER

        ifhelp.Name = pri.wszName.Trim
        ifhelp.Index = Convert.ToInt32(pri.dwIndex)
        ifhelp.Type = Convert.ToInt32(pri.dwType)
        ifhelp.Mtu = Convert.ToInt32(pri.dwMtu)
        ifhelp.Speed = Convert.ToInt64(pri.dwSpeed) ' Change from ToInt32 to ToInt64 for VISTA
        ifhelp.PhysAddrLen = Convert.ToInt32(pri.dwPhysAddrLen)
        ifhelp.PhysAddr = MAC2String(pri.bPhysAddr)
        ifhelp.AdminStatus = Convert.ToInt32(pri.dwAdminStatus)
        ifhelp.OperStatus = Convert.ToInt32(pri.dwOperStatus)
        'ifhelp.LastChange = Convert.ToInt32(pri.dwLastChange)
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

    ' Convert a byte array containing a MAC address to a hex string
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

            If (Len(hexStr) < 2) Then hexStr = "0" & hexStr
            aStr = aStr & hexStr
            If (i < 5) Then aStr = aStr & "-"
        Next i

        MAC2String = aStr

    End Function

End Class

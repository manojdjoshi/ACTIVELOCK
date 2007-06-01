Attribute VB_Name = "modTransfer"
Option Explicit


Public Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
   (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'------------------ begin Send Message API --------------------
'Declare API calls for SendMessage function
'which is used to provide Undo capability and closing apps.
'Specifically used with TextUndo and AppClose subs.
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
'Constants for SendMessage
Public Const EM_SETTABSTOPS = &HCB
Public Const WM_CLOSE = &H10
Public Const EM_UNDO = &H417
'------------------ end Send Message API --------------------

'------------------ begin Browse Folders API --------------------
' Currently selected option button
Dim m_wCurOptIdx As Integer

'///////////////////////////////////////////////////////////////////////////////////////////////////////////

' A little info...
' Objects in the shell’s namespace are assigned item identifiers and item
' identifier lists. An item identifier uniquely identifies an item within its parent
' folder. An item identifier list uniquely identifies an item within the shell’s
' namespace by tracing a path to the item from the desktop.

'///////////////////////////////////////////////////////////////////////////////////////////////////////////

' An item identifier is defined by the variable-length SHITEMID structure.
' The first two bytes of this structure specify its size, and the format of
' the remaining bytes depends on the parent folder, or more precisely
' on the software that implements the parent folder’s IShellFolder interface.
' Except for the first two bytes, item identifiers are not strictly defined, and
' applications should make no assumptions about their format.
Private Type SHITEMID   ' mkid
    cb As Long       ' Size of the ID (including cb itself)
    abID() As Byte  ' The item ID (variable length)
End Type

' The ITEMIDLIST structure defines an element in an item identifier list
' (the only member of this structure is an SHITEMID structure). An item
' identifier list consists of one or more consecutive ITEMIDLIST structures
' packed on byte boundaries, followed by a 16-bit zero value. An application
' can walk a list of item identifiers by examining the size specified in each
' SHITEMID structure and stopping when it finds a size of zero. A pointer
' to an item identifier list, is sometimes called a PIDL (pronounced piddle)
Private Type ITEMIDLIST   ' idl
    mkid As SHITEMID
End Type

' Converts an item identifier list to a file system path.
' Returns TRUE if successful or FALSE if an error occurs, for example,
' if the location specified by the pidl parameter is not part of the file system.
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
    (ByVal pidl As Long, ByVal pszPath As String) As Long

' Retrieves the location of a special (system) folder.
' Returns NOERROR if successful or an OLE-defined error result otherwise.
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
    (ByVal hwndOwner As Long, ByVal nFolder As Long, _
    pidl As Long) As Long

' SHGetSpecialFolderLocation successful rtn val
Private Const NoError = 0

' SHGetSpecialFolderLocation nFolder params:
' Most folder locations are stored in:
' [HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders]
' Value specifying the types of folders to be listed in the dialog box as well as other options.
' This member can be 0 or one of the following values:

' Windows desktop, virtual folder at the root of the name space.
Const CSIDL_DESKTOP = &H0
Const CSIDL_INTERNET = &H1
' File system directory that contains the user's program groups
' (which are also file system directories).
Const CSIDL_PROGRAMS = &H2
' Control Panel, virtual folder containing icons for the control panel applications.
Const CSIDL_CONTROLS = &H3
' Printers folder, virtual folder containing installed printers.
Const CSIDL_PRINTERS = &H4
' File system directory that serves as a common respository for documents.
Const CSIDL_PERSONAL = &H5   ' (Documents folder)
' File system directory that contains the user's favorite Internet Explorer URLs.
Const CSIDL_FAVORITES = &H6
' File system directory that corresponds to the user's Startup program group.
Const CSIDL_STARTUP = &H7
' File system directory that contains the user's most recently used documents.
Const CSIDL_RECENT = &H8   ' (Recent folder)
' File system directory that contains Send To menu items.
Const CSIDL_SENDTO = &H9

' Recycle bin, file system directory containing file objects in the user's recycle bin.
' The location of this directory is not in the registry; it is marked with the hidden and
' system attributes to prevent the user from moving or deleting it.
Const CSIDL_BITBUCKET = &HA
' File system directory containing Start menu items.
Const CSIDL_STARTMENU = &HB
Const CSIDL_MYDOCUMENTS = &HC
Const CSIDL_MYMUSIC = &HD
Const CSIDL_MYVIDEO = &HE

' File system directory used to physically store file objects on the desktop
' (not to be confused with the desktop folder itself).
Const CSIDL_DESKTOPDIRECTORY = &H10
' My Computer, virtual folder containing everything on the local computer: storage
' devices, printers, and Control Panel. The folder may also contain mapped network drives.
Const CSIDL_DRIVES = &H11
' Network Neighborhood, virtual folder representing the top level of the network hierarchy.
Const CSIDL_NETWORK = &H12
' File system directory containing objects that appear in the network neighborhood.
Const CSIDL_NETHOOD = &H13
' Virtual folder containing fonts.
Const CSIDL_FONTS = &H14
' File system directory that serves as a common repository for document templates.
Const CSIDL_TEMPLATES = &H15   ' (ShellNew folder)
Const CSIDL_COMMON_STARTMENU           As Long = &H16
Const CSIDL_COMMON_PROGRAMS            As Long = &H17
Const CSIDL_COMMON_STARTUP             As Long = &H18
Const CSIDL_COMMON_DESKTOPDIRECTORY    As Long = &H19
Const CSIDL_APPDATA                    As Long = &H1A
Const CSIDL_PRINTHOOD                  As Long = &H1B
Const CSIDL_LOCAL_APPDATA              As Long = &H1C
Const CSIDL_ALTSTARTUP                 As Long = &H1D
Const CSIDL_COMMON_ALTSTARTUP          As Long = &H1E
 
Const CSIDL_COMMON_FAVORITES           As Long = &H1F
Const CSIDL_INTERNET_CACHE             As Long = &H20
Const CSIDL_COOKIES                    As Long = &H21
Const CSIDL_HISTORY                    As Long = &H22
Const CSIDL_COMMON_APPDATA             As Long = &H23
Const CSIDL_WINDOWS                    As Long = &H24
Const CSIDL_SYSTEM                     As Long = &H25
Const CSIDL_PROGRAM_FILES              As Long = &H26
Const CSIDL_MYPICTURES                 As Long = &H27
Const CSIDL_PROFILE                    As Long = &H28
Const CSIDL_PROGRAM_FILES_COMMON       As Long = &H2B
Const CSIDL_COMMON_TEMPLATES           As Long = &H2D
Const CSIDL_COMMON_DOCUMENTS           As Long = &H2E
Const CSIDL_COMMON_ADMINTOOLS          As Long = &H2F
Const CSIDL_ADMINTOOLS                 As Long = &H30
Const CSIDL_CONNECTIONS                As Long = &H31
Const CSIDL_COMMON_MUSIC               As Long = &H35
Const CSIDL_COMMON_PICTURES            As Long = &H36
Const CSIDL_COMMON_VIDEO               As Long = &H37
Const CSIDL_RESOURCES                  As Long = &H38
Const CSIDL_RESOURCES_LOCALIZED        As Long = &H39
Const CSIDL_COMMON_OEM_LINKS           As Long = &H3A
Const CSIDL_CDBURN_AREA                As Long = &H3B
Const CSIDL_COMPUTERSNEARME            As Long = &H3D
'========================================================
' Frees memory allocated by SHBrowseForFolder()
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

' Displays a dialog box that enables the user to select a shell folder.
' Returns a pointer to an item identifier list that specifies the location
' of the selected folder relative to the root of the name space. If the user
' chooses the Cancel button in the dialog box, the return value is NULL.
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
    (lpBrowseInfo As BROWSEINFO) As Long ' ITEMIDLIST

Const BFFM_INITIALIZED = 1

'Constants ending in 'A' are for Win95 ANSI
'calls; those ending in 'W' are the wide Unicode
'calls for NT.

'Sets the status text to the null-terminated
'string specified by the lParam parameter.
'wParam is ignored and should be set to 0.
Const WM_USER = &H400
Const BFFM_SETSTATUSTEXTA As Long = (WM_USER + 100)
Const BFFM_SETSTATUSTEXTW As Long = (WM_USER + 104)

'If the lParam  parameter is non-zero, enables the
'OK button, or disables it if lParam is zero.
'(docs erroneously said wParam!)
'wParam is ignored and should be set to 0.
Const BFFM_ENABLEOK As Long = (WM_USER + 101)

'Selects the specified folder. If the wParam
'parameter is FALSE, the lParam parameter is the
'PIDL of the folder to select , or it is the path
'of the folder if wParam is the C value TRUE (or 1).
'Note that after this message is sent, the browse
'dialog receives a subsequent BFFM_SELECTIONCHANGED
'message.
Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)

'specific to the PIDL method
'Undocumented call for the example. IShellFolder's
'ParseDisplayName member function should be used instead.
Private Declare Function SHSimpleIDListFromPath Lib _
   "shell32" Alias "#162" _
   (ByVal szPath As String) As Long

'specific to the STRING method
Private Declare Function LocalAlloc Lib "kernel32" _
   (ByVal uFlags As Long, _
    ByVal uBytes As Long) As Long
    
Private Declare Function LocalFree Lib "kernel32" _
   (ByVal hMem As Long) As Long

Private Declare Function lstrcpyA Lib "kernel32" _
   (lpString1 As Any, lpString2 As Any) As Long

Private Declare Function lstrlenA Lib "kernel32" _
   (lpString As Any) As Long

Const LMEM_FIXED = &H0
Const LMEM_ZEROINIT = &H40
Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)

' Contains parameters for the the SHBrowseForFolder function and receives
' information about the folder selected by the user.
Private Type BROWSEINFO   ' bi
    
    ' Handle of the owner window for the dialog box.
    hOwner As Long
    
    ' Pointer to an item identifier list (an ITEMIDLIST structure) specifying the location
    ' of the "root" folder to browse from. Only the specified folder and its subfolders
    ' appear in the dialog box. This member can be NULL, and in that case, the
    ' name space root (the desktop folder) is used.
    pidlRoot As Long
    
    ' Pointer to a buffer that receives the display name of the folder selected by the
    ' user. The size of this buffer is assumed to be MAX_PATH bytes.
    pszDisplayName As String
    
    ' Pointer to a null-terminated string that is displayed above the tree view control
    ' in the dialog box. This string can be used to specify instructions to the user.
    lpszTitle As String
    
    ' Value specifying the types of folders to be listed in the dialog box as well as
    ' other options. This member can include zero or more of the following values below.
    ulFlags As Long
    
    ' Address an application-defined function that the dialog box calls when events
    ' occur. For more information, see the description of the BrowseCallbackProc
    ' function. This member can be NULL.
    lpfn As Long
    
    ' Application-defined value that the dialog box passes to the callback function
    ' (if one is specified).
    lParam As Long
    
    ' Variable that receives the image associated with the selected folder. The image
    ' is specified as an index to the system image list.
    iImage As Long

End Type

' BROWSEINFO ulFlags values:
' Value specifying the types of folders to be listed in the dialog box as well as
' other options. This member can include zero or more of the following values:

' Only returns file system directories. If the user selects folders
' that are not part of the file system, the OK button is grayed.
Private Const BIF_RETURNONLYFSDIRS = &H1

' Does not include network folders below the domain level in the tree view control.
' For starting the Find Computer
Private Const BIF_DONTGOBELOWDOMAIN = &H2

' Includes a status area in the dialog box. The callback function can set
' the status text by sending messages to the dialog box.
Private Const BIF_STATUSTEXT = &H4

' Only returns file system ancestors. If the user selects anything other
' than a file system ancestor, the OK button is grayed.
Private Const BIF_RETURNFSANCESTORS = &H8

' Only returns computers. If the user selects anything other
' than a computer, the OK button is grayed.
Private Const BIF_BROWSEFORCOMPUTER = &H1000

' Only returns (network) printers. If the user selects anything other
' than a printer, the OK button is grayed.
Private Const BIF_BROWSEFORPRINTER = &H2000

Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, _
                                                           ByVal X As Long, _
                                                           ByVal Y As Long, _
                                                           ByVal hIcon As Long) As Boolean

Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, _
                                                              ByVal xleft As Long, _
                                                              ByVal yTop As Long, _
                                                              ByVal hIcon As Long, _
                                                              ByVal cxWidth As Long, _
                                                              ByVal cyWidth As Long, _
                                                              ByVal istepIfAniCur As Long, _
                                                              ByVal hbrFlickerFreeDraw As Long, _
                                                              ByVal diFlags As Long) As Boolean
' DrawIconEx() diFlags values:
Private Const DI_MASK = &H1
Private Const DI_IMAGE = &H2
Private Const DI_vbNormal = &H3
Private Const DI_COMPAT = &H4
Private Const DI_DEFAULTSIZE = &H8

Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" _
                              (ByVal pszPath As Any, _
                              ByVal dwFileAttributes As Long, _
                              psfi As SHFILEINFO, _
                              ByVal cbFileInfo As Long, _
                              ByVal uFlags As Long) As Long

' pszPath:
' Pointer to a buffer that contains the path and filename. Both absolute and
' relative paths are valid. If uFlags includes the SHGFI_PIDL, value pszPath
' must be the address of an ITEMIDLIST structure that contains the list of
' item identifiers that uniquely identifies the file within the shell's name space.
' This string can use either short (the 8.3 form) or long filenames.

' dwFileAttributes:
' Array of file attribute flags (FILE_ATTRIBUTE_ values). If uFlags does not
' include the SHGFI_USEFILEATTRIBUTES value, this parameter is ignored.

Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_vbNormal = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800

' psfi and cbFileInfo:
' Address and size, in bytes, of the SHFILEINFO structure that receives the file information.

' Maximun long filename path length
Private Const MAX_PATH = 260

Private Type SHFILEINFO   ' shfi
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

' uFlags:
' Flag that specifies the file information to retrieve. This parameter can
' be a combination of the following values:


' Modifies SHGFI_ICON, causing the function to retrieve the file's large icon.
Private Const SHGFI_LARGEICON = &H0&

' Modifies SHGFI_ICON, causing the function to retrieve the file's small icon.
Private Const SHGFI_SMALLICON = &H1&

' Modifies SHGFI_ICON, causing the function to retrieve the file's open icon.
' A container object displays an open icon to indicate that the container is open.
Private Const SHGFI_OPENICON = &H2&

' Modifies SHGFI_ICON, causing the function to retrieve a shell-sized icon.
' If this flag is not specified, the function sizes the icon according to the system metric values.
Private Const SHGFI_SHELLICONSIZE = &H4&

' Indicates that pszPath is the address of an ITEMIDLIST structure rather than a path name.
Private Const SHGFI_PIDL = &H8&

' Indicates that the function should use the dwFileAttributes parameter.
Private Const SHGFI_USEFILEATTRIBUTES = &H10&

' Retrieves the handle of the icon that represents the file and the index of the
' icon within the system image list. The handle is copied to the hIcon member
' of the structure specified by psfi, and the index is copied to the iIcon member.
' The return value is the handle of the system image list.
Private Const SHGFI_ICON = &H100&

' Retrieves the display name for the file. The name is copied to the szDisplayName
' member of the structure specified by psfi. The returned display name uses the
' long filename, if any, rather than the 8.3 form of the filename.
Private Const SHGFI_DISPLAYNAME = &H200&

' Retrieves the string that describes the file's type. The string is copied to the
' szTypeName member of the structure specified by psfi.
Private Const SHGFI_TYPENAME = &H400&

' Retrieves the file attribute flags. The flags are copied to the dwAttributes
' member of the structure specified by psfi.
Private Const SHGFI_ATTRIBUTES = &H800&

' Retrieves the name of the file that contains the icon representing the file.
' The name is copied to the szDisplayName member of the structure specified by psfi.
Private Const SHGFI_ICONLOCATION = &H1000&

' Returns the type of the executable file if pszPath identifies an executable file.
' To retrieve the executable file type, uFlags must specify only SHGFI_EXETYPE.
' The return value specifies the type of the executable file:
' 0                                                                       Nonexecutable file or an error condition.
' LOWORD = NE or PEHIWORD = 3.0, 3.5, or 4.0  Windows application
' LOWORD = MZHIWORD = 0                               MS-DOS .EXE, .COM or .BAT file
' LOWORD = PEHIWORD = 0                               Win32 console application
Private Const SHGFI_EXETYPE = &H2000&

' Retrieves the index of the icon within the system image list. The index is copied to the iIcon
' member of the structure specified by psfi. The return value is the handle of the system image list.
Private Const SHGFI_SYSICONINDEX = &H4000&

' Modifies SHGFI_ICON, causing the function to add the link overlay to the file's icon.
Private Const SHGFI_LINKOVERLAY = &H8000&

' Modifies SHGFI_ICON, causing the function to blend the file's icon with the system highlight color.
Private Const SHGFI_SELECTED = &H10000
'------------------ end Browse Folders API --------------------
Public Function BrowseCallbackProcStr(ByVal hwnd As Long, _
                                      ByVal uMsg As Long, _
                                      ByVal lParam As Long, _
                                      ByVal lpData As Long) As Long
                                       
  'Callback for the Browse STRING method.
 
  'On initialization, set the dialog's
  'pre-selected folder from the pointer
  'to the path allocated as bi.lParam,
  'passed back to the callback as lpData param.
 
   Select Case uMsg
      Case BFFM_INITIALIZED
      
         Call SendMessage(hwnd, BFFM_SETSELECTIONA, _
                          True, ByVal lpData)
                          
         Case Else:
         
   End Select
          
End Function



Public Sub BrowseFolders(fromName As Form, givenTextbox As TextBox, Optional caption As Variant)
    'This sub uses the functionality of the API calls wo
    'use the Browse for Folder option in Windows
    'The name of the parent form and a textbox control on it
    'has to be passed to it
    Dim BI As BROWSEINFO
    Dim pidl As Long
    Dim sPath As String * MAX_PATH
    Dim sSelPath As String
    Dim lpSelPath As Long
    
    Dim nFolder As Long
    Dim IDL As ITEMIDLIST
    Dim SHFI As SHFILEINFO
    Dim i As Integer

    Dim RootID As Long
    'SHGetSpecialFolderLocation fromName.hwnd, CSIDL_DESKTOP, RootID
    With BI
        ' The dialog's owner window...
        .hOwner = fromName.hwnd
        .pidlRoot = 0
        .lpfn = FARPROC(AddressOf BrowseCallbackProcStr)
        'Prevent GPF caused by empty textboxes
        nFolder = CSIDL_DESKTOP
        If givenTextbox.Text = "" Then
            givenTextbox.Text = CurDir
        End If
        sSelPath = UnqualifyPath(givenTextbox.Text)
        lpSelPath = LocalAlloc(LPTR, Len(sSelPath) + 1)
        CopyMemory ByVal lpSelPath, ByVal sSelPath, Len(sSelPath) + 1
        .lParam = lpSelPath
            
        ' Set the Browse dialog root folder
        'm_wCurOptIdx = 13    'Limit the top folder to Desktop
        'nFolder = GetFolderValue(m_wCurOptIdx)
        ' Fill the item id list with the pointer of the selected folder item, rtns 0 on success
        ' ==================================================
        ' If this function fails because the selected folder doesn't exist,
        ' .pidlRoot will be uninitialized & will equal 0 (CSIDL_DESKTOP)
        ' and the root will be the Desktop.
        ' DO NOT specify the CSIDL_ constants for .pidlRoot !!!!
        ' The SHBrowseForFolder() call below will generate a fatal exception
        ' (GPF) if the folder indicated by the CSIDL_ constant does not exist!!
        ' ==================================================
'        If SHGetSpecialFolderLocation(ByVal fromName.hwnd, ByVal nFolder, IDL) = NoError Then
'          .pidlRoot = IDL.mkid.cb
'        End If
        
        ' Initialize the buffer that rtns the display name of the selected folder
        .pszDisplayName = String$(MAX_PATH, 0)
        
        ' Set the dialog's banner text
        If IsMissing(caption) Then
            .lpszTitle = "Select a folder from the tree below:"  '"Browsing is limited to: " & optFolder(m_wCurOptIdx).Caption
        Else
            .lpszTitle = "Select the " & caption
        End If
        
        ' Set the type of folders to display & return
        ' -play with these option constants to see what can be returned
        '.ulFlags = GetReturnType()
        
        'Public Enum BrowseType
        '    BrowseForFolders = &H1
        '    BrowseForComputers = &H1000
        '    BrowseForPrinters = &H2000
        '    BrowseForEverything = &H4000
        'End Enum
        .ulFlags = &H1
        
        If RootID <> 0 Then .pidlRoot = RootID
        
    End With
    
    
    ' Show the Browse dialog
    pidl = SHBrowseForFolder(BI)
    
    If pidl Then
        If SHGetPathFromIDList(pidl, sPath) Then
            ' Display the path and the name of the selected folder
            givenTextbox.Text = Left$(sPath, InStr(sPath, vbNullChar) - 1)
        End If
    Else
        givenTextbox.Text = ""
    End If
    ' Frees the memory SHBrowseForFolder()
    ' allocated for the pointer to the item id list
    Call CoTaskMemFree(pidl)
   
    Call LocalFree(lpSelPath)

End Sub
Public Function UnqualifyPath(sPath As String) As String
  'Qualifying a path involves assuring that its format
  'is valid, including a trailing slash, ready for a
  'filename. Since SHBrowseForFolder will not pre-select
  'the path if it contains the trailing slash, it must be
  'removed, hence 'unqualifying' the path.
   If Len(sPath) > 0 Then
      If Right$(sPath, 1) = "\" Then
         UnqualifyPath = Left$(sPath, Len(sPath) - 1)
         Exit Function
      End If
   End If
   UnqualifyPath = sPath
   
End Function


Public Function FARPROC(pfn As Long) As Long
  
  'A dummy procedure that receives and returns
  'the value of the AddressOf operator.
 
  'This workaround is needed as you can't assign
  'AddressOf directly to a member of a user-
  'defined type, but you can assign it to another
  'long and use that (as returned here)
  FARPROC = pfn

End Function




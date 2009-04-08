[Setup]
AppName=Activelock VB6 3.6
AppVersion=3.6
AppVerName=Activelock VB6 3.6
AppPublisher=Activelock Software Group
Uninstallable=true
; set installation folder based on target Windows version
DefaultDirName={code:SetInstallDir}\Activelock_VB6_3.6
OutputBaseFilename=Activelock_VB6_Setup_3.6_April_7_2009
OutputDir=.
DefaultGroupName=Activelock Software Group
WizardImageFile=C:\ActiveLockCommunity\Images\big-side.bmp
WizardSmallImageFile=C:\ActiveLockCommunity\Images\small-logo.bmp
WizardImageBackColor=clWhite
LicenseFile=C:\ActiveLockCommunity\License Agreement\Activelock3_License.txt
;MinVersion=0,5.0.2195sp2
PrivilegesRequired=admin
Compression=lzma
SolidCompression=true
; the following must be set to allow installation into the system directory
AllowRootDirectory=true

[Types]
Name: "full"; Description: "Full installation"
Name: "compact"; Description: "Compact installation"
Name: "custom"; Description: "Custom installation"; Flags: iscustom

[Components]
Name: "bin"; Description: "Activelock Development System"; Types: full compact custom; Flags: fixed
Name: "help"; Description: "Documents"; Types: full compact
Name: "examples"; Description: "Examples"; Types: full custom
Name: "sourcecode"; Description: "Source Code"; Types: full custom

[Files]
;begin VB system files
; (Note: Scroll to the right to see the full lines!)
Source: C:\ActiveLockCommunity\Redistribution\msvbvm60.dll; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags:  regserver restartreplace uninsneveruninstall sharedfile
Source: C:\ActiveLockCommunity\Redistribution\oleaut32.dll; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags:  regserver restartreplace uninsneveruninstall sharedfile
Source: C:\ActiveLockCommunity\Redistribution\olepro32.dll; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags:  regserver restartreplace uninsneveruninstall sharedfile
Source: C:\ActiveLockCommunity\Redistribution\asycfilt.dll; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags:  restartreplace uninsneveruninstall sharedfile
Source: C:\ActiveLockCommunity\Redistribution\stdole2.tlb; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags:  restartreplace uninsneveruninstall sharedfile regtypelib
Source: C:\ActiveLockCommunity\Redistribution\comcat.dll; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags:   restartreplace uninsneveruninstall sharedfile regserver
; end VB system files

Source: C:\ActiveLockCommunity\Redistribution\Mscomctl.ocx; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: C:\ActiveLockCommunity\Redistribution\comdlg32.ocx; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: C:\ActiveLockCommunity\Redistribution\comctl32.ocx; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: C:\ActiveLockCommunity\Redistribution\tabctl32.ocx; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: C:\ActiveLockCommunity\Redistribution\msflxgrd.ocx; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: C:\ActiveLockCommunity\Redistribution\mswinsck.ocx; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver

;license file and release notes
Source: C:\ActiveLockCommunity\License Agreement\Activelock3_License.txt; DestDir: "{app}\License Agreement"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Help\ReleaseNotes3.htm; DestDir: "{app}\Help"; Flags: ignoreversion

Source: C:\ActiveLockCommunity\Help\ReleaseNotes3_files\image001.jpg; DestDir: "{app}\Help"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Help\ReleaseNotes3_files\image002.jpg; DestDir: "{app}\Help"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Help\ReleaseNotes3_files\image003.jpg; DestDir: "{app}\Help"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Help\ReleaseNotes3_files\filelist.xml; DestDir: "{app}\Help"; Flags: ignoreversion

;this file
Source: C:\ActiveLockCommunity\Setup Packages\Activelock3.6.iss; DestDir: "{app}\Inno Setup Files"; Flags: ignoreversion

;Redistribution folder
Source: C:\ActiveLockCommunity\Redistribution\activelock3.6.dll; DestDir: "{app}\Redistribution"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Redistribution\alcrypto3.dll; DestDir: "{app}\Redistribution"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Redistribution\msvbvm60.dll; DestDir: "{app}\Redistribution"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Redistribution\oleaut32.dll; DestDir: "{app}\Redistribution"; Flags: ignoreversion; OnlyBelowVersion: 0,6
Source: C:\ActiveLockCommunity\Redistribution\olepro32.dll; DestDir: "{app}\Redistribution"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Redistribution\asycfilt.dll; DestDir: "{app}\Redistribution"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Redistribution\stdole2.tlb; DestDir: "{app}\Redistribution"; Flags: ignoreversion; OnlyBelowVersion: 0,6
Source: C:\ActiveLockCommunity\Redistribution\comcat.dll; DestDir: "{app}\Redistribution"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Redistribution\Mscomctl.ocx; DestDir: "{app}\Redistribution"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Redistribution\comdlg32.ocx; DestDir: "{app}\Redistribution"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Redistribution\comctl32.ocx; DestDir: "{app}\Redistribution"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Redistribution\tabctl32.ocx; DestDir: "{app}\Redistribution"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Redistribution\msflxgrd.ocx; DestDir: "{app}\Redistribution"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Redistribution\mswinsck.ocx; DestDir: "{app}\Redistribution"; Flags: ignoreversion

;Alcrypto3 C++ folder
Source: C:\ActiveLockCommunity\Redistribution\alcrypto3.dll; DestDir: "{sys}"; Flags: sharedfile; Components: bin sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\alcrypto3.cpp; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\alcrypto3.def; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\alcrypto3.dll; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\alcrypto3.dsp; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\alcrypto3.dsw; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
;Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\alcrypto3.exp; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\alcrypto3.h; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
;Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\alcrypto3.ilk; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
;Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\alcrypto3.lib; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
;Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\alcrypto3.ncb; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
;Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\alcrypto3.obj; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
;Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\alcrypto3.opt; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
;Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\alcrypto3.pch; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
;Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\alcrypto3.pdb; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
;Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\alcrypto3.plg; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\bignum.cpp; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\bignum.h; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
;Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\bignum.obj; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\diskid32.cpp; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\diskid32.h; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
;Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\diskid32.obj; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\md5.cpp; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\md5.h; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
;Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\md5.obj; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\memory.cpp; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\memory.h; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
;Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\memory.obj; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\misc.cpp; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\misc.h; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
;Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\misc.obj; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\noise.cpp; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\noise.h; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
;Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\noise.obj; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\port32.cpp; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\port32.h; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
;Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\port32.obj; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\prime.cpp; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\prime.h; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
;Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\prime.obj; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\rand.cpp; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\rand.h; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
;Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\rand.obj; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\resource.h; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\rsa.cpp; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\rsa.h; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
;Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\rsa.obj; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\rsag.cpp; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\rsag.h; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
;Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\rsag.obj; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\sha.cpp; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\sha.h; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
;Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\sha.obj; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\vc60.idb; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\vc60.pdb; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\version.rc; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\version.res; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\windowsversion.cpp; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\windowsversion.h; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
;Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\windowsversion.obj; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\winio.cpp; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\winio.h; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode
;Source: C:\ActiveLockCommunity\Alcrypto\Alcrypto3 C++\winio.obj; DestDir: "{app}\Alcrypto3 C++"; Flags: ignoreversion; Components: sourcecode

;Src_v3 folder for Activelock source files - including Alugen - for Classic VB6 only
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\Activelock3.6.dll; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver ignoreversion; Components: bin sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\ActiveLock3.res; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\ActiveLock3.vbp; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\ActiveLock.cls; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\ActiveLockEventNotifier.cls; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\al3.vbg; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\ALtestApp.ico; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\alugen3.ini; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\alugen3.xsd; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\alugen3_6.exe; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\alugen3_6.exe.manifest; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\Alugen.vbp; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\AlugenGlobals.cls; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\Authorizations.txt; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion uninsneveruninstall; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\BlowFish.cls; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\clsCryptoAPI.cls; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\clsCryptoAPIBase64.cls; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\clsHScroll.cls; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\clsMD5.cls; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\clsRijndael.cls; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\clsShortLicenseKey.cls; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\clsShortSerial.cls; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\clsTrialCryptAPI.cls; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\dlgRegisteredLevel.frm; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\dlgRegisteredLevel.frx; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\FileKeyStore.cls; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\frmAlugenDatabase.frm; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\frmLevelManager.frm; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\frmLevelManager.frx; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\frmMain3.frm; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\frmMain3.frx; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\frmProductsStorage.frm; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\frmProductsStorage.frx; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\frmSearch.frm; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\Globals3.cls; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\IActiveLock.cls; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\IALUGenerator.cls; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\IKeyStoreProvider.cls; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\INIFile3.cls; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\INIGenerator.cls; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\Licenses.ini; DestDir: "{app}\Activelock3.6 for VB6"; Flags: onlyifdoesntexist uninsneveruninstall; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\Licenses.mdb; DestDir: "{app}\Activelock3.6 for VB6"; Flags: onlyifdoesntexist uninsneveruninstall; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\Licenses.xml; DestDir: "{app}\Activelock3.6 for VB6"; Flags: onlyifdoesntexist uninsneveruninstall; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\Locale.cls; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\MDBGenerator.cls; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\modActiveLock.bas; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\modAlugen.bas; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\modBase64.bas; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\modComputerName.bas; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\modCryptoAPI.bas; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\modMD5.bas; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\modRegisteredLevel.bas; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\modRegistryAPIs.bas; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\modSHA1.bas; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\modTrialGlobals3.bas; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\modWindowsVersion.bas; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\ProductInfo.cls; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\ProductKeyGenerator.cls; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\ProductLicense.cls; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\RegisteredLevelDB.dat; DestDir: "{app}\Activelock3.6 for VB6"; Flags: onlyifdoesntexist uninsneveruninstall; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\RegistryKeyStore.cls; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6 for VB6\XMLGenerator.cls; DestDir: "{app}\Activelock3.6 for VB6"; Flags: ignoreversion; Components: sourcecode

;help folder
Source: C:\ActiveLockCommunity\help\ActiveLock3.chm; DestDir: "{app}\help"; Flags: ignoreversion; Components: help
Source: C:\ActiveLockCommunity\help\VBTutorial3.chm; DestDir: "{app}\help"; Flags: ignoreversion; Components: help

;Change Log
Source: C:\ActiveLockCommunity\ChangeLog\VB6_ChangeLog.txt; DestDir: "{app}\ChangeLog"; Flags: ignoreversion; Components: help

;examples\VB folder
Source: C:\ActiveLockCommunity\ALTestApp3.6 for VB6\ALVB6Sample3_6.exe; DestDir: "{app}\ALTestApp3.6 for VB6"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.6 for VB6\ALVB6Sample3_6.exe.manifest; DestDir: "{app}\ALTestApp3.6 for VB6"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.6 for VB6\ALVB6Sample.vbp; DestDir: "{app}\ALTestApp3.6 for VB6"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.6 for VB6\frmMain.frm; DestDir: "{app}\ALTestApp3.6 for VB6"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.6 for VB6\frmMain.frx; DestDir: "{app}\ALTestApp3.6 for VB6"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.6 for VB6\modMain.bas; DestDir: "{app}\ALTestApp3.6 for VB6"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.6 for VB6\modRegistryAPIs.bas; DestDir: "{app}\ALTestApp3.6 for VB6"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.6 for VB6\splash.frm; DestDir: "{app}\ALTestApp3.6 for VB6"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.6 for VB6\splash.frx; DestDir: "{app}\ALTestApp3.6 for VB6"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.6 for VB6\splash1.frm; DestDir: "{app}\ALTestApp3.6 for VB6"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.6 for VB6\splash1.frx; DestDir: "{app}\ALTestApp3.6 for VB6"; Flags: ignoreversion; Components: examples

[Tasks]
Name: startmenu; Description: Add icons to Start Menu;
Name: desktopicon; Description: Add icons to Desktop;

[Icons]
Name: {commondesktop}\Activelock3 Key Generator; Filename: "{app}\Activelock3.6 for VB6\alugen3_6.exe"; Parameters: ""; WorkingDir: {app}\Activelock3.6 for VB6; IconFilename: {app}\Activelock3.6 for VB6\alugen3_6.exe; IconIndex: 0; Flags: createonlyiffileexists; Tasks: desktopicon
Name: {group}\Activelock Home Page; Filename: "http://www.activelocksoftware.com"; Tasks: startmenu
Name: {group}\Activelock Bulletin Board; Filename: "http://forum.activelocksoftware.com"; Tasks: startmenu
Name: {group}\Activelock Download Engine; Filename: "http://download.activelocksoftware.com"; Tasks: startmenu
Name: "{group}\Activelock3 Key Generator [Alugen]"; Filename: "{app}\Activelock3.6 for VB6\alugen3_6.exe"; WorkingDir: {app}\Activelock3.6 for VB6; Components: bin sourcecode; Tasks: startmenu
Name: {group}\Activelock3 Test Application; Filename: "{app}\ALTestApp3.6 for VB6\ALVB6Sample3_6.exe"; WorkingDir: {app}\ALTestApp3.6 for VB6; Components: bin examples sourcecode; Tasks: startmenu
Name: {group}\Activelock3 Release Notes; Filename: "{app}\Help\ReleaseNotes3.htm"; WorkingDir: {app}\Help; Components: bin sourcecode; Tasks: startmenu
Name: {group}\VB6 Tutorial; Filename: "{app}\Help\VBTutorial3.chm"; WorkingDir: {app}\Help; Components: bin sourcecode; Tasks: startmenu
Name: {group}\Uninstall Activelock3; Filename: {uninstallexe}; Tasks: startmenu

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
var
  Filename: String;
begin
  if CurStep = ssInstall then
  begin
    Filename := ExpandConstant('{sys}\Activelock3.5.dll');
    Filename := ExpandConstant('{sys}\Activelock3.6.dll');
    UnregisterServer(False, Filename, True);
  end;
end;
Function SetInstallDir(Param: String): String;
Var
	Version: TWindowsVersion;
Begin
	GetWindowsVersionEx(Version);
	If Version.Major > 5 Then
		Result := ExpandConstant('{commonappdata}')
	Else
		Result := ExpandConstant('{pf}')
End;

[Registry]
Root: HKLM; ValueType: string; Subkey: Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers; ValueName: {sys}\regsvr32.exe; ValueData: DisableNXShowUI
; Uncomment the following if necessary
;Root: HKCU; ValueType: string; Subkey: Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers; ValueName: "{app}\Activelock3.6 for VB6\alugen3_6.exe"; ValueData: RunAsAdmin
;Root: HKCU; ValueType: string; Subkey: Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers; ValueName: "{app}\ALTestApp3.6 for VB6\ALVB6Sample3_6.exe"; ValueData: RunAsAdmin

[Run]
Filename: "{app}\ChangeLog\VB6_ChangeLog.txt"; Description: "View v3.6 Change Log"; Flags: unchecked shellexec postinstall skipifsilent




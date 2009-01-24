[Setup]
AppName=Activelock VB2003 3.5.4
AppVersion=3.5.4
AppVerName=Activelock VB2003 3.5.4
AppPublisher=Activelock Software Group
Uninstallable=true
DefaultDirName={pf}\Activelock VB2003 3.5.4
OutputBaseFilename=Activelock_VB2003_Setup_3.5.4
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

;DLLs and library files
Source: C:\ActiveLockCommunity\Redistribution\Mscomctl.ocx; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: C:\ActiveLockCommunity\Redistribution\comdlg32.ocx; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: C:\ActiveLockCommunity\Redistribution\comctl32.ocx; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: C:\ActiveLockCommunity\Redistribution\tabctl32.ocx; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: C:\ActiveLockCommunity\Redistribution\msflxgrd.ocx; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: C:\ActiveLockCommunity\Redistribution\mswinsck.ocx; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver

;license file and release notes
Source: C:\ActiveLockCommunity\License Agreement\Activelock3_License.txt; DestDir: "{app}\License Agreement"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Help\ReleaseNotes3.htm; DestDir: "{app}\Help"; Flags: ignoreversion isreadme
Source: C:\ActiveLockCommunity\Help\ReleaseNotes3_files\image001.jpg; DestDir: "{app}\Help"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Help\ReleaseNotes3_files\image002.jpg; DestDir: "{app}\Help"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Help\ReleaseNotes3_files\image003.jpg; DestDir: "{app}\Help"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Help\ReleaseNotes3_files\filelist.xml; DestDir: "{app}\Help"; Flags: ignoreversion

;this file
Source: C:\ActiveLockCommunity\Setup Packages\Activelock3.5_VBNET.iss; DestDir: "{app}\Inno Setup Files"; Flags: ignoreversion

;C:\ActiveLockCommunity\Redistribution folder
Source: C:\ActiveLockCommunity\Redistribution\Mscomctl.ocx; DestDir: "{app}\Redistribution"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Redistribution\comdlg32.ocx; DestDir: "{app}\Redistribution"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Redistribution\comctl32.ocx; DestDir: "{app}\Redistribution"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Redistribution\tabctl32.ocx; DestDir: "{app}\Redistribution"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Redistribution\msflxgrd.ocx; DestDir: "{app}\Redistribution"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Redistribution\mswinsck.ocx; DestDir: "{app}\Redistribution"; Flags: ignoreversion

;C:\ActiveLockCommunity\ALTestApp.NET folder
Source: C:\ActiveLockCommunity\ALTestApp3.5 NET\Activelock3_5NET.dll; DestDir: "{app}\ALTestApp3.5 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.5 NET\Alcrypto3NET.dll; DestDir: "{app}\ALTestApp3.5 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.5 NET\ALTestApp3.5.NET.exe; DestDir: "{app}\ALTestApp3.5 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.5 NET\ALTestApp3.5.NET.pdb; DestDir: "{app}\ALTestApp3.5 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.5 NET\ALTestApp3.5.NET.sln; DestDir: "{app}\ALTestApp3.5 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.5 NET\ALTestApp3.5.NET.suo; DestDir: "{app}\ALTestApp3.5 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.5 NET\ALTestApp3.5.NET.vbproj; DestDir: "{app}\ALTestApp3.5 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.5 NET\ALTestApp3.5.NET.vbproj.user; DestDir: "{app}\ALTestApp3.5 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.5 NET\ALTestApp.ico; DestDir: "{app}\ALTestApp3.5 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.5 NET\AssemblyInfo.vb; DestDir: "{app}\ALTestApp3.5 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.5 NET\CRC32.vb; DestDir: "{app}\ALTestApp3.5 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.5 NET\frmMain.resx; DestDir: "{app}\ALTestApp3.5 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.5 NET\frmMain.vb; DestDir: "{app}\ALTestApp3.5 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.5 NET\modMain.vb; DestDir: "{app}\ALTestApp3.5 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.5 NET\modRegistryAPIs.vb; DestDir: "{app}\ALTestApp3.5 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.5 NET\splash.resx; DestDir: "{app}\ALTestApp3.5 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.5 NET\splash.vb; DestDir: "{app}\ALTestApp3.5 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.5 NET\splash1.designer.vb; DestDir: "{app}\ALTestApp3.5 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.5 NET\splash1.resx; DestDir: "{app}\ALTestApp3.5 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\ALTestApp3.5 NET\splash1.vb; DestDir: "{app}\ALTestApp3.5 NET"; Flags: ignoreversion; Components: examples

;C:\ActiveLockCommunity\Alugen3.5.NET folder
Source: C:\ActiveLockCommunity\Alugen3.5.NET\Activelock3_5NET.dll; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\ALCrypto3NET.dll; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\Alugen3.5NET.sln; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\Alugen3.5NET.suo; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\Alugen3.5NET.vbproj; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\Alugen3.5NET.vbproj.user; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
;Source: C:\ActiveLockCommunity\Alugen3.5.NET\Alugen3.xsd; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\Alugen3_5NET.exe; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\Alugen3_5NET.pdb; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\Alugen3_5NET_ilmerge.exe; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\Alugen3_5NET_ilmerge_gilma.glm; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\Alugen3NET.resx; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\Alugen.ico; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\AssemblyInfo.vb; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\CRC32.vb; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\CryptoHelper.vb; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\daReport.dll; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\daReportDesigner.exe; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\daReportDesigner.pdb; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\DTPschema.xsd; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\frmAlugenDB.resX; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\frmAlugenDB.vb; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\frmAlugenDB_details.resX; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\frmAlugenDB_details.vb; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\frmLevelManager.resX; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\frmLevelManager.vb; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\frmMain3.resX; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\frmMain3.vb; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\frmProductsStorage.resX; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\frmProductsStorage.vb; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\frmRegisteredLevelEdit.resX; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\frmRegisteredLevelEdit.vb; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\GridLayoutHelper.vb; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\licenses.ini; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\licenses.mdb; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\licenses.xml; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\ListViewColumnSorter.vb; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\modActivelock.vb; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\modAlugen.vb; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\modBase64.vb; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\modINIFile.vb; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\modRegisteredLevel.vb; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\NullableDateTimePicker.vb; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\RegisteredLevelDB.dat; DestDir: "{app}\Alugen3.5.NET"; Flags: onlyifdoesntexist uninsneveruninstall; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\Resourcer_NET1.1.4322.exe; DestDir: "{app}\Alugen3.5.NET"; Flags: onlyifdoesntexist uninsneveruninstall; Components: examples
Source: C:\ActiveLockCommunity\Alugen3.5.NET\Startup.vb; DestDir: "{app}\Alugen3.5.NET"; Flags: ignoreversion; Components: examples

;Src_v3\Activelock3_5NET folder
Source: C:\ActiveLockCommunity\Activelock3.5.NET\Activelock3.5.NET.sln; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\Activelock3.5.NET.suo; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\Activelock3.5.NET.vbproj; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\Activelock3.5.NET.vbproj.user; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\Activelock3_5NET.dll; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\Activelock3_5NET.pdb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\Activelock.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\ActivelockEventNotifier.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\ADSFile.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\AlugenGlobals.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\AssemblyInfo.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\CCoder.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\clsGetMACs_IfTable.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\clsRijndael.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\clsShortLicenseKey.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\clsShortSerial.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\CRC32.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\Crypto2005.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\Daytime.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\EncryptionRoutines.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\FileKeyStore.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\frmTrialC_modified.resX; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\frmTrialC_modified.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\Globals3.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\IActiveLock.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\IALUGenerator.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\IKeyStoreProvider.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\INIFile3.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\INIGenerator.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\MDBGenerator.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\modActiveLock.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\modALUGEN.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\modBase64.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\modComputerName.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\modMD5.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\modRegistryAPIs.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\modSHA1.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\modTrialGlobals3.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\mSmartCall.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\OSVersion.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\ProductInfo.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\ProductLicense.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\ProjectResources.resx; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\RegistryKeyStore.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.5.NET\XMLGenerator.vb; DestDir: "{app}\Activelock3.5.NET"; Flags: ignoreversion; Components: examples

;Alcrypto3.NET C++ folder
Source: C:\ActiveLockCommunity\Alcrypto3.NET\alcrypto3.cpp; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\alcrypto3.exp; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\alcrypto3.h; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\alcrypto3.lib; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\alcrypto3.map; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\alcrypto3.obj; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\alcrypto3.vcproj; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\alcrypto3NET.def; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\alcrypto3NET.dll; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\alcrypto3NET.exp; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\alcrypto3NET.lib; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\alcrypto3NET.ncb; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\alcrypto3NET.sln; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\alcrypto3NET.suo; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\bignum.cpp; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\bignum.h; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\bignum.obj; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\diskid32.cpp; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\diskid32.h; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\diskid32.obj; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\md5.cpp; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\md5.h; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\md5.obj; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\memory.cpp; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\memory.h; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\memory.obj; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\misc.cpp; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\misc.h; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\misc.obj; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\noise.cpp; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\noise.h; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\noise.obj; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\port32.cpp; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\port32.h; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\port32.obj; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\prime.cpp; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\prime.h; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\prime.obj; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\rand.cpp; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\rand.h; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\rand.obj; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\resource.h; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\rsa.cpp; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\rsa.h; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\rsa.obj; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\rsag.cpp; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\rsag.h; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\rsag.obj; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\sha.cpp; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\sha.h; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\sha.obj; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\vc70.idb; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\version.rc; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\version.res; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\windowsversion.cpp; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\windowsversion.h; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\windowsversion.obj; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\winio.cpp; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\winio.h; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Alcrypto3.NET\winio.obj; DestDir: "{app}\Alcrypto3.NET"; Flags: ignoreversion; Components: sourcecode

[Code]
function InitializeSetup(): Boolean;
var
ErrorCode: Integer;
NetFrameWorkInstalled : Boolean;
Result1 : Boolean;

begin
//Get the key for the .net framework if installed
NetFrameWorkInstalled :=RegKeyExists(HKLM,'SOFTWARE\Microsoft\.NETFramework\policy\v1.1');
//The .Net Framework is installed
if NetFrameWorkInstalled = true then
  begin
  Result:=true;
  //Well OK it's installed
  end;
  //The .Net Framework is not installed, the user is prompted to install
  //the .Net Framework from the web
  if NetFrameWorkInstalled = false then
    begin
    Result1 := MsgBox('This setup requires the .NET Framework. Please download and install the .NET Framework and run this setup again. Do you want to download the .NET Framework now?', mbConfirmation, MB_YESNO) = idYes;
    if Result1 =false then
      begin
      Result:=false;
      end

    else
      begin
      Result:=false;
      //Exec('http://download.microsoft.com/download/a/a/c/aac39226-8825-44ce-90e3-bf8203e74006/dotnetfx.exe','','',SW_SHOWNORMAL,ewWaitUntilTerminated,ErrorCode);
      ShellExec('open', ExpandConstant('http://download.microsoft.com/download/a/a/c/aac39226-8825-44ce-90e3-bf8203e74006/dotnetfx.exe'),'', '', SW_SHOW, ewWaitUntilTerminated, ErrorCode);
    end;
  end;
end;

[Tasks]
Name: startmenu; Description: Add icons to Start Menu;
Name: desktopicon; Description: Add icons to Desktop;

[Icons]
Name: "{commondesktop}\Activelock3.5 VB2003 Key Generator"; Filename: "{app}\Alugen3.5.NET\Alugen3.5.NET.exe"; Parameters: ""; WorkingDir: {app}\Alugen3.5.NET; IconFilename: {app}\Alugen3.5.NET\Alugen3NET.exe; IconIndex: 0; Flags: createonlyiffileexists; Tasks: desktopicon
Name: {group}\Activelock Home Page; Filename: "http://www.activelocksoftware.com"; Tasks: startmenu
Name: {group}\Activelock Bulletin Board; Filename: "http://forum.activelocksoftware.com"; Tasks: startmenu
Name: {group}\Activelock Download Engine; Filename: "http://download.activelocksoftware.com"; Tasks: startmenu
Name: "{group}\Activelock3.5 VB2003 Key Generator [Alugen3NET]"; Filename: "{app}\Alugen3.5.NET\Alugen3.5.NET.exe"; WorkingDir: {app}\Alugen3.5.NET; Components: bin sourcecode; Tasks: startmenu
Name: "{group}\Activelock3.5 VB2003 Test Application"; Filename: "{app}\ALTestApp3.5 NET\ALTestApp3.5 NET.exe"; WorkingDir: {app}\ALTestApp3.5 NET; Components: bin examples sourcecode; Tasks: startmenu
;Name: {group}\Activelock3.5.NET Release Notes; Filename: "{app}\ReleaseNotes3.htm"; WorkingDir: {app}; Components: bin sourcecode; Tasks: startmenu
;Name: {group}\VB Tutorial; Filename: "{app}\Help\VBTutorial3.chm"; WorkingDir: {app}\Help; Components: bin sourcecode; Tasks: startmenu
Name: {group}\Uninstall Activelock3.5 for VB2003; Filename: {uninstallexe}; Tasks: startmenu

[Setup]
AppName=Activelock VB2008 3.6.0.4
AppVersion=3.6.0.4
AppVerName=Activelock VB2008 3.6.0.4
AppPublisher=Activelock Software Group
Uninstallable=true
; set installation folder based on target Windows version
DefaultDirName={code:SetInstallDir}\Activelock_VB2008_3.6
OutputBaseFilename=Activelock_VB2008_Setup_3.6.0.4_Oct_10_2009
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

;DLLs and library files
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
Source: C:\ActiveLockCommunity\Setup Packages\Activelock3.6_VBNET.iss; DestDir: "{app}\Inno Setup Files"; Flags: ignoreversion

;C:\ActiveLockCommunity\Redistribution folder
Source: C:\ActiveLockCommunity\Redistribution\Mscomctl.ocx; DestDir: "{app}\Redistribution"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Redistribution\comdlg32.ocx; DestDir: "{app}\Redistribution"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Redistribution\comctl32.ocx; DestDir: "{app}\Redistribution"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Redistribution\tabctl32.ocx; DestDir: "{app}\Redistribution"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Redistribution\msflxgrd.ocx; DestDir: "{app}\Redistribution"; Flags: ignoreversion
Source: C:\ActiveLockCommunity\Redistribution\mswinsck.ocx; DestDir: "{app}\Redistribution"; Flags: ignoreversion

;Change Log
Source: C:\ActiveLockCommunity\ChangeLog\VB_NET2008_ChangeLog.txt; DestDir: "{app}\ChangeLog"; Flags: ignoreversion; Components: help

;C:\ActiveLockCommunity\ALTestApp.NET folder
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\ALTestApp3.6.NET\Activelock3_6NET.dll; DestDir: "{app}\ALTestApp3.6 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\ALTestApp3.6.NET\Activelock3_6NET.xml; DestDir: "{app}\ALTestApp3.6 NET"; Flags: ignoreversion; Components: examples
;Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\ALTestApp3.6.NET\Alcrypto3NET.dll; DestDir: "{app}\ALTestApp3.6 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\ALTestApp3.6.NET\ALTestApp3.6.NET.exe; DestDir: "{app}\ALTestApp3.6 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\ALTestApp3.6.NET\ALTestApp3.6.NET.sln; DestDir: "{app}\ALTestApp3.6 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\ALTestApp3.6.NET\ALTestApp3.6.NET.vbproj; DestDir: "{app}\ALTestApp3.6 NET"; Flags: ignoreversion; Components: examples
;Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\ALTestApp3.6.NET\ALTestApp3.6.NET.vbproj.user; DestDir: "{app}\ALTestApp3.6 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\ALTestApp3.6.NET\ALTestApp3.6.NET.vshost.exe; DestDir: "{app}\ALTestApp3.6 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\ALTestApp3.6.NET\ALTestApp3.6.NET.vshost.exe.manifest; DestDir: "{app}\ALTestApp3.6 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\ALTestApp3.6.NET\ALTestApp3.6.NET.xml; DestDir: "{app}\ALTestApp3.6 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\ALTestApp3.6.NET\ALTestApp.ico; DestDir: "{app}\ALTestApp3.6 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\ALTestApp3.6.NET\app.manifest; DestDir: "{app}\ALTestApp3.6 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\ALTestApp3.6.NET\AssemblyInfo.vb; DestDir: "{app}\ALTestApp3.6 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\ALTestApp3.6.NET\CRC32.vb; DestDir: "{app}\ALTestApp3.6 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\ALTestApp3.6.NET\frmMain.resx; DestDir: "{app}\ALTestApp3.6 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\ALTestApp3.6.NET\frmMain.vb; DestDir: "{app}\ALTestApp3.6 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\ALTestApp3.6.NET\modMain.vb; DestDir: "{app}\ALTestApp3.6 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\ALTestApp3.6.NET\modRegistryAPIs.vb; DestDir: "{app}\ALTestApp3.6 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\ALTestApp3.6.NET\splash1.designer.vb; DestDir: "{app}\ALTestApp3.6 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\ALTestApp3.6.NET\splash1.resx; DestDir: "{app}\ALTestApp3.6 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\ALTestApp3.6.NET\splash1.vb; DestDir: "{app}\ALTestApp3.6 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\ALTestApp3.6.NET\splash.resx; DestDir: "{app}\ALTestApp3.6 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\ALTestApp3.6.NET\splash.vb; DestDir: "{app}\ALTestApp3.6 NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\ALTestApp3.6.NET\My Project\app.manifest; DestDir: "{app}\ALTestApp3.6 NET\My Project"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\ALTestApp3.6.NET\My Project\Resources.Designer.vb; DestDir: "{app}\ALTestApp3.6 NET\My Project"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\ALTestApp3.6.NET\My Project\Resources.resx; DestDir: "{app}\ALTestApp3.6 NET\My Project"; Flags: ignoreversion; Components: examples

;C:\ActiveLockCommunity\Alugen3.6.NET folder
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\Activelock3_6NET.dll; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\ALCrypto3NET.dll; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\Alugen3.6NET.sln; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\Alugen3NET.vbproj; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
;Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\Alugen3NET.vbproj.user; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\Alugen3_6NET.vshost.exe; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\Alugen3_6NET.vshost.exe.manifest; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\Alugen3_6NET.exe; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\Alugen3_6NET.xml; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\Alugen3_6NET.pdb; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\Alugen3NET.Designer.vb; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\Alugen3NET.resx; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\Alugen.ico; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\Alugen.ini; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\AssemblyInfo.vb; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\CRC32.vb; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\CryptoHelper.vb; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\daReport.dll; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\daReportDesigner.exe; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\daReportDesigner.pdb; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\DTPschema.xsd; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\frmAlugenDB.resX; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\frmAlugenDB.vb; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\frmAlugenDB_details.resX; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\frmAlugenDB_details.vb; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\frmLevelManager.resX; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\frmLevelManager.vb; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\frmMain3.resX; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\frmMain3.vb; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\frmProductsStorage.resX; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\frmProductsStorage.vb; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\frmRegisteredLevelEdit.resX; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\frmRegisteredLevelEdit.vb; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\GridLayoutHelper.vb; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\licenses.ini; DestDir: "{app}\Alugen3.6.NET"; Flags: onlyifdoesntexist uninsneveruninstall; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\licenses.mdb; DestDir: "{app}\Alugen3.6.NET"; Flags: onlyifdoesntexist uninsneveruninstall; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\licenses.xml; DestDir: "{app}\Alugen3.6.NET"; Flags: onlyifdoesntexist uninsneveruninstall; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\ListViewColumnSorter.vb; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\modActivelock.vb; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\modAlugen.vb; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\modBase64.vb; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\modINIFile.vb; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\modRegisteredLevel.vb; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\NullableDateTimePicker.vb; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\RegisteredLevelDB.dat; DestDir: "{app}\Alugen3.6.NET"; Flags: onlyifdoesntexist uninsneveruninstall; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\Resourcer_NET1.1.4322.exe; DestDir: "{app}\Alugen3.6.NET"; Flags: onlyifdoesntexist uninsneveruninstall; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\Startup.vb; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\Wizard.gif; DestDir: "{app}\Alugen3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Alugen3.6.NET\My Project\app.manifest; DestDir: "{app}\Alugen3.6.NET\My Project"; Flags: ignoreversion; Components: examples

;Src_v3\Activelock3_6NET folder
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\Activelock3.6.NET.sln; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\Activelock3.6.NET.vbproj; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
;Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\Activelock3.6.NET.vbproj.user; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\Activelock3_6NET.dll; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\Activelock.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\ActivelockEventNotifier.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\ADSFile.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\AlugenGlobals.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\AssemblyInfo.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\CCoder.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\clsGetMACs_IfTable.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\clsRijndael.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\clsShortLicenseKey.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\clsShortSerial.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\CRC32.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\Crypto2005.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\Daytime.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\EncryptionRoutines.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\FileKeyStore.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\Fingerprint.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\frmTrialC_modified.resX; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\frmTrialC_modified.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\Globals3.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\IActiveLock.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\IALUGenerator.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\IKeyStoreProvider.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\INIFile3.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\INIGenerator.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\MDBGenerator.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\modActiveLock.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\modALUGEN.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\modBase64.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\modComputerName.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\modMD5.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\modRegistryAPIs.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\modSHA1.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\modTrialGlobals3.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\mSmartCall.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\OSVersion.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\ProductInfo.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\ProductLicense.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\ProjectResources.resx; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\RegistryKeyStore.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock3.6 for VB2008\Activelock3.6.NET\XMLGenerator.vb; DestDir: "{app}\Activelock3.6.NET"; Flags: ignoreversion; Components: examples

; VB2005 Solution and Project Files
Source: C:\ActiveLockCommunity\Activelock3.6.NETVB2005\Activelock3.6.NETVB2005.sln; DestDir: "{app}\Activelock3.6.NETVB2005"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6.NETVB2005\Activelock3.6.NETVB2005.vbproj; DestDir: "{app}\Activelock3.6.NETVB2005"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6.NETVB2005\Activelock3.6.NETVB2005.vbproj.user; DestDir: "{app}\Activelock3.6.NETVB2005"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6.NETVB2005\ALTestApp3.6.NETVB2005.sln; DestDir: "{app}\Activelock3.6.NETVB2005"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6.NETVB2005\ALTestApp3.6.NETVB2005.suo; DestDir: "{app}\Activelock3.6.NETVB2005"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6.NETVB2005\ALTestApp3.6.NETVB2005.vbproj; DestDir: "{app}\Activelock3.6.NETVB2005"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6.NETVB2005\Alugen3.6NETVB2005.sln; DestDir: "{app}\Activelock3.6.NETVB2005"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6.NETVB2005\Alugen3.6NETVB2005.vbproj; DestDir: "{app}\Activelock3.6.NETVB2005"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock3.6.NETVB2005\Alugen3.6NETVB2005.vbproj.user; DestDir: "{app}\Activelock3.6.NETVB2005"; Flags: ignoreversion; Components: sourcecode

[Code]
function InitializeSetup(): Boolean;
var
ErrorCode: Integer;
NetFrameWorkInstalled : Boolean;
Result1 : Boolean;

begin
//Get the key for the .net framework if installed
NetFrameWorkInstalled :=RegKeyExists(HKLM,'SOFTWARE\Microsoft\.NETFramework\policy\v2.0');
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
    Result1 := MsgBox('This setup requires the .NET Framework v2.0 Please download and install the .NET Framework v2.0 and run this setup again. Do you want to download the .NET Framework now?', mbConfirmation, MB_YESNO) = idYes;
    if Result1 =false then
      begin
      Result:=false;
      end
    else
      begin
      Result:=false;
      ShellExec('open', ExpandConstant('http://download.microsoft.com/download/5/6/7/567758a3-759e-473e-bf8f-52154438565a/dotnetfx.exe'),'', '', SW_SHOW, ewWaitUntilTerminated, ErrorCode);
    end;
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
;Root: HKCU; ValueType: string; Subkey: Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers; ValueName: "{app}\Alugen3.6.NET\Alugen3_6NET.exe"; ValueData: RunAsAdmin
;Root: HKCU; ValueType: string; Subkey: Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers; ValueName: "{app}\ALTestApp3.6 NET\ALTestApp3.6.NET.exe"; ValueData: RunAsAdmin

[Tasks]
Name: startmenu; Description: Add icons to Start Menu;
Name: desktopicon; Description: Add icons to Desktop;

[Icons]
Name: "{commondesktop}\Activelock3.6 VB2008 Key Generator"; Filename: "{app}\Alugen3.6.NET\Alugen3_6NET.exe"; Parameters: ""; WorkingDir: {app}\Alugen3.6.NET; IconFilename: {app}\Alugen3.6.NET\Alugen3_6NET.exe; IconIndex: 0; Flags: createonlyiffileexists; Tasks: desktopicon
Name: {group}\Activelock Home Page; Filename: "http://www.activelocksoftware.com"; Tasks: startmenu
Name: {group}\Activelock Bulletin Board; Filename: "http://forum.activelocksoftware.com"; Tasks: startmenu
Name: {group}\Activelock Download Page; Filename: "http://download.activelocksoftware.com"; Tasks: startmenu
Name: "{group}\Activelock3.6 VB2008 Key Generator [Alugen3NET]"; Filename: "{app}\Alugen3.6.NET\Alugen3_6NET.exe"; WorkingDir: {app}\Alugen3.6.NET; Components: bin sourcecode; Tasks: startmenu
Name: "{group}\Activelock3.6 VB2008 Test Application"; Filename: "{app}\ALTestApp3.6 NET\ALTestApp3.6.NET.exe"; WorkingDir: {app}\ALTestApp3.6 NET; Components: bin examples sourcecode; Tasks: startmenu
Name: {group}\Uninstall Activelock3.6 for VB2008; Filename: {uninstallexe}; Tasks: startmenu

[Run]
Filename: "{app}\ChangeLog\VB_NET2008_ChangeLog.txt"; Description: "View v3.6 Change Log"; Flags: unchecked shellexec postinstall skipifsilent


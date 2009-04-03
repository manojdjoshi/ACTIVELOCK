[Setup]
AppId={{F1401524-A9CC-4432-A137-4E5014CFD0D4}
AppName=Activelock Wizard 3.6
AppVersion=3.6
AppVerName=Activelock Wizard 3.6
AppPublisher=Activelock Software Group
Uninstallable=true
; set installation folder based on target Windows version
DefaultDirName={pf}\Activelock_Wizard_3.6
;DefaultDirName={code:SetInstallDir}\Activelock_Wizard_3.6
; Permissions: everyone-full
OutputBaseFilename=Activelock_Wizard_3.6_RC2_Setup_April_03_2009
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
SetupIconFile=C:\ActiveLockCommunity\Activelock Wizard\Resources\ActivelockWizrd.ico

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Types]
Name: "full"; Description: "Full installation"
Name: "compact"; Description: "Compact installation"
Name: "custom"; Description: "Custom installation"; Flags: iscustom

[Components]
Name: "bin"; Description: "Activelock Wizard Program"; Types: full compact custom; Flags: fixed
Name: "help"; Description: "Documents"; Types: full compact
Name: "examples"; Description: "Examples"; Types: full custom
Name: "sourcecode"; Description: "Source Code"; Types: full custom

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
;this file
Source: C:\ActiveLockCommunity\Setup Packages\Activelock Wizard 3.6.iss; DestDir: "{app}\Inno Setup Files"; Flags: ignoreversion
;UserGuide And Changelog
Source: C:\ActiveLockCommunity\Activelock Wizard\Userguide.doc; DestDir: "{app}\Help"; Flags: ignoreversion; Components: help
Source: C:\ActiveLockCommunity\ChangeLog\ActivelockWizard_ChangeLog.txt; DestDir: "{app}\Help"; Flags: ignoreversion; Components: help
;Code
Source: C:\ActiveLockCommunity\Activelock Wizard\Activelock Wizard.sln; DestDir: "{app}\Code"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\Activelock Wizard.vbproj; DestDir: "{app}\Code"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\Activelock Wizard 2008.sln; DestDir: "{app}\Code"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\ApplicationEvents.vb; DestDir: "{app}\Code"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\CRC32.vb; DestDir: "{app}\Code"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\frmMain.Designer.vb; DestDir: "{app}\Code"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\frmMain.resx; DestDir: "{app}\Code"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\frmMain.vb; DestDir: "{app}\Code"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\modGlobal.vb; DestDir: "{app}\Code"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\modVB6.vb; DestDir: "{app}\Code"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\modVB2003.vb; DestDir: "{app}\Code"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\modVB2005.vb; DestDir: "{app}\Code"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\ModVB2008.vb; DestDir: "{app}\Code"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\My Project\app.manifest; DestDir: "{app}\Code\My Project"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\My Project\Application.Designer.vb; DestDir: "{app}\Code\My Project"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\My Project\Application.myapp; DestDir: "{app}\Code\My Project"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\My Project\AssemblyInfo.vb; DestDir: "{app}\Code\My Project"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\My Project\Resources.Designer.vb; DestDir: "{app}\Code\My Project"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\My Project\Resources.resx; DestDir: "{app}\Code\My Project"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\My Project\Settings.Designer.vb; DestDir: "{app}\Code\My Project"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\My Project\Settings.settings; DestDir: "{app}\Code\My Project"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\Resources\activelock.ico; DestDir: "{app}\Code\Resources"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\Resources\ActivelockWizrd.ico; DestDir: "{app}\Code\Resources"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\Resources\Help.png; DestDir: "{app}\Code\Resources"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\Resources\small-logo.bmp; DestDir: "{app}\Code\Resources"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\Resources\VB6FormHeader.txt; DestDir: "{app}\Code\Resources"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\Resources\VB6FormRoutines.txt; DestDir: "{app}\Code\Resources"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\Resources\VB6FormTop.txt; DestDir: "{app}\Code\Resources"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\Resources\VB2005FormHeader.txt; DestDir: "{app}\Code\Resources"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\Resources\VB2005FormRoutines.txt; DestDir: "{app}\Code\Resources"; Flags: ignoreversion; Components: sourcecode
Source: C:\ActiveLockCommunity\Activelock Wizard\Resources\VB2005FormTop.txt; DestDir: "{app}\Code\Resources"; Flags: ignoreversion; Components: sourcecode
;examples
;VB6
Source: C:\ActiveLockCommunity\Activelock Wizard\Samples\VB6\Form1.frm; DestDir: "{app}\Examples\VB6"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock Wizard\Samples\VB6\Form1.frx; DestDir: "{app}\Examples\VB6"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock Wizard\Samples\VB6\modALVB6.bas; DestDir: "{app}\Examples\VB6"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock Wizard\Samples\VB6\Project1.exe; DestDir: "{app}\Examples\VB6"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock Wizard\Samples\VB6\Project1.vbp; DestDir: "{app}\Examples\VB6"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock Wizard\Samples\VB6\Project1.vbw; DestDir: "{app}\Examples\VB6"; Flags: ignoreversion; Components: examples
;VB2005/VB2008
Source: C:\ActiveLockCommunity\Activelock Wizard\Samples\VB2005\AL_SAMPLE_VB2005.sln; DestDir: "{app}\Examples\VB2005"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock Wizard\Samples\VB2005\AL_SAMPLE_VB2005.vbproj; DestDir: "{app}\Examples\VB2005"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock Wizard\Samples\VB2005\AssemblyInfo.vb; DestDir: "{app}\Examples\VB2005"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock Wizard\Samples\VB2005\frmMain.Designer.vb; DestDir: "{app}\Examples\VB2005"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock Wizard\Samples\VB2005\frmMain.resx; DestDir: "{app}\Examples\VB2005"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock Wizard\Samples\VB2005\frmMain.vb; DestDir: "{app}\Examples\VB2005"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock Wizard\Samples\VB2005\modALVB2005.vb; DestDir: "{app}\Examples\VB2005"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock Wizard\Samples\VB2005\My Project\Resources.Designer.vb; DestDir: "{app}\Examples\VB2005\My Project"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock Wizard\Samples\VB2005\My Project\Resources.resx; DestDir: "{app}\Examples\VB2005\My Project"; Flags: ignoreversion; Components: examples
Source: C:\ActiveLockCommunity\Activelock Wizard\Samples\VB2005\Resources\ICON.ico; DestDir: "{app}\Examples\VB2005\Resources"; Flags: ignoreversion; Components: examples

;Application
Source: "C:\ActiveLockCommunity\Activelock Wizard\Activelock Wizard.exe"; DestDir: "{app}"; Flags: ignoreversion; Components: bin

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

[Icons]
Name: "{group}\Activelock Wizard"; Filename: "{app}\Activelock Wizard.exe"
Name: "{group}\{cm:ProgramOnTheWeb,Activelock Wizard}"; Filename: "http://www.activelocksoftware.com/"
Name: "{group}\{cm:UninstallProgram,Activelock Wizard}"; Filename: "{uninstallexe}"
Name: "{commondesktop}\Activelock Wizard"; Filename: "{app}\Activelock Wizard.exe"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\Activelock Wizard"; Filename: "{app}\Activelock Wizard.exe"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\Activelock Wizard.exe"; Description: "{cm:LaunchProgram,Activelock Wizard}"; Flags: nowait postinstall skipifsilent


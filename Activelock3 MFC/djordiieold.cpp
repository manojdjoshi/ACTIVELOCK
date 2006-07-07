
  BOOL IsLockWindowsOnly()      { return (m_lLockType == lockWindows);      } 
  BOOL IsLockCompNameOnly()     { return (m_lLockType == lockComp);         } 
  BOOL IsLockHDOnly()           { return (m_lLockType == lockHD);            } 
  BOOL IsLockMACOnly()          { return (m_lLockType == lockMAC);            } 
  BOOL IsLockHDFirmwareOnly()   { return (m_lLockType == lockHDFirmware);   } 
  BOOL IsLockBIOSOnly()         { return (m_lLockType == lockBIOS);         } 
  BOOL IsLockIPOnly()           { return (m_lLockType == lockIP);            } 
  BOOL IsLockMotherboardOnly()  { return (m_lLockType == lockMotherboard);   } 

  BOOL IsLockNone()             { return (m_lLockType == lockNone);   } 
  BOOL IsLockWindows()          { return (m_lLockType != (m_lLockType & ~lockWindows));      } 
  BOOL IsLockCompName()         { return (m_lLockType != (m_lLockType & ~lockComp));         } 
  BOOL IsLockHD()               { return (m_lLockType != (m_lLockType & ~lockHD));         } 
  BOOL IsLockMAC()              { return (m_lLockType != (m_lLockType & ~lockMAC));         } 
  BOOL IsLockHDFirmware()       { return (m_lLockType != (m_lLockType & ~lockHDFirmware));   } 
  BOOL IsLockBIOS()             { return (m_lLockType != (m_lLockType & ~lockBIOS));         } 
  BOOL IsLockIP()               { return (m_lLockType != (m_lLockType & ~lockIP));         } 
  BOOL IsLockMotherboard()      { return (m_lLockType != (m_lLockType & ~lockMotherboard));} 

  long GetLockNoneCode()        { return lockNone;            } 
  long GetLockWindowsCode()     { return lockWindows;         } 
  long GetLockCompNameCode()    { return lockComp;            } 
  long GetLockHDCode()          { return lockHD;               } 
  long GetLockMACCode()         { return lockMAC;            } 
  long GetLockHDFirmwareCode()  { return lockHDFirmware;      } 
  long GetLockBIOSCode()        { return lockBIOS;            } 
  long GetLockIPCode()          { return lockIP;               } 
  long GetLockMotherboardCode() { return lockMotherboard;      } 

  void SetLicTypeNone()         { m_lLicType = allicNone;         } 
  void SetLicTypePeriodic()     { m_lLicType = allicPeriodic;      } 
  void SetLicTypePermanent()    { m_lLicType = allicPermanent;   } 
  void SetLicTypeTimeLocked()   { m_lLicType = allicTimeLocked;   } 

  BOOL IsLicTypeNone()          { return (m_lLicType == allicNone);         } 
  BOOL IsLicTypePeriodic()      { return (m_lLicType == allicPeriodic);   } 
  BOOL IsLicTypePermanent()     { return (m_lLicType == allicPermanent);   } 
  BOOL IsLicTypeTimeLocked()    { return (m_lLicType == allicTimeLocked);   } 

  long GetLicTypeNoneCode()     { return allicNone;         } 
  long GetLicTypePeriodicCode() { return allicPeriodic;      } 
  long GetLicTypePermanentCode()   { return allicPermanent;   } 
  long GetLicTypeTimeLockedCode()  { return allicTimeLocked;   } 

  void SetLockNone()            { m_lLockType = lockNone;          } 
  void SetLockWindows()         { m_lLockType = lockWindows;       } 
  void SetLockCompName()        { m_lLockType = lockComp;          } 
  void SetLockHD()              { m_lLockType = lockHD;            } 
  void SetLockMAC()             { m_lLockType = lockMAC;           } 
  void SetLockHDFirmware()      { m_lLockType = lockHDFirmware;    } 
  void SetLockBIOS()            { m_lLockType = lockBIOS;          } 
  void SetLockIP()              { m_lLockType = lockIP;            } 
  void SetLockMotherboard()     { m_lLockType = lockMotherboard;   } 

  void PutLockNone()            { m_lLockType = lockNone;         PutLockType();  } 
  void PutLockWindows()         { m_lLockType = lockWindows;      PutLockType();  } 
  void PutLockCompName()        { m_lLockType = lockComp;         PutLockType();  } 
  void PutLockHD()              { m_lLockType = lockHD;           PutLockType();  } 
  void PutLockMAC()             { m_lLockType = lockMAC;          PutLockType();  } 
  void PutLockHDFirmware()      { m_lLockType = lockHDFirmware;   PutLockType();  } 
  void PutLockBIOS()            { m_lLockType = lockBIOS;         PutLockType();  } 
  void PutLockIP()              { m_lLockType = lockIP;           PutLockType();  } 
  void PutLockMotherboard()     { m_lLockType = lockMotherboard;  PutLockType();  } 

  CString GetInstallationNoneCode()        { return m_strInstallationCode = GetInstallationCode(m_strRegisteredUser);            } 
  CString GetInstallationWindowsCode()     { return m_strInstallationCode = GetInstallationCode(m_strRegisteredUser);         } 
  CString GetInstallationCompNameCode()    { return m_strInstallationCode = GetInstallationCode(m_strRegisteredUser);            } 
  CString GetInstallationHDCode()          { return m_strInstallationCode = GetInstallationCode(m_strRegisteredUser);               } 
  CString GetInstallationMACCode()         { return m_strInstallationCode = GetInstallationCode(m_strRegisteredUser);            } 
  CString GetInstallationHDFirmwareCode()  { return m_strInstallationCode = GetInstallationCode(m_strRegisteredUser);      } 
  CString GetInstallationBIOSCode()        { return m_strInstallationCode = GetInstallationCode(m_strRegisteredUser);            } 
  CString GetInstallationIPCode()          { return m_strInstallationCode = GetInstallationCode(m_strRegisteredUser);               } 
  CString GetInstallationMotherboardCode() { return m_strInstallationCode = GetInstallationCode(m_strRegisteredUser);      } 

  BOOL IsSingleUser()           { return (m_lFlagUsers == lFlagSingleUser);   } 
  BOOL IsMultiUser()            { return (m_lFlagUsers == lFlagMultiUser);   } 

  long GetFlagSingleUserCode()  { return lFlagSingleUser;      } 
  long GetFlagMultiUserCode()   { return lFlagMultiUser;      } 

  BOOL IsTrialTypeNone()        { return (m_lTrialType == trialNone);   } 
  BOOL IsTrialTypeDays()        { return (m_lTrialType == trialDays);   } 
  BOOL IsTrialTypeRuns()        { return (m_lTrialType == trialRuns);   } 

  long GetTrialTypeNoneCode()   { return trialNone;   } 
  long GetTrialTypeDaysCode()   { return trialDays;   } 
  long GetTrialTypeRunsCode()   { return trialRuns;   } 

  void SetTrialLength(long length)   { m_lTrialLength = length; }
  void SetTrialTypeNone()            { m_lTrialType = trialNone; } 
  void SetTrialTypeDays(long lDays); 
  void SetTrialTypeRuns(long lRuns); 
  void SetTrialHideSteganography()  { m_lTrialHideType = trialSteganography;   } 
  void SetTrialHideHiddenFolder()   { m_lTrialHideType = trialHiddenFolder;   } 
  void SetTrialHideRegistry()       { m_lTrialHideType = trialRegistry;         } 

  BOOL IsTrialHideSteganography()   { return (m_lTrialHideType == trialSteganography);   } 
  BOOL IsTrialHideHiddenFolder()    { return (m_lTrialHideType == trialHiddenFolder);   } 
  BOOL IsTrialHideRegistry()        { return (m_lTrialHideType == trialRegistry);      } 

  long GetTrialHideSteganographyCode()      { return trialSteganography;   } 
  long GetTrialHideHiddenFolderCode()       { return trialHiddenFolder;   } 
  long GetTrialHideRegistryCode()           { return trialRegistry;         } 


  inline void CActiveLockMFC::SetTrialTypeDays(long lDays) 
  { 
    if(lDays > 0) 
    { 
      m_lTrialType   = trialDays; 
      m_lTrialLength   = lDays; 
    } 
    else 
    { 
      SetTrialTypeNone(); 
      m_lTrialLength   = 0; 
    } 
  } 


  inline void CActiveLockMFC::SetTrialTypeRuns(long lRuns) 
  { 
    if(lRuns > 0) 
    { 
      m_lTrialType   = trialRuns; 
      m_lTrialLength   = lRuns; 
    } 
    else 
    { 
      SetTrialTypeNone(); 
      m_lTrialLength   = 0; 
    } 
  } 

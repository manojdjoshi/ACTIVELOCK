#pragma once 


inline CString S(char *c)
{
  return c ? c : "";
}

// This class wraps the class _IActivelock contained in the dll
// Only changes - procedures which require a license all provide a NLL ptr for this parameter

class CActiveLockWrap 
{ 
public: 
  CActiveLockWrap(){}; 
  ~CActiveLockWrap(){}; 

public: 
  _IActiveLockPtr&            ActiveLock()              { return m_ActiveLock;      } 
  const _IActiveLockPtr&      ActiveLock()      const   { return m_ActiveLock;      } 

  const CString    GetRegisteredLevel()  const                        {return S(m_ActiveLock->GetRegisteredLevel()); } 

  void             PutLockType(enum ALLockTypes arg)                  {m_ActiveLock->PutLockType(arg); };
  void             PutTrialHideType ( enum ALTrialHideTypes arg )     {m_ActiveLock->PutTrialHideType(arg); };
  void             PutTrialType (enum ALTrialTypes arg )              {m_ActiveLock->PutTrialType(arg); };
  void             PutTrialLength (long arg )                         {m_ActiveLock->PutTrialLength(arg); };
  enum ALLockTypes      GetLockType()       const                     {return m_ActiveLock->GetLockType(); };
  enum ALTrialHideTypes GetTrialHideType ( )const                     {return m_ActiveLock->GetTrialHideType(); };
  enum ALTrialTypes     GetTrialType ( )    const                     {return m_ActiveLock->GetTrialType(); };
  long             GetTrialLength() const                             {return m_ActiveLock->GetTrialLength();  };

  void             PutSoftwareName (const CString& arg)               {m_ActiveLock->PutSoftwareName(LPCTSTR(arg)); };
  void             PutSoftwarePassword (const CString& arg)           {m_ActiveLock->PutSoftwarePassword(LPCTSTR(arg)); };
  const CString    GetSoftwareName()      const                       { return S(m_ActiveLock->GetSoftwareName());    };
  const CString    GetSoftwarePassword()  const                       { return S(m_ActiveLock->GetSoftwarePassword());    };
  void             PutSoftwareCode (const CString& arg)               {m_ActiveLock->PutSoftwareCode(LPCTSTR(arg)); };
  void             PutSoftwareVersion (const CString& arg)            {m_ActiveLock->PutSoftwareVersion(LPCTSTR(arg)); };
  const CString    GetSoftwareVersion()   const                       { return S(m_ActiveLock->GetSoftwareVersion());    } ;
  void             PutKeyStoreType (enum LicStoreType arg)            {m_ActiveLock->PutKeyStoreType(arg); };
  void             PutKeyStorePath (const CString&  arg )             {m_ActiveLock->PutKeyStorePath(LPCTSTR(arg)); };
  const CString    GetInstallationCode(const CString& strUsn) const; 
  void             PutAutoRegisterKeyPath (const CString& arg)        {m_ActiveLock->PutAutoRegisterKeyPath(LPCTSTR(arg)); };
  void             Transfer (const CString& arg)                      {m_ActiveLock->Transfer(LPCTSTR(arg)); };
  const CString    InitBSTR(); 
  const CString    AcquireBSTR(
        CString & strRemainingTrialDays,
        CString & strRemainingTrialRuns,
        CString & strTrialLength,
        CString & strUsedDays,
        CString & strExpirationDate,
        CString & strRegisteredUser,
        CString & strRegisteredLevel,
        CString & strLicenseClass,
        CString & strMaxCount,
        CString & strLicenseFileType,
        CString & strLicenseType,
        CString & strUsedLockType ); 
  void             ResetTrial ()                                      {m_ActiveLock->ResetTrial(); };
  void             KillTrial ()                                       {m_ActiveLock->KillTrial(); };
  _ActiveLockEventNotifier* GetEventNotifier () const                 {return m_ActiveLock->GetEventNotifier(); };
  long             GetUsedDays()            const                     {return m_ActiveLock->GetUsedDays();  };
  const CString    GetRegisteredDate()      const                     {return S(m_ActiveLock->GetRegisteredDate());    }; 
  const CString    GetRegisteredUser()      const                     {return S(m_ActiveLock->GetRegisteredUser());    };
  const CString    GetExpirationDate()      const                     {return S(m_ActiveLock->GetExpirationDate());    };


  void DisconnectEvent(); 
  void DestroyActiveLock(); 

public: 
  //DW should be protected but not certain being used in about.cpp correctly 
  CString               m_strInitAnswer; 
  CString               m_strAcquireAnswer; 

protected: 
  CString& BSTRToCString(BSTR& bs);

  _IActiveLockPtr       m_ActiveLock;
  CActiveLockEventSink  m_ActiveLockEventSink;
}; 


// the user can change both Registered User name and lockType
// this class handles these 2 variables

class CActiveLockWrapEx : public CActiveLockWrap
{ 
public: 
  CActiveLockWrapEx(){}; 
  ~CActiveLockWrapEx(){}; 

public: 

  void             SetRegisteredUser(CString& user)                   {m_strRegisteredUser = user;};
  const CString    RegisteredUser()      const                        {return m_strRegisteredUser;    };
  void             SetLockType(enum ALLockTypes arg)                  {m_lLockType = arg;};
  enum ALLockTypes LockType()      const                              {return m_lLockType;    };
  // following used in reading/writing value to ini file
  enum ALLockTypes LockTypeFromInt(int i)      const ;
  int              LockTypeToInt(enum ALLockTypes  alt)      const    {return (int)alt;};

protected: 

  CString               m_strRegisteredUser; 
  enum ALLockTypes      m_lLockType; 

}; 


// top level class - not really necessary
// matter of style
// this is only an example of what you can do


class CActiveLockMFC : public CActiveLockWrapEx
{ 
public: 
  CActiveLockMFC(); 
  ~CActiveLockMFC(); 

  const CString    LicenseStatus()       const      { return m_strLicenseStatus;   } 
  BOOL             RegStatus()           const      { return m_bRegStatus;        } 
  BOOL             IsTrialLicense()      const      { return m_bTrial;            } 
  BOOL             IsLimitedLicense()    const      { return m_bLimited;          } 

  void Create(); 
  BOOL Initialize(); 
  BOOL CheckLicense(); 
  BOOL Register(const CString& strLibKey); 

  void SetSoftwareVersion (const CString& arg)      {m_strSoftwareVersion = arg; PutSoftwareVersion (arg);};
  void SetTrialType (enum ALTrialTypes arg )        {m_lTrialType = arg; PutTrialType(arg); };
  CString RegisteredDate ()                         {return m_strRegisteredDate;};
  CString RemainingTrialDays ()                     {return m_strRemainingTrialDays;};
  CString RemainingTrialRuns ()                     {return m_strRemainingTrialRuns;};
  CString TrialLength ()                            {return m_strTrialLength;};
  

protected: 
public:  //dw remove me
  void Init();
  BOOL                  m_bRegStatus; 
  CString               m_strLicenseStatus; 


  //   CGlobals alGlobals; 
  BOOL                  m_bTrial; 
  BOOL                  m_bLimited; 

  CString               m_strInstallationCode ;

  CString               m_strSoftwareName; 
  CString               m_strSoftwareVersion; 
  // we do not want next variable easy to find so lets not define it
  //CString               m_strSoftwareCode; 
  CString               m_strKeyStorePath; 
  CString               m_strAutoRegisterKeyPath; 
  CString               m_strRegisteredLevel; 

  enum LicStoreType     m_licStoreType;

  enum ALLicType        m_lLicType; 
  enum ALTrialTypes     m_lTrialType; 
  enum ALTrialHideTypes m_lTrialHideType; 

  CString               m_strRegisteredDate; 
  CString               m_strExpirationDate; 
  CString               m_strUsedDays; 
  
  CString m_strRemainingTrialDays;    
  CString m_strRemainingTrialRuns;    
  CString m_strTrialLength;           
  CString m_strRegisteredUser;        
  CString m_strLicenseClass;          
  CString m_strMaxCount;              
  CString m_strLicenseFileType;       
  CString m_strLicenseType;           
  CString m_strUsedLockType;        
  
}; 


#include "stdafx.h"

#include <iostream>
#include <strstream>
#include <iomanip>
#include <string>
using namespace std;
#include "comutilfix.h"

#import "C:\\windows\\system32\\ActiveLock3.4.dll" 
using namespace ActiveLock3; 

#include "ConnectionAdvisor.h"
#include "EventSink.h"

/*----------------------------------------------------------------------------*/
/*
#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif
*/


extern IID iid_IActiveLockEventSink;



BEGIN_MESSAGE_MAP(CActiveLockEventSink, CCmdTarget)
	//{{AFX_MSG_MAP(CActiveLockEventSink)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/*----------------------------------------------------------------------------*/

BEGIN_DISPATCH_MAP(CActiveLockEventSink, CCmdTarget)
	//{{AFX_DISPATCH_MAP(CActiveLockEventSink)
	DISP_FUNCTION(CActiveLockEventSink, "ValidateValue", ValidateValue,	VT_I4, VTS_PBSTR)
  //}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

/*----------------------------------------------------------------------------*/

BEGIN_INTERFACE_MAP(CActiveLockEventSink, CCmdTarget)
	INTERFACE_PART(CActiveLockEventSink, iid_IActiveLockEventSink, Dispatch)
END_INTERFACE_MAP()

/*----------------------------------------------------------------------------*/
/*
BEGIN_EVENTSINK_MAP(CActiveLockEventSink, CCmdTarget)
    //{{AFX_EVENTSINK_MAP(CActiveLockEventSink)
   ON_EVENT_REFLECT(CActiveLockEventSink, -600, ValidateValue, VTS_PBSTR )
   //}}AFX_EVENTSINK_MAP
END_EVENTSINK_MAP()
*/

IMPLEMENT_DYNCREATE(CActiveLockEventSink, CCmdTarget)

/*----------------------------------------------------------------------------*/

CActiveLockEventSink::CActiveLockEventSink() :
					m_AppEventsAdvisor(iid_IActiveLockEventSink) 
{
//	m_pWordLauncher = NULL;
	EnableAutomation();
}

/*----------------------------------------------------------------------------*/

CActiveLockEventSink::~CActiveLockEventSink()
{
}

/*----------------------------------------------------------------------------*/

void CActiveLockEventSink::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CCmdTarget::OnFinalRelease();
}

/*----------------------------------------------------------------------------*/

HRESULT CActiveLockEventSink::ValidateValue(BSTR* pbstr) 
{
  char *pc = _com_util_fix::ConvertBSTRToString(*pbstr);
  string str(pc);
  delete pc;
  string sans = Enc(str);
  SysFreeString(*pbstr);
  BSTR bans =_com_util_fix::ConvertStringToBSTR(sans.c_str());
  *pbstr=::SysAllocString(bans);
  return TRUE;
}

string CActiveLockEventSink::Enc(string strData){
  // 2 versions - choose one
  if(1){
    int i; int n;
    int numbs[]={13,29,7,5,2,9,3,1,17,11,7,13,29,31,6,17,5,13,23,47,13,29,13,19,7,1,11};
    string sResult;
    n = (int)strData.size();
    long l;
    for( i = 0; i < n; i++){
      string achar = strData.substr( i, 1);
      l = achar[0];
      switch(i%3){
          case 0:
            l *= numbs[i]; 
            break;
          case 1:
            l += numbs[i]; 
            break;
          case 2:
            l /= numbs[i];
            break;
          default:  // should n't happen
            l *= numbs[i]; 
            break;
      }
      char buf[10];
      ostrstream os(buf,10);
      os << setfill('7') << setw(2) << l << '\0' ;
      if( sResult == "" ){
        sResult = os.str();
      } else {
        sResult = sResult + os.str();
      }
    }
    return sResult;

  } else {
/*
    short len = strData.length(); 
    string sResult; 
    char xx[10]; 
    long  i; 
    for(i=0; i<len; i++) { 
      sprintf(xx, "%X", strData[i]*11); 
      sResult+=xx; 
      if(i<len-1) { 
        sResult+="."; 
      } 
    } 
    return sResult; 
*/
  }
    return strData; // not possible
}

/*----------------------------------------------------------------------------*/


BOOL CActiveLockEventSink::Advise(IUnknown* pSource, REFIID iid)
{
	// This GetInterface does not AddRef

	IUnknown* pUnknownSink = GetInterface(&IID_IUnknown);
	if (pUnknownSink == NULL)
	{
		return FALSE;
	}

	if (iid == iid_IActiveLockEventSink)
	{
		return m_AppEventsAdvisor.Advise(pUnknownSink, pSource);
	}
	else 
	{
		return FALSE;
	}
}

/*----------------------------------------------------------------------------*/
	
BOOL CActiveLockEventSink::Unadvise(REFIID iid)
{
	if (iid == iid_IActiveLockEventSink)
	{
		return m_AppEventsAdvisor.Unadvise();
	}
	else 
	{
		return FALSE;
	}
}

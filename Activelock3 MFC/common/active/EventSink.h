#pragma once

/*----------------------------------------------------------------------------*/

class CActiveLockEventSink : public CCmdTarget 
{
	DECLARE_DYNCREATE(CActiveLockEventSink)

public:
	CActiveLockEventSink();
	virtual ~CActiveLockEventSink();
	BOOL Advise(IUnknown* pSource, REFIID iid);
	BOOL Unadvise(REFIID iid);

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CActiveLockEventSink)
	public:
	virtual void OnFinalRelease();
	//}}AFX_VIRTUAL

protected:

	// Generated message map functions
	//{{AFX_MSG(CActiveLockEventSink)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CActiveLockEventSink)
	afx_msg HRESULT ValidateValue(BSTR* str);
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
  //DECLARE_EVENTSINK_MAP()

private:
  string Enc(string strData);

  CConnectionAdvisor m_AppEventsAdvisor;
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.


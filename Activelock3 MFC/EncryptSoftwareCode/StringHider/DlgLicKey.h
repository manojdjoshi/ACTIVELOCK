#pragma once
#include "afxwin.h"


// DlgLicKey dialog

class DlgLicKey : public CDialog
{
	DECLARE_DYNAMIC(DlgLicKey)

public:
	DlgLicKey(CWnd* pParent = NULL);   // standard constructor
	virtual ~DlgLicKey();

// Dialog Data
	enum { IDD = IDD_DIALOG1 };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
  CString licKey;
  int lang;
};

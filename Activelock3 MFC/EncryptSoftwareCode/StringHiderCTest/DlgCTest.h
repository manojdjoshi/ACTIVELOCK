#pragma once


// DlgCTest dialog

class DlgCTest : public CDialog
{
	DECLARE_DYNAMIC(DlgCTest)

public:
	DlgCTest(CWnd* pParent = NULL);   // standard constructor
	virtual ~DlgCTest();

// Dialog Data
	enum { IDD = IDD_TEST };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
};

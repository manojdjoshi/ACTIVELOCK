// DlgLicKey.cpp : implementation file
//

#include "stdafx.h"
#include "StringHider.h"
#include "DlgLicKey.h"


// DlgLicKey dialog

IMPLEMENT_DYNAMIC(DlgLicKey, CDialog)
DlgLicKey::DlgLicKey(CWnd* pParent /*=NULL*/)
	: CDialog(DlgLicKey::IDD, pParent)
  , licKey(_T(""))
  , lang(0)
{
}

DlgLicKey::~DlgLicKey()
{
}

void DlgLicKey::DoDataExchange(CDataExchange* pDX)
{
  CDialog::DoDataExchange(pDX);
  DDX_Text(pDX, IDC_EDIT1, licKey);
  DDX_Radio(pDX, IDC_RADIO1, lang);
}


BEGIN_MESSAGE_MAP(DlgLicKey, CDialog)
END_MESSAGE_MAP()


// DlgLicKey message handlers

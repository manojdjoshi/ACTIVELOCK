// DlgCTest.cpp : implementation file
//

#include "stdafx.h"
#include "StringHiderCTest.h"
#include "DlgCTest.h"


// DlgCTest dialog

IMPLEMENT_DYNAMIC(DlgCTest, CDialog)
DlgCTest::DlgCTest(CWnd* pParent /*=NULL*/)
	: CDialog(DlgCTest::IDD, pParent)
{
}

DlgCTest::~DlgCTest()
{
}

void DlgCTest::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(DlgCTest, CDialog)
END_MESSAGE_MAP()


// DlgCTest message handlers

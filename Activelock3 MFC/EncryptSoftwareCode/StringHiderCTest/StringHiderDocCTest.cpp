// StringHiderDoc.cpp : implementation of the CStringHiderDoc class
//

#include "stdafx.h"
#include <strstream>
#include <iomanip>

using namespace std;
#include <stdlib.h>

#include "StringHiderCTest.h"

#include "StringHiderDocCTest.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CStringHiderDoc

IMPLEMENT_DYNCREATE(CStringHiderDoc, CDocument)

BEGIN_MESSAGE_MAP(CStringHiderDoc, CDocument)
END_MESSAGE_MAP()


// CStringHiderDoc construction/destruction

CStringHiderDoc::CStringHiderDoc()
{
	// TODO: add one-time construction code here

}

CStringHiderDoc::~CStringHiderDoc()
{
}

BOOL CStringHiderDoc::OnNewDocument()
{
	if (!CDocument::OnNewDocument())
		return FALSE;

	// TODO: add reinitialization code here
	// (SDI documents will reuse this document)

	return TRUE;
}




// CStringHiderDoc serialization

void CStringHiderDoc::Serialize(CArchive& ar)
{
	if (ar.IsStoring())
	{
		// TODO: add storing code here
	}
	else
	{
		// TODO: add loading code here
	}
}


// CStringHiderDoc diagnostics

#ifdef _DEBUG
void CStringHiderDoc::AssertValid() const
{
	CDocument::AssertValid();
}

void CStringHiderDoc::Dump(CDumpContext& dc) const
{
	CDocument::Dump(dc);
}
#endif //_DEBUG


// CStringHiderDoc commands

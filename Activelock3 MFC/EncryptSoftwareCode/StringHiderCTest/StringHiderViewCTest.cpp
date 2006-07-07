// StringHiderView.cpp : implementation of the CStringHiderView class
//

#include "stdafx.h"

#include <iostream>
#include <functional>
#include <string>
#include <fstream>
#include <vector>
#include <algorithm>
#include <strstream>
#include <iomanip>

using namespace std;



#include "StringHiderCTest.h"

#include "StringHiderDocCTest.h"
#include "StringHiderViewCTest.h"
#include "DlgCTest.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// development model of what is required
string TemplateGetIt(){
  typedef unsigned int integers;
  const int charsPerInt = 4;
  const int licSize     = 200;
  const int splitSize   = licSize/charsPerInt;
  const int ileft       = 0;
  int untwist[splitSize] ={
    43 , 35 , 6 , 15 , 23 , 8 , 5 , 4 , 32 , 20
      , 3 , 10 , 21 , 41 , 11 , 30 , 47 , 39 , 1 , 7
      , 48 , 38 , 36 , 31 , 46 , 13 , 45 , 49 , 16 , 37
      , 22 , 24 , 26 , 25 , 29 , 34 , 14 , 19 , 18 , 40
      , 33 , 12 , 44 , 42 , 2 , 28 , 9 , 17 , 27 , 0
  };
  integers licInt[splitSize] ={
    2054876766 , -1293097852 , -892807578 , -1997215074 , -493581068 , 1921290882 , -228423998 , -1403612452 , -2105367916 , -2031450402
      , -257107350 , -255169814 , -2105363732 , -1734040356 , -697276176 , -2104859450 , -865154326 , -459885838 , 1855626326 , -959263510
      , -393303906 , -2004181820 , 1819188336 , -2071821694 , 1751147110 , -597761308 , -1763679576 , -392141694 , 1458356910 , 1616818858
      , -1297193834 , -188028310 , -1565862186 , -2070351164 , 1450242254 , -191076732 , -1802343190 , -455288180 , 1925897318 , 1922358368
      , -221868322 , -623088528 , 1688912596 , -2105376126 , -1461792024 , -1368759166 , -1700352348 , -1870215468 , -966865188 , -897132386
  };
  string stLic;
  stLic.reserve(licSize);
  for(int* cit = &untwist[0]; cit != &untwist[splitSize]; cit++)
  {
    char cstr[charsPerInt+1];
    int index = *cit;
    unsigned int i = licInt[index];
    i = _rotr( i, 1 );
    const unsigned int * pi = &i;
    const char * pc = reinterpret_cast<const char*>(pi);
    for(int ic=0; ic < charsPerInt; ic++){
      cstr[ic] = *pc++;
    }
    cstr[charsPerInt] = 0;
    stLic += cstr;
    const char* pTest = stLic.c_str();  // test only
  }
  if(ileft){ stLic = stLic.substr(0, (int)stLic.length() - (charsPerInt-ileft)); }
  return stLic;
}
// end of development version

// this is the code produced by this program. On reflection might be better to have seperate test program
// rather than this 2 pass version

//  *********** b e g i n n i n g  o f  t e m p . c
// code here is all in one lump. Seperate and Hide ...
// Advice. If you are new to ActiveLock, then get it working by just using the long license key first
// When that is working implement this as its own stage
// Remember if you produce a new key, eg a new version YOU MUST redo this step
string GetIt(){
  typedef unsigned int integers;
  const int charsPerInt = 4;
  const int licSize     = 200;
  const int splitSize   = licSize/charsPerInt;
  const int ileft       = 0;
  int untwist[splitSize] ={
    43 , 35 , 6 , 15 , 23 , 8 , 5 , 4 , 32 , 20
      , 3 , 10 , 21 , 41 , 11 , 30 , 47 , 39 , 1 , 7
      , 48 , 38 , 36 , 31 , 46 , 13 , 45 , 49 , 16 , 37
      , 22 , 24 , 26 , 25 , 29 , 34 , 14 , 19 , 18 , 40
      , 33 , 12 , 44 , 42 , 2 , 28 , 9 , 17 , 27 , 0
  };
  integers licInt[splitSize] ={
    2054876766 , -1293097852 , -892807578 , -1997215074 , -493581068 , 1921290882 , -228423998 , -1403612452 , -2105367916 , -2031450402
      , -257107350 , -255169814 , -2105363732 , -1734040356 , -697276176 , -2104859450 , -865154326 , -459885838 , 1855626326 , -959263510
      , -393303906 , -2004181820 , 1819188336 , -2071821694 , 1751147110 , -597761308 , -1763679576 , -392141694 , 1458356910 , 1616818858
      , -1297193834 , -188028310 , -1565862186 , -2070351164 , 1450242254 , -191076732 , -1802343190 , -455288180 , 1925897318 , 1922358368
      , -221868322 , -623088528 , 1688912596 , -2105376126 , -1461792024 , -1368759166 , -1700352348 , -1870215468 , -966865188 , -897132386
  };
  string stLic;
  stLic.reserve(licSize);
  for(int* cit = &untwist[0]; cit != &untwist[splitSize]; cit++)
  {
    char cstr[charsPerInt+1];
    int index = *cit;
    unsigned int i = licInt[index];
    i = _rotr( i, 1 );
    const unsigned int * pi = &i;
    const char * pc = reinterpret_cast<const char*>(pi);
    for(int ic=0; ic < charsPerInt; ic++){
      cstr[ic] = *pc++;
    }
    cstr[charsPerInt] = 0;
    stLic += cstr;
    const char* pTest = stLic.c_str();  // test only
  }
  if(ileft){ stLic = stLic.substr(0, (int)stLic.length() - (charsPerInt-ileft)); }
  return stLic;
}
// ******** e n d  o f   t e m p . c

// CStringHiderView

IMPLEMENT_DYNCREATE(CStringHiderView, CView)

BEGIN_MESSAGE_MAP(CStringHiderView, CView)
  // Standard printing commands
  ON_COMMAND(ID_FILE_PRINT, CView::OnFilePrint)
  ON_COMMAND(ID_FILE_PRINT_DIRECT, CView::OnFilePrint)
  ON_COMMAND(ID_FILE_PRINT_PREVIEW, CView::OnFilePrintPreview)
  ON_COMMAND(ID_LICENSEKEY_ENTERKEY, OnLicensekeyEnterkey)
END_MESSAGE_MAP()

// CStringHiderView construction/destruction

CStringHiderView::CStringHiderView()
{
  // TODO: add construction code here

}

CStringHiderView::~CStringHiderView()
{
}

BOOL CStringHiderView::PreCreateWindow(CREATESTRUCT& cs)
{
  // TODO: Modify the Window class or styles here by modifying
  //  the CREATESTRUCT cs

  return CView::PreCreateWindow(cs);
}

// CStringHiderView drawing

void CStringHiderView::OnDraw(CDC* /*pDC*/)
{
  CStringHiderDoc* pDoc = GetDocument();
  ASSERT_VALID(pDoc);
  if (!pDoc)
    return;

  // TODO: add draw code for native data here
}


// CStringHiderView printing

BOOL CStringHiderView::OnPreparePrinting(CPrintInfo* pInfo)
{
  // default preparation
  return DoPreparePrinting(pInfo);
}

void CStringHiderView::OnBeginPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
  // TODO: add extra initialization before printing
}

void CStringHiderView::OnEndPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
  // TODO: add cleanup after printing
}


// CStringHiderView diagnostics

#ifdef _DEBUG
void CStringHiderView::AssertValid() const
{
  CView::AssertValid();
}

void CStringHiderView::Dump(CDumpContext& dc) const
{
  CView::Dump(dc);
}

CStringHiderDoc* CStringHiderView::GetDocument() const // non-debug version is inline
{
  ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(CStringHiderDoc)));
  return (CStringHiderDoc*)m_pDocument;
}
#endif //_DEBUG


// CStringHiderViewCTest message handlers


void CStringHiderView::OnLicensekeyEnterkey()
{
  typedef  int integers;

  DlgCTest dlg;
  INT_PTR nResponse = dlg.DoModal();
  if (nResponse == IDOK)
  {
    string OrigLicKey = "AAAAB3NzaC1yc2EAAAABJQAAAIB9zFJqkkUQOTGtOuzD5mVxbNED86nmu5exK2WYjcCH0nJ9BrvYnI+Vng/c3ne9u6IJ5uezRWSMnRRLA25WOlCeue7fFmnr8N763104T1pKrq/nUY/0gx8+x48kufic+NM7oGcybyLBvYAAtioTjaU23kdeWav+oCuCyYKrA2Pt/w==";
    {
      string licKey = TemplateGetIt();
      if(licKey != OrigLicKey){
        AfxMessageBox("Template reconstruction failed");
      } else {
        AfxMessageBox("Template reconstruction OK");
      }
    }
    {
      string licKey = GetIt();
      if(licKey != OrigLicKey){
        AfxMessageBox("Full reconstruction failed");
      } else {
        AfxMessageBox("Full reconstruction OK");
      }
    }
  }
  else if (nResponse == IDCANCEL)
  {
    AfxMessageBox("You have cancelled ! So I have done nothing!");
  }
}


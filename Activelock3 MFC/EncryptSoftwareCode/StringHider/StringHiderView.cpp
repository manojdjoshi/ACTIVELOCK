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



#include "StringHider.h"
#include "DlgLicKey.h"

#include "StringHiderDoc.h"
#include "StringHiderView.h"
#include ".\stringhiderview.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


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


// CStringHiderView message handlers

// A l l   T h e  S t r i n g H i d e r   C o d e

// except dialogue which is minor

// return an integral random number in the range 0 - (n - 1)
int Rand(int n)
{
  return rand() % n ;
}

void CStringHiderView::OutputIncludes(ostrstream &cd)
{
  cd << "Typical includes should be, these are the ones in the program StringHider. Not all are nescessary !" << endl;
  cd << "#include <iostream>      " << endl;
  cd << "#include <functional>    " << endl;
  cd << "#include <string>        " << endl;
  cd << "#include <fstream>       " << endl;
  cd << "#include <vector>        " << endl;
  cd << "#include <algorithm>     " << endl;
  cd << "#include <strstream>     " << endl;
  cd << "#include <iomanip>       " << endl;
  cd << "using namespace std;     " << endl;
}

void CStringHiderView::OutputVBProcs(ostrstream &cd)
{
  cd << "                                                                             " << endl;
  cd << "    ' Rotate a Long to the right the specified number of times               " << endl;
  cd << "    ' Treats long as a psudo unsigned int                                    " << endl;
  cd << "                                                                             " << endl;
  cd << "    Function RotateRight(ByVal value As Long, ByVal times As Long) As Long   " << endl;
  cd << "        Dim i As Long, signBits As Long                                      " << endl;
  cd << "                                                                             " << endl;
  cd << "        ' no need to rotate more times than required                         " << endl;
  cd << "        times = times Mod 32                                                 " << endl;
  cd << "        ' return the number if it's a multiple of 32                         " << endl;
  cd << "        If times = 0 Then RotateRight = value : Exit Function                " << endl;
  cd << "                                                                             " << endl;
  cd << "        For i = 1 To times                                                   " << endl;
  cd << "            ' remember the sign bit and bit 0                                " << endl;
  cd << "            signBits = value And &H80000001                                  " << endl;
  cd << "            ' clear those bits and shift to the right by one position        " << endl;
  cd << "            value = (value And &H7FFFFFFE) \\ 2                              " << endl;
  cd << "            ' if the number was negative, then re-insert the bit             " << endl;
  cd << "            If signBits And &H80000000 Then                                  " << endl;
  cd << "                value = value Or &H40000000                                  " << endl;
  cd << "            End If                                                           " << endl;
  cd << "            ' if bit 0 was set, then set the sign bit                        " << endl;
  cd << "            If signBits And 1 Then                                           " << endl;
  cd << "               value = value Or &H80000000                                   " << endl;                           
  cd << "            End If                                                           " << endl;
  cd << "        Next                                                                 " << endl;
  cd << "        RotateRight = value                                                  " << endl;
  cd << "    End Function                                                             " << endl;
}                                                   
void CStringHiderView::OnLicensekeyEnterkey()
{
  typedef  int integers;

  DlgLicKey dlg;
  INT_PTR nResponse = dlg.DoModal();
  if (nResponse == IDOK)
  {
    string OrigLicKey = LPCTSTR(dlg.licKey);
    string licKey = LPCTSTR(dlg.licKey);
    int lang = dlg.lang;
    const int charsPerInt    = sizeof(int);
    const int licSizeMax     = 500;                         // only needs to be big enough not accurate
    const int splitSizeMax   = licSizeMax/charsPerInt;

    // if not a multiple of integer size add some crap
    const int tmplicSize  = (int)licKey.size();
    if (!tmplicSize)
    {
      AfxMessageBox("More fun if you Enter Something. Exiting!");
      return;
    }
    if (tmplicSize > licSizeMax) 
    {
      AfxMessageBox("License is too large, recompile program with licSizeMax increased. Exiting!");
      return;
    }
    int ileft = tmplicSize%charsPerInt;
    if(ileft){
      for(int i=0; i< charsPerInt-ileft; i++){
        licKey += "x";   // could do better
      }
    }

    const int licSize  = (int)licKey.size();
    const int splitSize   = licSize/charsPerInt;
    int iend  = licSize/charsPerInt;

    // split into small strings all same size as integers
    vector<string> vs;
    for(int i=0; i < iend; i++){
      string st = licKey.substr(i*charsPerInt, charsPerInt);
      vs.push_back(st);
    }

    // convert all the small strings above to integers 
    vector<integers> vi;
    for(vector<string>::const_iterator cit = vs.begin(); cit != vs.end(); cit++)
    {
      string st = *cit;
      const char * pc = st.c_str();
      const integers * pi = reinterpret_cast<const integers*>(pc);
      integers i = *pi;
      i = _rotl( i, 1 );
      vi.push_back(i);
    }

    // ************** this bit is not nescessary but does no harm
    // just a test that the integers can all be converted back to the little strings
    // and the little strings back to the big string
    // before we do more complex things
    int debugSimpleTest = 1;  
    if(debugSimpleTest)
    {
      string stTest;
      for(vector<integers>::const_iterator cit = vi.begin(); cit != vi.end(); cit++)
      {
        char cstr[charsPerInt+1];
        integers i = *cit;
        i = _rotr( i, 1 );                    // make it look less like ASCI - could do better
        const integers * pi = &i;
        const char * pc = reinterpret_cast<const char*>(pi);
        for(int ic=0; ic < charsPerInt; ic++){
          cstr[ic] = *pc++;
        }
        cstr[charsPerInt] = 0;
        stTest += cstr;
      }
      if(ileft){ stTest = stTest.substr(0, (int)stTest.length() - (charsPerInt-ileft)); }
      if(stTest != OrigLicKey){   
        AfxMessageBox("Simple reconstruction failed");
      }
    }
    // end of test

    // twist should contain the randomised No.s 0 to splitSize-1
    // and untwist their inverse
    integers twist[splitSizeMax];
    integers untwist[splitSizeMax];
    vector <int> letsShuffle;
    for (int i = 0 ; i < splitSize ; i++ ){
      letsShuffle.push_back( i );  // non randomised Nos 0 - 
    }

    // Seed the random-number generator with current time so that
    // the numbers will be different every time we run.
    // can be useful to comment out if you have problems until working
    srand( (unsigned)time( NULL ) );

    random_shuffle( letsShuffle.begin( ), letsShuffle.end( ) , pointer_to_unary_function<int, int>(Rand));

    int ind = 0;
    for (vector <int>::iterator it = letsShuffle.begin( ) ; it != letsShuffle.end( ) ; it++ ){
      twist[ind] = *it;
      untwist[*it] = ind++;          
    }

    // test string used AAAAB3NzaC1yc2EAAAABJQAAAIB9zFJqkkUQOTGtOuzD5mVxbNED86nmu5exK2WYjcCH0nJ9BrvYnI+Vng/c3ne9u6IJ5uezRWSMnRRLA25WOlCeue7fFmnr8N763104T1pKrq/nUY/0gx8+x48kufic+NM7oGcybyLBvYAAtioTjaU23kdeWav+oCuCyYKrA2Pt/w==

    // produce the code ugh ugh

    // for this code we know the exact sizes so we do not need both licSize and licSizeMax etc.
    ostrstream cd; // code
    string sComment[] ={"//","'"};
    cd << sComment[lang] << " code here is all in one lump. Seperate and Hide ..."                  << endl;
    cd << sComment[lang] << " Advice. If you are new to ActiveLock, then get it working by just using the long license key first"   << endl;
    cd << sComment[lang] << " When that is working implement this as its own stage"                 << endl;
    cd << sComment[lang] << " Remember if you produce a new key, eg a new version YOU MUST redo this step"      << endl;
    switch(lang)
    {
    case 0:  // C / C++
      {
        OutputIncludes(cd);
        cd << "string GetIt(){"                                  << endl;

        cd << "typedef unsigned int integers;"                   << endl;
        cd << "const int charsPerInt = " << charsPerInt << ";"   << endl;
        cd << "const int licSize     = " << licSize     << ";"   << endl;  
        cd << "const int splitSize   = licSize/charsPerInt;"     << endl;     
        cd << "const int ileft       = " << ileft       << ";"   << endl;     
        cd << " int untwist[splitSize] ={"                       << endl;
        cd << "   ";
        for(int i=0; i<splitSize; i++){
          cd << untwist[i] ;
          if((i+1) % 10 == 0) cd                                 << endl ;
          if(i != (splitSize-1)) cd << " , " ;
        }
        cd << "};"                                               << endl;

        cd << " integers licInt[splitSize] ={"                   << endl;
        cd << "   ";
        for(int i=0; i<splitSize; i++){
          cd << vi[ twist[i] ] ;
          if((i+1) % 10 == 0) cd                                 << endl ;
          if(i != (splitSize-1)) cd << " , " ;
        }
        cd << "};"                                               << endl;

        cd << "string stLic;"                                                << endl;
        cd << "stLic.reserve(licSize);"                                      << endl;
        cd << "for(int* cit = &untwist[0]; cit != &untwist[splitSize]; cit++)"   << endl;
        cd << "{"                                                            << endl;
        cd << "char cstr[charsPerInt+1];"                                    << endl;
        cd << "int index = *cit;"                                            << endl;
        cd << "unsigned int i = licInt[index];"                              << endl;
        cd << "i = _rotr( i, 1 );"                                           << endl;
        cd << "const unsigned int * pi = &i;"                                << endl;
        cd << "const char * pc = reinterpret_cast<const char*>(pi);"         << endl;
        cd << "for(int ic=0; ic < charsPerInt; ic++){"                       << endl;
        cd << "cstr[ic] = *pc++;"                                            << endl;
        cd << "}"                                                            << endl;
        cd << "cstr[charsPerInt] = 0;"                                       << endl;
        cd << "stLic += cstr;"                                               << endl;
        cd << "const char* pTest = stLic.c_str();  // test only"             << endl;
        cd << "}"                                                            << endl;
        cd << "if(ileft){ stLic = stLic.substr(0, (int)stLic.length() - (charsPerInt-ileft)); }" << endl;
        cd << "return stLic;"                                                << endl;
        cd << "}"                                                            << endl;

        cd << "// add following one of the following lines to access procedure" << endl;
        cd << "const char* pLic = GetIt().c_str();"                          << endl;
        cd << "string licKey = GetIt();"                                     << endl;
        cd << "     // The idea is not to have the following 2 lines in the program                          " << endl;
        cd << "     // so remove them. Only here too check that the procedure works. When satisfied remove them or convert to comment" << endl;
        cd << "     string OrigLicKey;" << endl;
        cd << "     OrigLicKey = \"" << OrigLicKey.c_str() << "\"" << endl;

        cd << '\0';     // so that str() below produces right result
      }
      break;
    case 1: // VB
      {
        OutputVBProcs(cd);
        cd << " Function GetIt() As String                                                                  " << endl;
        cd << "     ' The idea is not to have the following 2 lines in the program                          " << endl;
        cd << "     ' so remove them. Only here too check that the procedure works. When satisfied remove them or convert to comment" << endl;
        cd << "     Dim strLicOrig As String                                                                " << endl;
        cd << "     strLicOrig = \"" << OrigLicKey.c_str() << "\"" << endl;
        cd << "     ' end of remove section"                << endl;
        cd << "Const charsPerInt = " << charsPerInt         << endl;
        cd << "Const licSize     = " << licSize             << endl;  
        cd << "Const splitSize   = licSize/charsPerInt"     << endl;     
        cd << "Const ileft       = " << ileft               << endl;  
        cd << "     Dim untwist As Integer() = { _ "        << endl;
        cd << "   ";
        for(int i=0; i<splitSize; i++){
          cd << untwist[i] ;
          if((i+1) % 10 == 0) cd << " _"                    << endl ;
          if(i != (splitSize-1)) cd << " , " ;
        }
        cd << "}" << endl;
        cd << "     Dim licInt As Long() = { _"             << endl;
        cd << "   ";
        for(int i=0; i<splitSize; i++){
          cd << vi[ twist[i] ] ;
          if((i+1) % 10 == 0) cd << " _"                    << endl ;
          if(i != (splitSize-1)) cd << " , " ;
        }
        cd << "}"                                               << endl;
        cd << "     Dim stLic As String                       " << endl;
        cd << "     Dim I As Integer                          " << endl;
        cd << "     Dim Iv As Long                            " << endl;
        cd << "     For I = 0 To splitSize - 1                " << endl;
        cd << "         Iv = untwist(I)                       " << endl;
        cd << "         Iv = licInt(Iv)                       " << endl;
        cd << "         Iv = RotateRight(Iv, 1)               " << endl;
        cd << "         Dim Ic As Integer                     " << endl;
        cd << "         For Ic = 0 To charsPerInt - 1         " << endl;
        cd << "             stLic &= Chr(Iv And 255)          " << endl;
        cd << "             Iv = Iv >> 8                      " << endl;
        cd << "         Next Ic                               " << endl;
        cd << "     Next I                                    " << endl;
        cd << "     If ileft Then                             " << endl;
        cd << "         stLic = Microsoft.VisualBasic.Left(stLic, Len(stLic) - (charsPerInt - ileft))       " << endl;
        cd << "     End If                                        " << endl;
        cd << "     ' remove following section when all working   " << endl;
        cd << "     If StrComp(stLic, strLicOrig) Then            " << endl;
        cd << "         MsgBox(\"Failed to prroduce original license string\")                              " << endl;
        cd << "         MsgBox(stLic & vbCrLf & strLicOrig)   " << endl;
        cd << "     End If                                    " << endl;
        cd << "     ' end of remove section                   " << endl;
        cd << "     GetIt = stLic                             " << endl;
        cd << " End Function                                  " << endl;

        cd << '\0';     // so that str() below produces right result
      }
      break;
    }

    {
      // output to temp.c
      ofstream of("temp.c");
      of << cd.str();
    }

    AfxMessageBox("Code has been Generated in file   temp.c");

  }
  else if (nResponse == IDCANCEL)
  {
    AfxMessageBox("You have cancelled ! So I have done nothing!");
  }
}


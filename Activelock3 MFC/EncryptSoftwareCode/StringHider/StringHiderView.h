// StringHiderView.h : interface of the CStringHiderView class
//


#pragma once



class CStringHiderView : public CView
{
protected: // create from serialization only
	CStringHiderView();
	DECLARE_DYNCREATE(CStringHiderView)

// Attributes
public:
	CStringHiderDoc* GetDocument() const;

// Operations
public:

// Overrides
	public:
	virtual void OnDraw(CDC* pDC);  // overridden to draw this view
virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
protected:
	virtual BOOL OnPreparePrinting(CPrintInfo* pInfo);
	virtual void OnBeginPrinting(CDC* pDC, CPrintInfo* pInfo);
	virtual void OnEndPrinting(CDC* pDC, CPrintInfo* pInfo);

  void OutputVBProcs(ostrstream &cd);
  void OutputIncludes(ostrstream &cd);

// Implementation
public:
	virtual ~CStringHiderView();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:

// Generated message map functions
protected:
	DECLARE_MESSAGE_MAP()
public:
  afx_msg void OnLicensekeyEnterkey();
};

#ifndef _DEBUG  // debug version in StringHiderView.cpp
inline CStringHiderDoc* CStringHiderView::GetDocument() const
   { return reinterpret_cast<CStringHiderDoc*>(m_pDocument); }
#endif


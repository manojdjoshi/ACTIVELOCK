// StringHiderDoc.h : interface of the CStringHiderDoc class
//


#pragma once

class CStringHiderDoc : public CDocument
{
protected: // create from serialization only
	CStringHiderDoc();
	DECLARE_DYNCREATE(CStringHiderDoc)

// Attributes
public:

// Operations
public:

// Overrides
	public:
	virtual BOOL OnNewDocument();
	virtual void Serialize(CArchive& ar);

// Implementation
public:
	virtual ~CStringHiderDoc();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:

// Generated message map functions
protected:
	DECLARE_MESSAGE_MAP()
};



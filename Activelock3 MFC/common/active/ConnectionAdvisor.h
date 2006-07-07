#pragma once

/*----------------------------------------------------------------------------*/

class CConnectionAdvisor  
{
public:
	CConnectionAdvisor(REFIID iid);
	BOOL Advise(IUnknown* pSink, IUnknown* pSource);
	BOOL Unadvise();
	virtual ~CConnectionAdvisor();

private:
	CConnectionAdvisor();
	CConnectionAdvisor(const CConnectionAdvisor& ConnectionAdvisor);
	REFIID m_iid;
	IConnectionPoint* m_pConnectionPoint;
	DWORD m_AdviseCookie;
};

/*----------------------------------------------------------------------------*/


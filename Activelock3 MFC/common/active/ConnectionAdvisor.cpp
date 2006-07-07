
#include "stdafx.h"
#include "ConnectionAdvisor.h"

/*----------------------------------------------------------------------------*/
/*
#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif
*/
/*----------------------------------------------------------------------------*/

CConnectionAdvisor::CConnectionAdvisor(REFIID iid) : m_iid(iid)
{
	m_pConnectionPoint = NULL;
	m_AdviseCookie = 0;
}

/*----------------------------------------------------------------------------*/

CConnectionAdvisor::~CConnectionAdvisor()
{
	Unadvise();
}


/*----------------------------------------------------------------------------*/

BOOL CConnectionAdvisor::Advise(IUnknown* pSink, IUnknown* pSource)
{	
	// Advise already done 
	if (m_pConnectionPoint != NULL)
	{
		return FALSE;
	}

	BOOL Result = FALSE;

	IConnectionPointContainer* pConnectionPointContainer;

	if (FAILED(pSource->QueryInterface(
					IID_IConnectionPointContainer,
					(void**)&pConnectionPointContainer)))
	{
		return FALSE;
	}

	if (SUCCEEDED(pConnectionPointContainer->FindConnectionPoint(m_iid, &m_pConnectionPoint)))
	{
		if (SUCCEEDED(m_pConnectionPoint->Advise(pSink, &m_AdviseCookie)))
		{
			Result = TRUE;
		}
		else
		{
			m_pConnectionPoint->Release();
			m_pConnectionPoint = NULL;
			m_AdviseCookie = 0;
		}
	}
	pConnectionPointContainer->Release();
	return Result;
}

/*----------------------------------------------------------------------------*/

BOOL CConnectionAdvisor::Unadvise()
{	
	if (m_pConnectionPoint != NULL)
	{
		HRESULT hr = m_pConnectionPoint->Unadvise(m_AdviseCookie);
		// If the server is gone, ignore the error
		// ASSERT(SUCCEEDED(hr));
		m_pConnectionPoint->Release();
		m_pConnectionPoint = NULL;
		m_AdviseCookie = 0;
	}
	return TRUE;
}

/*----------------------------------------------------------------------------*/

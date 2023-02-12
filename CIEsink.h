#include "StdAfx.h"

class ATL_NO_VTABLE CIESink:
	public CComObjectRootEx<CComSingleThreadModel>,
	public IDispEventImpl<0,CIESink,&DIID_DWebBrowserEvents2>
{
private:
	CComPtr<IUnknown> m_pUnkIE;

public:
	CIESink() {}

	HRESULT AdviseToIE( CComPtr<IUnknown> pUnkIE)
	{
		m_pUnkIE=pUnkIE;
		AtlGetObjectSourceInterface( pUnkIE,&m_libid,&m_iid,&m_wMajorVerNum,&m_wMinorVerNum) ;
		HRESULT hr=this->DispEventAdvise( pUnkIE);
		return hr ;
	}

BEGIN_COM_MAP(CIESink)
	COM_INTERFACE_ENTRY_IID(DIID_DWebBrowserEvents2,CIESink)
END_COM_MAP()

BEGIN_SINK_MAP(CIESink)
	SINK_ENTRY_EX(0,DIID_DWebBrowserEvents2,DISPID_DOCUMENTCOMPLETE,OnDocumentComplete)
END_SINK_MAP()

void _stdcall OnDocumentComplete(IDispatch*,VARIANT*)
{
	PostMessage(hWnd,WM_DOCUMENTCOMPLETE,0,0);
}
};


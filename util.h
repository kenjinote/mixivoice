HHOOK g_hHook;
LRESULT CALLBACK CBTProc(
						 int nCode,
						 WPARAM wParam,
						 LPARAM lParam)
{
	if(nCode==HCBT_ACTIVATE)
	{
        UnhookWindowsHookEx(g_hHook);
        HWND hMes=(HWND)wParam;
        RECT m,w;
        GetWindowRect(hMes,&m);
        GetWindowRect(hWnd,&w);
        SetWindowPos(
			hMes,
			hWnd,
			(w.right+w.left-m.right+m.left)/2,
			(w.bottom+w.top-m.bottom+m.top)/2,
			0,
			0,
			SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
    }
    return 0;
}

BOOL CreateBoundaryA(LPSTR lpszBoundary)
{
	GUID guid=GUID_NULL;
	HRESULT hr=UuidCreate(&guid);
	if(HRESULT_CODE(hr)!=RPC_S_OK){return FALSE;}
	if(guid==GUID_NULL){return FALSE;}
	wsprintfA(lpszBoundary,
		"{%08lX-%04X-%04x-%02X%02X"
		"-%02X%02X%02X%02X%02X%02X}",
		guid.Data1,guid.Data2,guid.Data3,
		guid.Data4[0],guid.Data4[1],
		guid.Data4[2],guid.Data4[3],
		guid.Data4[4],guid.Data4[5],
		guid.Data4[6],guid.Data4[7]);
	return TRUE;
}

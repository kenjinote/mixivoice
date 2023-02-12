#define UNICODE
#pragma comment(lib,"wininet")
#pragma comment(lib,"rpcrt4")
#pragma comment(lib,"shlwapi")

#include<atlbase.h>
#include<windows.h>
#include<shlwapi.h>
#include<string>
#include<wininet.h>
CComModule _Module;
HWND hWnd;
#define WM_DOCUMENTCOMPLETE WM_APP
#include "CIEsink.h"
#include "util.h"
TCHAR szClassName[]=TEXT("Window");
#define CONSUMERKEY ★取得してください★
#define CONSUMERSECRET ★取得してください★
#define RELEASE(x) if(x){GlobalFree(x);x=0;}

BEGIN_OBJECT_MAP(ObjectMap)
END_OBJECT_MAP()

BOOL PostMixiVoice(
				  LPCSTR lpszAccessToken,
				  LPCTSTR lpszMessage,
				  LPCTSTR lpszFilePath=0) 
{
	BOOL bRet=FALSE;
	HINTERNET hInet=0;
	HINTERNET hConn=0;
	HINTERNET hHttp=0;
	BOOL bRead=FALSE;
	DWORD nRead=0L;
	std::string sHead;
	std::string sTail;
	if(!lpszMessage&&!lpszFilePath)goto END;
	CHAR lpszBoundary[39];
	if(CreateBoundaryA(lpszBoundary)==0)goto END;
	HANDLE hFile;
	hFile=INVALID_HANDLE_VALUE;
	DWORD dwFileSize;
	dwFileSize=0;
	if(lpszFilePath)
	{
		LPSTR pFileName;
		const DWORD dwSize=WideCharToMultiByte(CP_ACP,
			0,PathFindFileName(lpszFilePath),-1,0,0,0,0);
		pFileName=(LPSTR)GlobalAlloc(GMEM_FIXED,dwSize);
		WideCharToMultiByte(CP_ACP,0,PathFindFileName(
			lpszFilePath),-1,pFileName,dwSize,0,0);
		sHead+="--";
		sHead+=lpszBoundary;
		sHead+="\r\n";
		sHead+="Content-Disposition: form-data; "
			   "name=\"photo\"; filename=\"";
		sHead+=pFileName;
		GlobalFree(pFileName);
		sHead+="\"\r\n";
		sHead+="Content-Type: application/octet-stream"
		       "\r\n\r\n";
		hFile=CreateFile(lpszFilePath,GENERIC_READ,
			FILE_SHARE_READ,0,OPEN_EXISTING,
			FILE_ATTRIBUTE_NORMAL,0);
		if(hFile==INVALID_HANDLE_VALUE)goto END;
		dwFileSize=GetFileSize(hFile,0);
	}
	hInet=InternetOpen(TEXT("PostMixiVoice"),
		INTERNET_OPEN_TYPE_DIRECT,0,0,0);
	if(!hInet)goto END;
	hConn=InternetConnect(hInet,TEXT("api.mixi-platform.com"),
		INTERNET_DEFAULT_HTTPS_PORT,
		0,0,INTERNET_SERVICE_HTTP,
		0,(DWORD_PTR)0);
	if(!hConn)goto END;
	hHttp=HttpOpenRequest(hConn,TEXT("POST"),
		TEXT("/2/voice/statuses"),
		HTTP_VERSION,0,0,
		(INTERNET_FLAG_PRAGMA_NOCACHE|
		INTERNET_FLAG_RELOAD|INTERNET_FLAG_SECURE),0);
	if(!hHttp)goto END;
	LPSTR pMessage;
	pMessage=0;
	if(lpszMessage)
	{
		const DWORD dwSize=WideCharToMultiByte(
			CP_UTF8,0,lpszMessage,-1,0,0,0,0);
		pMessage=(LPSTR)GlobalAlloc(GMEM_FIXED,dwSize);
		WideCharToMultiByte(CP_UTF8,0,lpszMessage,-1,pMessage,dwSize,0,0);
	}
	{
		std::string sHeader;
		sHeader+="Content-Type: multipart/form-data, boundary=";
		sHeader+=lpszBoundary;
		sHeader+="\r\n";
		sHeader+="Authorization: OAuth ";
		sHeader+=lpszAccessToken;
		sHeader+="\r\n";
		if(!HttpAddRequestHeadersA(hHttp,sHeader.c_str(),
			-1,HTTP_ADDREQ_FLAG_ADD))goto END;
	}
	if(lpszFilePath)sTail+="\r\n";
	if(pMessage)
	{
		sTail+="--";
		sTail+=lpszBoundary;
		sTail+="\r\nContent-Disposition: form-data; "
			"name=\"status\"\r\n\r\n";
		sTail+=pMessage;
		GlobalFree(pMessage);
		sTail+="\r\n";
	}
	sTail+="--";
	sTail+=lpszBoundary;
	sTail+="--\r\n";
	INTERNET_BUFFERS BufferIn;
	ZeroMemory(&BufferIn,sizeof(INTERNET_BUFFERS));
	BufferIn.dwStructSize=sizeof(INTERNET_BUFFERS);
	BufferIn.dwBufferTotal=sTail.size()+
		((lpszFilePath)?(sHead.size()+dwFileSize):0);
	HttpSendRequestEx(hHttp,&BufferIn,0,HSR_INITIATE,0);
	DWORD dwWrote;
	if(lpszFilePath)
	{
		InternetWriteFile(
			hHttp,sHead.c_str(),sHead.size(),&dwWrote);
		DWORD dwBytesRead;
		LPBYTE lpBuffer=(LPBYTE)GlobalAlloc(
			GMEM_FIXED,dwFileSize);
		do
		{
			ReadFile(hFile,lpBuffer,dwFileSize,&dwBytesRead,0);
			InternetWriteFile(hHttp,lpBuffer,dwBytesRead,&dwWrote);
		}
		while(dwBytesRead==dwFileSize);
		GlobalFree(lpBuffer);
	}
	InternetWriteFile(
		hHttp,sTail.c_str(),sTail.size(),&dwWrote);
	HttpEndRequest(hHttp,0,HSR_INITIATE,0);
	TCHAR szStatus[16];
	DWORD dwSize;
	dwSize=16;
	if(HttpQueryInfo(hHttp,HTTP_QUERY_STATUS_TEXT,
		szStatus,&dwSize,0))
	{
		if(dwSize>1&&szStatus[0]==TEXT('O')&&
			szStatus[1]==TEXT('K'))bRet=TRUE;
	}
END:
	if(hHttp)InternetCloseHandle(hHttp);
	if(hConn)InternetCloseHandle(hConn);
	if(hInet)InternetCloseHandle(hInet);
	if(hFile!=INVALID_HANDLE_VALUE)CloseHandle(hFile);
	return bRet;
}

BOOL GetAccessToken(LPCTSTR lpszCode,
					LPSTR* lpszAccessToken,
					LPSTR* lpszRefreshToken)
{
	BOOL bRet=FALSE;
	HINTERNET hInet=0;
	HINTERNET hConn=0;
	HINTERNET hHttp=0;
	std::string sHeader;
	std::string sRequest;
	hInet=InternetOpen(TEXT("GetAccessToken"),
		INTERNET_OPEN_TYPE_DIRECT,0,0,0);
	if(!hInet)goto END;
	hConn=InternetConnect(hInet,TEXT("secure.mixi-platform.com"),
		INTERNET_DEFAULT_HTTPS_PORT,0,0,INTERNET_SERVICE_HTTP,
		0,(DWORD_PTR)0);
	if(!hConn)goto END;
	hHttp=HttpOpenRequest(hConn,TEXT("POST"),TEXT("/2/token"),
		HTTP_VERSION,0,0,(INTERNET_FLAG_PRAGMA_NOCACHE|
		INTERNET_FLAG_RELOAD|INTERNET_FLAG_SECURE),0);
	if(!hHttp)goto END;
	sHeader+="Content-Type: application/x-www-form-urlencoded\r\n";
	if(!HttpAddRequestHeadersA(hHttp,sHeader.c_str(),
		-1,HTTP_ADDREQ_FLAG_ADD))goto END;
	LPSTR pCode;
	pCode=0;
	if(lpszCode)
	{
		const DWORD dwSize=WideCharToMultiByte(
			CP_ACP,0,lpszCode,-1,0,0,0,0);
		pCode=(LPSTR)GlobalAlloc(GMEM_FIXED,dwSize);
		WideCharToMultiByte(CP_ACP,0,lpszCode,-1,pCode,dwSize,0,0);
	}
	sRequest+="grant_type=authorization_code";
	sRequest+="&client_id="CONSUMERKEY;
	sRequest+="&client_secret="CONSUMERSECRET;
	sRequest+="&code=";
	sRequest+=pCode;
	GlobalFree(pCode);
	sRequest+="&redirect_uri=http%3A%2F%2Fwww.a.zaq.jp%2Fsdk%2Fmixi.html";
	INTERNET_BUFFERS BufferIn;
	ZeroMemory(&BufferIn,sizeof(INTERNET_BUFFERS));
	BufferIn.dwStructSize=sizeof(INTERNET_BUFFERS);
	BufferIn.dwBufferTotal=sRequest.size();
	HttpSendRequestEx(hHttp,&BufferIn,0,HSR_INITIATE,0);
	DWORD dwWrote;
	InternetWriteFile(
		hHttp,sRequest.c_str(),sRequest.size(),&dwWrote);
	HttpEndRequest(hHttp,0,HSR_INITIATE,0);
	CHAR szReadData[512];
	DWORD dwBufSize;
	dwBufSize=512;
	DWORD dwReadSize;
	if(InternetReadFile(hHttp,szReadData,dwBufSize,&dwReadSize))
	{
		szReadData[dwReadSize]=0;
		LPSTR lpszChars;
		lpszChars="{\":,}";
		LPSTR token;
		token=strtok(szReadData,lpszChars);
		BOOL bAccessTokenFlag=0;
		BOOL bRefreshTokenFlag=0;
		RELEASE(*lpszAccessToken)
		RELEASE(*lpszRefreshToken)
		while(token)
		{
			if(bAccessTokenFlag)
			{
				const DWORD dwSize=lstrlenA(token)+1;
				*lpszAccessToken=(LPSTR)GlobalAlloc(GMEM_FIXED,dwSize);
				lstrcpyA(*lpszAccessToken,token);
				bAccessTokenFlag=0;
			}
			else if(bRefreshTokenFlag)
			{
				const DWORD dwSize=lstrlenA(token)+1;
				*lpszRefreshToken=(LPSTR)GlobalAlloc(GMEM_FIXED,dwSize);
				lstrcpyA(*lpszRefreshToken,token);
				bRefreshTokenFlag=0;
			}
			if(lstrcmpiA(token,"access_token")==0)
				bAccessTokenFlag=1;
			else if(lstrcmpiA(token,"refresh_token")==0)
				bRefreshTokenFlag=1;
			else
			{
				bAccessTokenFlag=0;
				bRefreshTokenFlag=0;
			}
			token=strtok(0,lpszChars);
		}
		if(*lpszRefreshToken&&*lpszRefreshToken)bRet=TRUE;
	}
END:
	if(hHttp)InternetCloseHandle(hHttp);
	if(hConn)InternetCloseHandle(hConn);
	if(hInet)InternetCloseHandle(hInet);
	return bRet;
}

LRESULT CALLBACK WndProc(
						 HWND hWnd,
						 UINT msg,
						 WPARAM wParam,
						 LPARAM lParam)
{
	static CComQIPtr<IWebBrowser2>pWB;
	static HWND	hBrowser;
	static HWND	hEditMessage;
	static HWND	hEditFilePath;
	static HWND	hButtonSubmit;
	static HWND	hButtonFileOpen;
	static LPSTR lpszAccessToken;
	static LPSTR lpszRefreshToken;
	switch(msg)
	{
	case WM_CREATE:
		{
			AtlAxWinInit();
			hBrowser=CreateWindow(TEXT("AtlAxWin"),
				TEXT("https://mixi.jp/connect_authorize.pl")
				TEXT("?client_id=")TEXT(CONSUMERKEY)
				TEXT("&response_type=code")
				TEXT("&scope=w_voice"),
				WS_CHILD|WS_VISIBLE,0,0,0,0,hWnd,
				0,((LPCREATESTRUCT)lParam)->hInstance,0);
			hEditMessage=CreateWindowEx(WS_EX_CLIENTEDGE,
				TEXT("EDIT"),0,WS_CHILD|WS_VSCROLL|WS_TABSTOP|
				ES_MULTILINE|ES_WANTRETURN,0,0,0,0,hWnd,0,
				((LPCREATESTRUCT)lParam)->hInstance,0);
			SendMessage(hEditMessage,EM_LIMITTEXT,150,0);
			hEditFilePath=CreateWindowEx(WS_EX_CLIENTEDGE,
				TEXT("EDIT"),0,WS_CHILD|WS_TABSTOP|ES_AUTOHSCROLL,
				0,0,0,0,hWnd,0,((LPCREATESTRUCT)lParam)->hInstance,0);
			hButtonFileOpen=CreateWindow(TEXT("BUTTON"),TEXT("..."),
				WS_CHILD|WS_TABSTOP,0,0,0,0,hWnd,(HMENU)101,
				((LPCREATESTRUCT)lParam)->hInstance,0);
			hButtonSubmit=CreateWindow(TEXT("BUTTON"),TEXT("投稿"),
				WS_CHILD|WS_TABSTOP,0,0,0,0,hWnd,(HMENU)100,
				((LPCREATESTRUCT)lParam)->hInstance,0);
			DragAcceptFiles(hWnd,1);
			CComPtr<IUnknown>ie;
			if(AtlAxGetControl(hBrowser,&ie)==S_OK)
			{
				pWB=ie;
				if(pWB)
				{
					CComObject<CIESink>*sink;
					CComObject<CIESink>::CreateInstance(&sink);
					HRESULT hr=sink->AdviseToIE(ie);
					if(SUCCEEDED(hr))return 0;
				}
			}
			return -1;
		}
		break;
	case WM_DROPFILES:
		{
			HDROP hDrop=(HDROP)wParam; 
			DragQueryFile(hDrop,0xFFFFFFFF,0,0);
			TCHAR szFilePath[MAX_PATH];
			DragQueryFile(hDrop,0,szFilePath,sizeof(szFilePath));
			DragFinish(hDrop);
			SetWindowText(hEditFilePath,szFilePath);
			SendMessage(hEditFilePath,EM_SETSEL,0,-1);
			SetFocus(hEditFilePath);
		}
		break;
	case WM_DOCUMENTCOMPLETE:
		{
			BSTR HtmlURL;
			if(SUCCEEDED(pWB->get_LocationURL(&HtmlURL)))
			{
				LPTSTR p=StrStr(HtmlURL,TEXT("code="));
				if(p)
				{
					p+=5;
					LPTSTR q=p;
					while(*q)
					{
						if(*q==TEXT('&'))
						{
							*q=0;
							break;
						}
						++q;
					}
					LPTSTR lpszCode=(LPTSTR)GlobalAlloc(GMEM_FIXED,
						sizeof(TCHAR)*(lstrlen(p)+1));
					lstrcpy(lpszCode,p);
					SysFreeString(HtmlURL);
					GetAccessToken(lpszCode,&lpszAccessToken,
						&lpszRefreshToken);
					GlobalFree(lpszCode);
					ShowWindow(hBrowser,SW_HIDE);
					ShowWindow(hEditMessage,SW_SHOWNA);
					ShowWindow(hEditFilePath,SW_SHOWNA);
					ShowWindow(hButtonFileOpen,SW_SHOWNA);
					ShowWindow(hButtonSubmit,SW_SHOWNA);
				}
			}
		}
		break;
	case WM_COMMAND:
		if(LOWORD(wParam)==100)
		{
			DWORD dwMessageSize=GetWindowTextLength(hEditMessage);
			DWORD dwFilePathSize=GetWindowTextLength(hEditFilePath);
			if(!dwMessageSize&&!dwFilePathSize)return 0;
			LPTSTR lpszMessage=0;
			if(dwMessageSize)
			{
				lpszMessage=(LPTSTR)GlobalAlloc(GMEM_FIXED,
					sizeof(TCHAR)*(dwMessageSize+1));
				GetWindowText(hEditMessage,lpszMessage,dwMessageSize+1);
			}
			LPTSTR lpszFilePath=0;
			if(dwFilePathSize)
			{
				lpszFilePath=(LPTSTR)GlobalAlloc(GMEM_FIXED,
					sizeof(TCHAR)*(dwFilePathSize+1));
				GetWindowText(hEditFilePath,lpszFilePath,dwFilePathSize+1);
			}
			g_hHook=SetWindowsHookEx(WH_CBT,CBTProc,0,GetCurrentThreadId());
			if(PostMixiVoice(lpszAccessToken,lpszMessage,lpszFilePath))
			{
				MessageBox(hWnd,TEXT("投稿されました。"),TEXT("確認"),0);
				SetWindowText(hEditMessage,0);
				SetWindowText(hEditFilePath,0);
			}
			else MessageBox(hWnd,TEXT("投稿に失敗しました。"),TEXT("確認"),0);
			RELEASE(lpszMessage)
			RELEASE(lpszFilePath)
		}
		else if(LOWORD(wParam)==101)
		{
			TCHAR szFilePath[MAX_PATH];
			GetWindowText(hEditFilePath,szFilePath,MAX_PATH);
			OPENFILENAME of;
			ZeroMemory(&of,sizeof(OPENFILENAME));
			of.lStructSize=sizeof(OPENFILENAME);
			of.hwndOwner=hWnd;
			of.lpstrFilter=
				TEXT("画像ファイル\0*.jpg;*.bmp;*.png;*.gif;*.tiff\0\0");
			of.lpstrFile=szFilePath;
			of.nMaxFile=MAX_PATH;
			of.nMaxFileTitle=MAX_PATH;
			of.Flags=OFN_FILEMUSTEXIST|OFN_HIDEREADONLY;
			of.lpstrTitle=TEXT("アップロードする画像ファイルを選択");
			if(GetOpenFileName(&of))
			{
				SetWindowText(hEditFilePath,szFilePath);
				SendMessage(hEditFilePath,EM_SETSEL,0,-1);
				SetFocus(hEditFilePath);
			}
		}
		break;
	case WM_PAINT:
		{
			PAINTSTRUCT ps;
			HDC hdc=BeginPaint(hWnd,&ps);
			DWORD dwBkMode=SetBkMode(hdc,TRANSPARENT);
			RECT rect;
			GetClientRect(hWnd,&rect);
			TextOut(hdc,10,10,TEXT("メッセージ"),5);
			TextOut(hdc,10,rect.bottom-84,TEXT("画像ファイル"),6);
			SetBkMode(hdc,dwBkMode);
			EndPaint(hWnd,&ps);
		}
		break;
	case WM_SIZE:
		{
			const int w=LOWORD(lParam);
			const int h=HIWORD(lParam);
			MoveWindow(hBrowser,0,0,w,h,1);
			MoveWindow(hEditMessage,128,10,w-138,h-104,1);
			MoveWindow(hEditFilePath,128,h-84,w-170,32,1);
			MoveWindow(hButtonFileOpen,w-42,h-84,32,32,1);
			MoveWindow(hButtonSubmit,w-138,h-42,128,32,1);
		}
		break;
	case WM_CLOSE:
		DestroyWindow(hWnd);
		break;
	case WM_DESTROY:
		pWB.Release();
		RELEASE(lpszAccessToken)
		RELEASE(lpszRefreshToken)
		PostQuitMessage(0);
		break;
	default:
		return(DefDlgProc(hWnd,msg,wParam,lParam));
	}
	return 0;
}

int WINAPI WinMain(HINSTANCE hInstance,HINSTANCE hPreInst,
	LPSTR pCmdLine,int nCmdShow)
{
	MSG msg;
	_Module.Init(ObjectMap,hInstance);
	WNDCLASS wndclass={CS_HREDRAW|CS_VREDRAW,WndProc,0,DLGWINDOWEXTRA,
		hInstance,0,LoadCursor(0,IDC_ARROW),0,0,szClassName};
	RegisterClass(&wndclass);
	hWnd=CreateWindow(szClassName,TEXT("mixiボイスに投稿"),
		WS_OVERLAPPEDWINDOW,CW_USEDEFAULT,0,CW_USEDEFAULT,0,0,0,
		hInstance,0);
	ShowWindow(hWnd,nCmdShow);
	UpdateWindow(hWnd);
	while(GetMessage(&msg,0,0,0))
	{
		if(!IsDialogMessage(hWnd,&msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
	_Module.Term();
	return msg.wParam;
}

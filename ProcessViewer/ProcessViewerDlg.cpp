
// ProcessViewerDlg.cpp : implementation file
//

#include "pch.h"
#include "framework.h"
#include "ProcessViewer.h"
#include "ProcessViewerDlg.h"
#include "afxdialogex.h"
#include <Psapi.h>
#include <WbemCli.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#pragma comment(lib, "wbemuuid.lib")

// CAboutDlg dialog used for App About

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CProcessViewerDlg dialog



CProcessViewerDlg::CProcessViewerDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_PROCESSVIEWER_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CProcessViewerDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST1, ProcessList);
}

BEGIN_MESSAGE_MAP(CProcessViewerDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON_REFRESH, &CProcessViewerDlg::OnBnClickedButtonRefresh)
END_MESSAGE_MAP()


// CProcessViewerDlg message handlers

BOOL CProcessViewerDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != nullptr)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// TODO: Add extra initialization here
	ProcessList.InsertColumn(0, _T("PID"), LVCFMT_LEFT, 70);
	ProcessList.InsertColumn(1, _T("Name"), LVCFMT_LEFT, 140);
	ProcessList.InsertColumn(2, _T("Path"), LVCFMT_LEFT, 380);
	ProcessList.InsertColumn(3, _T("CmdLine"), LVCFMT_LEFT, 460);

	GetProcessInfo();

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CProcessViewerDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CProcessViewerDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CProcessViewerDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CProcessViewerDlg::OnBnClickedButtonRefresh()
{
	// TODO: Add your control notification handler code here

	ProcessList.DeleteAllItems();
	GetProcessInfo();
}

void CProcessViewerDlg::GetProcessInfo()
{
	// mot so process khong lay du thong tin - can them quyen
	HRESULT hres;

	hres = CoInitializeEx(0, COINITBASE_MULTITHREADED);
	if (FAILED(hres))
	{
		MessageBox(_T("Failed to initialize COM library."), _T("Error"), MB_OK);
		return;
	}

	hres = CoInitializeSecurity(
		NULL,
		-1,
		NULL,
		NULL,
		RPC_C_AUTHN_LEVEL_DEFAULT,
		RPC_C_IMP_LEVEL_IMPERSONATE,
		NULL,
		EOAC_NONE,
		NULL
	);
	if (FAILED(hres))
	{
		MessageBox(_T("Failed to initialize security"), _T("Error"), MB_OK);
		CoUninitialize();
		return;
	}

	IWbemLocator* pLoc = NULL;
	hres = CoCreateInstance(
		CLSID_WbemLocator,
		0,
		CLSCTX_INPROC_SERVER,
		IID_IWbemLocator, (LPVOID*)&pLoc);
	if (FAILED(hres))
	{
		MessageBox(_T("Failed to create IWbemLocator object"), _T("Error"), MB_OK);
		CoUninitialize();
		return;
	}

	IWbemServices* pSvc = NULL;
	hres = pLoc->ConnectServer(
		_bstr_t(L"ROOT\\CIMV2"),
		NULL,
		NULL,
		0,
		NULL,
		0,
		0,
		&pSvc
	);
	if (FAILED(hres))
	{
		MessageBox(_T("Could not connect"), _T("Error"), MB_OK);
		pLoc->Release();
		CoUninitialize();
		return;
	}

	hres = CoSetProxyBlanket(
		pSvc,
		RPC_C_AUTHN_WINNT,
		RPC_C_AUTHZ_NONE,
		NULL,
		RPC_C_AUTHN_LEVEL_CALL,
		RPC_C_IMP_LEVEL_IMPERSONATE,
		NULL,
		EOAC_NONE
	);
	if (FAILED(hres))
	{
		MessageBox(_T("Could not set proxy blanket"), _T("Error"), MB_OK);
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
		return;
	}

	IEnumWbemClassObject* pEnumerator = NULL;
	hres = pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t("SELECT ProcessID,Name,ExecutablePath,CommandLine FROM Win32_Process"),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		NULL,
		&pEnumerator);
	if (FAILED(hres))
	{
		MessageBox(_T("Query for process information failed"), _T("Error"), MB_OK);
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
		return;
	}

	IWbemClassObject* pclsObj = NULL;
	ULONG uReturn = 0;

	while (pEnumerator)
	{
		HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
		if (uReturn == 0)
		{
			break;
		}

		VARIANT ProcessID;
		VARIANT ProcessName;
		VARIANT FullPath;
		VARIANT CmdLine;

		VariantInit(&ProcessID);
		VariantInit(&ProcessName);
		VariantInit(&FullPath);
		VariantInit(&CmdLine);

		hr = pclsObj->Get(L"ProcessID", 0, &ProcessID, 0, 0);
		hr = pclsObj->Get(L"Name", 0, &ProcessName, 0, 0);
		hr = pclsObj->Get(L"ExecutablePath", 0, &FullPath, 0, 0);
		hr = pclsObj->Get(L"CommandLine", 0, &CmdLine, 0, 0);

		CString ProcessIDStr;
		CString NameStr;
		CString FullPathStr;
		CString CmdLineStr;

		ProcessIDStr.Format(_T("%u"), ProcessID.uintVal);
		NameStr = ProcessName.bstrVal;
		FullPathStr = FullPath.bstrVal;
		CmdLineStr = CmdLine.bstrVal;

		int nItem;
		nItem = ProcessList.InsertItem(0, ProcessIDStr);
		ProcessList.SetItemText(nItem, 1, NameStr);
		ProcessList.SetItemText(nItem, 2, FullPathStr);
		ProcessList.SetItemText(nItem, 3, CmdLineStr);

		VariantClear(&ProcessID);
		VariantClear(&ProcessName);
		VariantClear(&FullPath);
		VariantClear(&CmdLine);

		pclsObj->Release();

	}
	
	pSvc->Release();
	pLoc->Release();
	pEnumerator->Release();
	CoUninitialize();
	return;
}
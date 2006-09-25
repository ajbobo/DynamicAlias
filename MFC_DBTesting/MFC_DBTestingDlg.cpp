// MFC_DBTestingDlg.cpp : implementation file
//

#include "stdafx.h"
#include "odbcinst.h"
#include "MFC_DBTesting.h"
#include "MFC_DBTestingDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// CMFC_DBTestingDlg dialog



CMFC_DBTestingDlg::CMFC_DBTestingDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CMFC_DBTestingDlg::IDD, pParent)
	, SQL_Query(_T(""))
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CMFC_DBTestingDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_SQL, SQL_Query);
	DDX_Control(pDX, IDC_OPEN, OpenButton);
	DDX_Control(pDX, IDC_CLOSE, CloseButton);
	DDX_Control(pDX, IDC_EXECUTE_SQL, SQLButton);
	DDX_Control(pDX, IDC_SQL, SQLEdit);
	DDX_Control(pDX, IDC_RESULTLIST, ResultList);
}

BEGIN_MESSAGE_MAP(CMFC_DBTestingDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_OPEN, OnBnClickedOpen)
	ON_BN_CLICKED(IDC_CLOSE, OnBnClickedClose)
	ON_BN_CLICKED(IDC_EXECUTE_SQL, OnBnClickedExecuteSql)
	ON_NOTIFY(LVN_ITEMCHANGED, IDC_RESULTLIST, OnLvnItemchangedResultlist)
END_MESSAGE_MAP()


// CMFC_DBTestingDlg message handlers

BOOL CMFC_DBTestingDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
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
	DynamicAlias_LoadFunctions();
	
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CMFC_DBTestingDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else if (nID == SC_CLOSE)
	{
		DynamicAlias_UnloadFunctions();
		CDialog::OnSysCommand(nID,lParam);
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CMFC_DBTestingDlg::OnPaint() 
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
		CDialog::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CMFC_DBTestingDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CMFC_DBTestingDlg::OnBnClickedOpen()
{
	UpdateData(true);

	// Create the DSN
	CFileDialog *filedlg = new CFileDialog(true,"Excel or Access Files","*.xls; *.mdb",NULL,"Excel or Access Files (*.xls; *.mdb)| *.xls;*.mdb|Excel Files (*.xls)|*.xls|Access Files (*.mdb)|*.mdb||");
	if (filedlg->DoModal() == IDCANCEL)
	{
		UpdateData(false);
		return;
	}
	CString filename = filedlg->GetPathName();
	delete filedlg;

	if (filename.Find(".xls") != -1)
	{
		CreateAlias("aj_dsn",filename.GetBuffer(),EXCEL_ENGLISH);
	}
	else if (filename.Find(".mdb") != -1)
	{
		CreateAlias("aj_dsn",filename.GetBuffer(),ACCESS_ENGLISH);
	}

	CString connectstr = "DSN=";
	connectstr += "aj_dsn";
	connectstr += ";";
	if (!TheDB.OpenEx(connectstr,0))
	{
		MessageBox("Unable to open database file");
	}
	else
	{
		OpenButton.EnableWindow(false);
		CloseButton.EnableWindow(true);
		SQLEdit.EnableWindow(true);
		SQLButton.EnableWindow(true);
	}

	UpdateData(false);
}

void CMFC_DBTestingDlg::OnBnClickedClose()
{
	UpdateData(true);

	TheDB.Close();

	RemoveAlias("aj_dsn",EXCEL_ENGLISH);

	OpenButton.EnableWindow(true);
	CloseButton.EnableWindow(false);
	SQLEdit.EnableWindow(false);
	SQLButton.EnableWindow(false);
	
	while (ResultList.DeleteColumn(0)) ;

	UpdateData(false);
}

void CMFC_DBTestingDlg::OnBnClickedExecuteSql()
{
	UpdateData(true);

	CRecordset *theRecord = new CRecordset(&TheDB);

	theRecord->Open(CRecordset::snapshot,SQL_Query,0);

	//Count the number of records returned
	if (!theRecord->IsBOF())
		theRecord->MoveFirst();
	while (!theRecord->IsEOF())
		theRecord->MoveNext();

	// Add the columns to the List
	while (ResultList.DeleteColumn(0)) /*delete all of the columns*/ ;
	for (int col = 0; col < theRecord->GetODBCFieldCount(); col++)
	{
		CODBCFieldInfo info;
		theRecord->GetODBCFieldInfo(col,info);
		ResultList.InsertColumn(col,info.m_strName,LVCFMT_LEFT,50,0);
	}


	// Add the records to the List
	if (!theRecord->IsBOF())
		theRecord->MoveFirst();
	ResultList.DeleteAllItems();
	for (int row = 0; row < theRecord->GetRecordCount(); row++)
	{
		// Add each list item
		CString info;
		theRecord->GetFieldValue(0,info);
		ResultList.InsertItem(LVIF_TEXT|LVIF_STATE,row,info,0,0,0,col + 1);
		// Add each sub-item
		for (int col = 1; col < theRecord->GetODBCFieldCount(); col++)
		{
			theRecord->GetFieldValue(col,info);
			ResultList.SetItemText(row,col,info);
		}
		theRecord->MoveNext();
	}

	theRecord->Close();

	delete theRecord;

	UpdateData(false);
}

void CMFC_DBTestingDlg::OnLvnItemchangedResultlist(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
	// TODO: Add your control notification handler code here
	*pResult = 0;
}

// ProjectTrafficDlg.cpp : implementation file
//

#include "stdafx.h"
#include "excel.h"
#include "ProjectTraffic.h"
#include "ProjectTrafficDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/*  ExcelRandWDlg.cpp  */
//并设置全局变量
_Application g_app;
Workbooks g_books;
_Workbook g_book;
Sheets g_sheets;        //低版本Office请将这改为 WorkSheets
_Worksheet g_sheet;
Range g_range;

/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	//{{AFX_MSG(CAboutDlg)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
	//{{AFX_MSG_MAP(CAboutDlg)
		// No message handlers
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CProjectTrafficDlg dialog

CProjectTrafficDlg::CProjectTrafficDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CProjectTrafficDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CProjectTrafficDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CProjectTrafficDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CProjectTrafficDlg)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CProjectTrafficDlg, CDialog)
	//{{AFX_MSG_MAP(CProjectTrafficDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON_READ, OnButtonRead)
	ON_BN_CLICKED(IDC_BUTTON_WRITE, OnButtonWrite)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CProjectTrafficDlg message handlers

BOOL CProjectTrafficDlg::OnInitDialog()
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
	
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CProjectTrafficDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CProjectTrafficDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

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

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CProjectTrafficDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

void CProjectTrafficDlg::OnButtonRead() 
{	
    COleVariant VOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR); 
    VARIANT ret,val;
    CString str,strPath;
    long index[2];
    COleSafeArray ole;
    CFileDialog dlg(true,"配置文件",NULL,0,
        "表格文件(.xlsx)|*.xlsx|表格文件(.xls)|*.xls");//打开文件夹
    if(IDOK==dlg.DoModal())
         strPath = dlg.GetPathName();
    if(!g_app.CreateDispatch(_T("Excel.Application")))
    {
        AfxMessageBox("创建Excel服务失败");
        return;
    }
    //获取所有的工作簿
    g_books = g_app.GetWorkbooks();               
    //打开工作簿
    g_book = g_books.Open(strPath,VOptional,VOptional,        
             VOptional, VOptional,VOptional,VOptional,
             VOptional, VOptional,VOptional,VOptional,
             VOptional, VOptional,VOptional,VOptional);
    //获取所有表
    g_sheets = g_book.GetWorksheets();                
    //得到第一个表 
    g_sheet = g_sheets.GetItem(COleVariant(short(1)));
    g_range = g_sheet.GetRange(COleVariant("A1"), COleVariant("A5")); //设置操作范围
    ret = g_range.GetValue2();
    ole = &ret;
    for(int i=1;i<=5;i++)
    {
        index[0]=i;                         
        //COleSafeArray的引索 行    
        index[1]=1;                         //列
        ole.GetElement(index,&val);
        switch (val.vt)
        {
        case VT_BSTR:
            str.Format((char*)val.bstrVal);
            AfxMessageBox(str);
            break;
        case VT_R8:
            str.Format("%2f",val.dblVal);
            AfxMessageBox(str);
            break;
        default:
            break;
        }
    }
    g_range.ReleaseDispatch();
    g_sheet.ReleaseDispatch();
    g_sheets.ReleaseDispatch();
    g_book.ReleaseDispatch();
    g_books.ReleaseDispatch();
    g_app.Quit();
    g_app.ReleaseDispatch();
}

void CProjectTrafficDlg::OnButtonWrite() 
{
    CString strPath;
    COleVariant covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR); 

    CFileDialog dlg(true,"配置文件",NULL,0,
        "表格文件(.xlsx)|*.xlsx|表格文件(.xls)|*.xls");
    //打开文件
    if(IDOK==dlg.DoModal())
        strPath =dlg.GetPathName();
    if(!g_app.CreateDispatch(_T("Excel.Application")))
    {
        AfxMessageBox("创建Excel服务失败");
        return;
    }
    //获取所有的工作簿
    g_books = g_app.GetWorkbooks();   
    //用来锁定对应的工作簿           
    g_books.AttachDispatch(g_app.GetWorkbooks(),true);       
    g_book = g_books.Open( strPath,covOptional,covOptional,
          covOptional,covOptional,covOptional,covOptional,
          covOptional,covOptional,covOptional,covOptional,
          covOptional,covOptional,covOptional,covOptional);

    //得到Worksheets
    g_sheets.AttachDispatch(g_book.GetWorksheets(),true);
    g_sheet=g_sheets.GetItem(COleVariant((short)1));

    //得到全部Cells，此时

    //设置5行第一列的单元的值
    g_range=g_sheet.GetRange(COleVariant(_T("A1")),COleVariant(_T("A1")));
    g_range.SetValue2(COleVariant(_T("1")));
    g_range=g_sheet.GetRange(COleVariant(_T("A2")),COleVariant(_T("A2")));
    g_range.SetValue2(COleVariant(_T("2")));
    g_range=g_sheet.GetRange(COleVariant(_T("A3")),COleVariant(_T("A3")));
    g_range.SetValue2(COleVariant(_T("3")));
    g_range=g_sheet.GetRange(COleVariant(_T("A4")),COleVariant(_T("A4")));
    g_range.SetValue2(COleVariant(_T("4")));
    g_range=g_sheet.GetRange(COleVariant(_T("A5")),COleVariant(_T("A5")));
    g_range.SetValue2(COleVariant(_T("6")));
     g_book.Save(); //保存
        g_range.ReleaseDispatch(); //退出
    g_sheet.ReleaseDispatch();
    g_sheets.ReleaseDispatch();
    g_book.ReleaseDispatch();
    g_books.ReleaseDispatch();
    g_app.Quit();
    g_app.ReleaseDispatch();
}

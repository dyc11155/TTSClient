// CTTTSTestDlg.cpp : implementation file
//

#include "stdafx.h"
#include "CTTTSTest.h"
#include "CTTTSTestDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

typedef int (__cdecl *String2AudioFile)(char*, char*); // __cdecl | __stdcall | __fastcall 
typedef int (__cdecl *STTSinit)();  

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
// CCTTTSTestDlg dialog

CCTTTSTestDlg::CCTTTSTestDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CCTTTSTestDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CCTTTSTestDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CCTTTSTestDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CCTTTSTestDlg)
	// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
	DDX_Control(pDX, IDC_EDIT_LOG, m_editLog);
}

BEGIN_MESSAGE_MAP(CCTTTSTestDlg, CDialog)
	//{{AFX_MSG_MAP(CCTTTSTestDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BTN_TRANSE, OnBtnTranse)
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_BUTTON1, &CCTTTSTestDlg::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON2, &CCTTTSTestDlg::OnBnClickedReadExcel)
	ON_BN_CLICKED(IDC_BTN_SYNTHETICVOICE, &CCTTTSTestDlg::OnBnClickedBtnSyntheticvoice)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CCTTTSTestDlg message handlers

BOOL CCTTTSTestDlg::OnInitDialog()
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
	GetDlgItem(IDC_EDIT_DLL)->SetWindowText("TTSClient.dll");
	GetDlgItem(IDC_EDIT_CONTENT)->SetWindowText("TTS测试测试");
	GetDlgItem(IDC_EDIT1)->SetWindowText("A");
	GetDlgItem(IDC_EDIT2)->SetWindowText("B");
	GetDlgItem(IDC_EDIT3)->SetWindowText("1");
	GetDlgItem(IDC_EDIT4)->SetWindowText("2");
	GetDlgItem(IDC_EDIT5)->SetWindowText("1");
	GetDlgItem(IDC_EDIT6)->SetWindowText("2");
	GetDlgItem(IDC_EDIT8)->SetWindowText("F:\\bjzx.xlsx");
	GetDlgItem(IDC_EDIT9)->SetWindowText("1");

	//CString sDllName = "TTSClient.dll";
	////GetDlgItem(IDC_EDIT_DLL)->GetWindowText(sDllName);

	//HMODULE hSApi = NULL;
	//hSApi = LoadLibrary(_T(sDllName));
	//if(hSApi == NULL) 
	//{
	//	AfxMessageBox("加载dll出错");
	//	return;
	//}
	
	
	
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CCTTTSTestDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

void CCTTTSTestDlg::OnPaint() 
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
HCURSOR CCTTTSTestDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

void CCTTTSTestDlg::OnBtnTranse() 
{
	// TODO: Add your control notification handler code here
	//获取dll名字
	CString sDllName = "";
	GetDlgItem(IDC_EDIT_DLL)->GetWindowText(sDllName);

	HMODULE hSApi = NULL;
	hSApi = LoadLibrary(_T(sDllName));
	if(hSApi == NULL) 
	{
		AfxMessageBox("加载dll出错");
		return;
	}
	//语音总数
	char VoiceNumString[50] = _T("");
	int VoiceNum = 0;
	CString total = _T("");

	GetPrivateProfileString("TTS", "VoiceTotal", "", VoiceNumString, 49, "c:\\tts.ini");

	VoiceNum = _ttoi(VoiceNumString);
	total.Format("语音总数为%d",VoiceNum);

	AfxMessageBox(total);
	
	//合成语音
	int i;
	char VoiceContent[2000];
	char VoiceName[50];
	CString sContent = _T(""), sRet = _T("");
	CString VoiceID = _T("");
	CString sName;
	
	for(i=1;i<VoiceNum;i++)
	{
		VoiceID.Format("%d",i);
		GetPrivateProfileString("TTSContent", VoiceID, "", VoiceContent, 1999, "c:\\tts.ini");
		sContent = VoiceContent;
		GetPrivateProfileString("TTSName", VoiceID, "", VoiceName, 49, "c:\\tts.ini");
		sName = VoiceName;
		

		if(sDllName.CompareNoCase("TTSClient.dll")==0)
		{
			String2AudioFile ProcS2F = (String2AudioFile)GetProcAddress(hSApi, "String2AudioFile"); 
			if(NULL != ProcS2F)
			{
				
				//GetDlgItem(IDC_EDIT_CONTENT)->GetWindowText(sContent);
				int nRet = (ProcS2F) ((LPSTR)(LPCSTR)sContent,(LPSTR)(LPCSTR)sName); 
				sRet.Format("调用返回%d", nRet);
				GetDlgItem(IDC_EDIT_IP)->SetWindowText(sRet);
			} 
		} else if(sDllName.CompareNoCase("TTStranslate.dll")==0) {
		} 

	}

	//设置定义
	/*if(sDllName.CompareNoCase("TTSClient.dll")==0)
	{
		String2AudioFile ProcS2F = (String2AudioFile)GetProcAddress(hSApi, "String2AudioFile"); 
		if(NULL != ProcS2F)
		{
			CString sContent = "", sRet = "";
			GetDlgItem(IDC_EDIT_CONTENT)->GetWindowText(sContent);
			int nRet = (ProcS2F) ((LPSTR)(LPCSTR)sContent, "C:\\test.wav"); 
			sRet.Format("调用返回%d", nRet);
			GetDlgItem(IDC_EDIT_IP)->SetWindowText(sRet);
		} 
	} else if(sDllName.CompareNoCase("TTStranslate.dll")==0) {
	} 
	*/
	// Free the DLL module.   	
	FreeLibrary(hSApi);	
}


void CCTTTSTestDlg::OnBnClickedButton1()
{
	// TODO: 在此添加控件通知处理程序代码
	/********************excel操作************************/
	//导出
	CoInitialize(NULL);

    COleVariant covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR); 
    if (!app.CreateDispatch(_T("Excel.Application")))
    {   
        this->MessageBox(_T("无法创建Excel应用！")); 
        return;  
    }   
    books = app.get_Workbooks(); 
    //打开Excel，其中pathname为Excel表的路径名  
    
    book = books.Add(covOptional);  
    sheets = book.get_Worksheets();  
    sheet = sheets.get_Item(COleVariant((short)1));  //获得坐标为（A，1）和（B，1）的两个单元格 
    range = sheet.get_Range(COleVariant(_T("A1")), COleVariant(_T("B1")));  //设置单元格类容为Hello Exce
    range.put_Value2(COleVariant(_T("Hello Excel")));  //选择整列，并设置宽度为自适应 
    cols = range.get_EntireColumn();  
    cols.AutoFit();  
    //设置字体为粗体 
    font = range.get_Font();  
    font.put_Bold(COleVariant((short)TRUE)); 
    //获得坐标为（C，2）单元格 
    range = sheet.get_Range(COleVariant(_T("C2")), COleVariant(_T("C2"))); 
    //设置公式“=RAND()*100000”
    range.put_Formula(COleVariant(_T("=2*100000"))); 
    //设置数字格式为货币型  
    range.put_NumberFormat(COleVariant(_T("$0.00"))); 
    //选择整列，并设置宽度为自适应  
    cols = range.get_EntireColumn(); 
    cols.AutoFit(); 
    //显示Excel表
    app.put_Visible(TRUE);   
    app.put_UserControl(TRUE);
}


void CCTTTSTestDlg::OnBnClickedReadExcel()
{
	this->PrintLog(_T("===============开始读取EXCEL============"));
	// TODO: 在此添加控件通知处理程序代码
	int fileColumnStartNum = 0;
	int fileColumnEndNum = 0;
	int contentColumnStartNum = 0;
	int contentColumnEndNum = 0;

	
	CString sLog = _T("");
	CString sFileColumn = _T("");
	CString sContentColumn = _T("");
	CString sFileColumnStartNum = _T("");
	CString sFileColumnEndNum = _T("");
	CString sContentColumnStartNum = _T("");
	CString sContentColumnEndNum = _T("");
	CString sExcelPath = _T("");
	CString sExcelSheetNum = _T("");
	GetDlgItem(IDC_EDIT1)->GetWindowTextA(sFileColumn);
	GetDlgItem(IDC_EDIT2)->GetWindowTextA(sContentColumn);
	GetDlgItem(IDC_EDIT3)->GetWindowTextA(sFileColumnStartNum);
	GetDlgItem(IDC_EDIT4)->GetWindowTextA(sFileColumnEndNum);
	GetDlgItem(IDC_EDIT5)->GetWindowTextA(sContentColumnStartNum);
	GetDlgItem(IDC_EDIT6)->GetWindowTextA(sContentColumnEndNum);
	GetDlgItem(IDC_EDIT8)->GetWindowTextA(sExcelPath);
	GetDlgItem(IDC_EDIT9)->GetWindowTextA(sExcelSheetNum);

	if(sFileColumn == _T("")||sContentColumn == _T("")||sFileColumnStartNum == _T("")||sFileColumnEndNum == _T("")||sContentColumnStartNum == _T("")||sContentColumnEndNum == _T(""))
	{
		this->MessageBox(_T("请输入信息"));
		return;
	}

	fileColumnStartNum = atoi(sFileColumnStartNum);
	fileColumnEndNum = atoi(sFileColumnEndNum);
	contentColumnStartNum = atoi(sFileColumnStartNum);
	contentColumnEndNum = atoi(sFileColumnEndNum);
	
	this->PrintLog(_T("读取数据完成"));

	//导入
    COleVariant covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR); 
    if (!app.CreateDispatch(_T("Excel.Application")))
    {   
        this->MessageBox(_T("无法创建Excel应用！")); 
        return;  
    }
    books = app.get_Workbooks(); 
    //打开Excel，其中pathname为Excel表的路径名  
    lpDisp = books.Open(sExcelPath,covOptional ,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional);
    book.AttachDispatch(lpDisp); 
    sheets = book.get_Worksheets(); 
    sheet = sheets.get_Item(COleVariant((short)atoi(sExcelSheetNum))); 

	this->PrintLog(_T("打开文件完成"));
    
	//获得坐标为（A，1）的单元格 ,获取并写入-----文件名=文件位置
	for(int i = fileColumnStartNum;i <= fileColumnEndNum;i++)
	{
		sLog.Format("<<循环开始i=%d，开始%d，结束%d>>",i,fileColumnStartNum,fileColumnEndNum);
		this->PrintLog(sLog);
		CString sFileName = _T("");
		CString sFileContent = _T("");
		CString sFileNameResult = _T("");
		CString sFileContentResult = _T("");
		//获取文件名
		sFileName.Format("%s%d",sFileColumn,i);
		sLog.Format("读取单元格为：%s",sFileName);
		this->PrintLog(sLog);
		range = sheet.get_Range(COleVariant(sFileName) ,COleVariant(sFileName));  
		//获得单元格的内容 
		COleVariant rValue;
		rValue =   COleVariant(range.get_Value2()); 
		//转换成宽字符  
		rValue.ChangeType(VT_BSTR); 
		//转换格式，并输出 
		sFileNameResult = CString(rValue.bstrVal);
		sLog.Format("文件名为：%s",sFileNameResult);
		this->PrintLog(sLog);
		
		//获取文件内容
		sFileContent.Format("%s%d",sContentColumn,i);
		sLog.Format("读取单元格为：%s",sFileContent);
		this->PrintLog(sLog);
		range = sheet.get_Range(COleVariant(sFileContent) ,COleVariant(sFileContent));  
		//获得单元格的内容 
		//COleVariant rValue;
		rValue =   COleVariant(range.get_Value2()); 
		//转换成宽字符  
		rValue.ChangeType(VT_BSTR); 
		//转换格式，并输出 
		sFileContentResult = CString(rValue.bstrVal);
		sLog.Format("文件名为：%s",sFileContentResult);
		this->PrintLog(sLog);


		//this->MessageBox(CString(rValue.bstrVal));
		/*CString value;
		value = CString(rValue.bstrVal);*/
		::WritePrivateProfileStringA("VALUE",sFileNameResult,sFileContentResult,"D:\\tts.ini");
		sLog.Format("<<循环结束i=%d，开始%d，结束%d>>",i,fileColumnStartNum,fileColumnEndNum);
		this->PrintLog(sLog);
		
	}

	this->PrintLog(_T("=============全部完成============"));

    book.put_Saved(TRUE); 
	book.Close(covOptional,covOptional,covOptional);
	books.Close();	
	app.Quit();
	range.ReleaseDispatch();
	sheet.ReleaseDispatch();
	sheets.ReleaseDispatch();
	book.ReleaseDispatch();
	books.ReleaseDispatch();
	//app.DetachDispatch();
	app.ReleaseDispatch();
	CoInitialize(NULL);
	
}


void CCTTTSTestDlg::OnBnClickedBtnSyntheticvoice()
{
	// TODO: 在此添加控件通知处理程序代码
	CString sDllName = "";
	GetDlgItem(IDC_EDIT_DLL)->GetWindowText(sDllName);

	HMODULE hSApi = NULL;
	hSApi = LoadLibrary(_T(sDllName));
	if(hSApi == NULL) 
	{
		AfxMessageBox("加载dll出错");
		return;
	}


	this->PrintLog(_T("===============开始读取EXCEL============"));
	
	int fileColumnStartNum = 0;
	int fileColumnEndNum = 0;
	int contentColumnStartNum = 0;
	int contentColumnEndNum = 0;

	
	CString sLog = _T("");
	CString sFileColumn = _T("");
	CString sContentColumn = _T("");
	CString sFileColumnStartNum = _T("");
	CString sFileColumnEndNum = _T("");
	CString sContentColumnStartNum = _T("");
	CString sContentColumnEndNum = _T("");
	CString sExcelPath = _T("");
	CString sExcelSheetNum = _T("");
	GetDlgItem(IDC_EDIT1)->GetWindowTextA(sFileColumn);
	GetDlgItem(IDC_EDIT2)->GetWindowTextA(sContentColumn);
	GetDlgItem(IDC_EDIT3)->GetWindowTextA(sFileColumnStartNum);
	GetDlgItem(IDC_EDIT4)->GetWindowTextA(sFileColumnEndNum);
	GetDlgItem(IDC_EDIT5)->GetWindowTextA(sContentColumnStartNum);
	GetDlgItem(IDC_EDIT6)->GetWindowTextA(sContentColumnEndNum);
	GetDlgItem(IDC_EDIT8)->GetWindowTextA(sExcelPath);
	GetDlgItem(IDC_EDIT9)->GetWindowTextA(sExcelSheetNum);

	if(sFileColumn == _T("")||sContentColumn == _T("")||sFileColumnStartNum == _T("")||sFileColumnEndNum == _T("")||sContentColumnStartNum == _T("")||sContentColumnEndNum == _T(""))
	{
		this->MessageBox(_T("请输入信息"));
		return;
	}

	fileColumnStartNum = atoi(sFileColumnStartNum);
	fileColumnEndNum = atoi(sFileColumnEndNum);
	contentColumnStartNum = atoi(sFileColumnStartNum);
	contentColumnEndNum = atoi(sFileColumnEndNum);
	
	this->PrintLog(_T("读取数据完成"));

	//导入
    COleVariant covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR); 
    if (!app.CreateDispatch(_T("Excel.Application")))
    {   
        this->MessageBox(_T("无法创建Excel应用！")); 
        return;  
    }
    books = app.get_Workbooks(); 
    //打开Excel，其中pathname为Excel表的路径名  
    lpDisp = books.Open(sExcelPath,covOptional ,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional);
    book.AttachDispatch(lpDisp); 
    sheets = book.get_Worksheets(); 
    sheet = sheets.get_Item(COleVariant((short)atoi(sExcelSheetNum))); 

	this->PrintLog(_T("打开文件完成"));
    
	//获得坐标为（A，1）的单元格 ,获取并写入-----文件名=文件位置
	for(int i = fileColumnStartNum;i <= fileColumnEndNum;i++)
	{
		sLog.Format("<<循环开始i=%d，开始%d，结束%d>>",i,fileColumnStartNum,fileColumnEndNum);
		this->PrintLog(sLog);
		CString sFileName = _T("");
		CString sFileContent = _T("");
		CString sFileNameResult = _T("");
		CString sFileContentResult = _T("");
		//获取文件名
		sFileName.Format("%s%d",sFileColumn,i);
		sLog.Format("读取单元格为：%s",sFileName);
		this->PrintLog(sLog);
		range = sheet.get_Range(COleVariant(sFileName) ,COleVariant(sFileName));  
		//获得单元格的内容 
		COleVariant rValue;
		rValue =   COleVariant(range.get_Value2()); 
		//转换成宽字符  
		rValue.ChangeType(VT_BSTR); 
		//转换格式，并输出 
		sFileNameResult = CString(rValue.bstrVal);
		sLog.Format("文件名为：%s",sFileNameResult);
		this->PrintLog(sLog);
		
		//获取文件内容
		sFileContent.Format("%s%d",sContentColumn,i);
		sLog.Format("读取单元格为：%s",sFileContent);
		this->PrintLog(sLog);
		range = sheet.get_Range(COleVariant(sFileContent) ,COleVariant(sFileContent));  
		//获得单元格的内容 
		//COleVariant rValue;
		rValue =   COleVariant(range.get_Value2()); 
		//转换成宽字符  
		rValue.ChangeType(VT_BSTR); 
		//转换格式，并输出 
		sFileContentResult = CString(rValue.bstrVal);
		sLog.Format("文件名为：%s",sFileContentResult);
		this->PrintLog(sLog);

		sFileNameResult = _T("F:\\bjzx\\") + sFileNameResult;

		//this->MessageBox(CString(rValue.bstrVal));
		/*CString value;
		value = CString(rValue.bstrVal);*/
		


		this->PrintLog(_T("===========开始合成语音。。。。========"));
		if(sDllName.CompareNoCase("TTSClient.dll")==0)
		{
			String2AudioFile ProcS2F = (String2AudioFile)GetProcAddress(hSApi, "String2AudioFile"); 
			if(NULL != ProcS2F)
			{

				//GetDlgItem(IDC_EDIT_CONTENT)->GetWindowText(sContent);
				int nRet = (ProcS2F) ((LPSTR)(LPCSTR)sFileContentResult,(LPSTR)(LPCSTR)sFileNameResult); 
				//sRet.Format("调用返回%d", nRet);
				//GetDlgItem(IDC_EDIT_IP)->SetWindowText(sRet);
			} 
		} else if(sDllName.CompareNoCase("TTStranslate.dll")==0) {
		} 

		this->PrintLog(_T("===========合成结束========="));

		::WritePrivateProfileStringA("VALUE",sFileNameResult,sFileContentResult,"D:\\tts.ini");
		sLog.Format("<<循环结束i=%d，开始%d，结束%d>>",i,fileColumnStartNum,fileColumnEndNum);
		this->PrintLog(sLog);

		
	}

	this->PrintLog(_T("=============全部完成============"));

	
    book.put_Saved(TRUE);
	book.Close(covOptional,covOptional,covOptional);
	books.Close();	
    app.Quit();
	range.ReleaseDispatch();
	sheet.ReleaseDispatch();
	sheets.ReleaseDispatch();
	book.ReleaseDispatch();
	books.ReleaseDispatch();
	//app.DetachDispatch();
	app.ReleaseDispatch();
	CoInitialize(NULL);
	FreeLibrary(hSApi);	
}

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
	GetDlgItem(IDC_EDIT_CONTENT)->SetWindowText("TTS���Բ���");
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
	//	AfxMessageBox("����dll����");
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
	//��ȡdll����
	CString sDllName = "";
	GetDlgItem(IDC_EDIT_DLL)->GetWindowText(sDllName);

	HMODULE hSApi = NULL;
	hSApi = LoadLibrary(_T(sDllName));
	if(hSApi == NULL) 
	{
		AfxMessageBox("����dll����");
		return;
	}
	//��������
	char VoiceNumString[50] = _T("");
	int VoiceNum = 0;
	CString total = _T("");

	GetPrivateProfileString("TTS", "VoiceTotal", "", VoiceNumString, 49, "c:\\tts.ini");

	VoiceNum = _ttoi(VoiceNumString);
	total.Format("��������Ϊ%d",VoiceNum);

	AfxMessageBox(total);
	
	//�ϳ�����
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
				sRet.Format("���÷���%d", nRet);
				GetDlgItem(IDC_EDIT_IP)->SetWindowText(sRet);
			} 
		} else if(sDllName.CompareNoCase("TTStranslate.dll")==0) {
		} 

	}

	//���ö���
	/*if(sDllName.CompareNoCase("TTSClient.dll")==0)
	{
		String2AudioFile ProcS2F = (String2AudioFile)GetProcAddress(hSApi, "String2AudioFile"); 
		if(NULL != ProcS2F)
		{
			CString sContent = "", sRet = "";
			GetDlgItem(IDC_EDIT_CONTENT)->GetWindowText(sContent);
			int nRet = (ProcS2F) ((LPSTR)(LPCSTR)sContent, "C:\\test.wav"); 
			sRet.Format("���÷���%d", nRet);
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
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	/********************excel����************************/
	//����
	CoInitialize(NULL);

    COleVariant covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR); 
    if (!app.CreateDispatch(_T("Excel.Application")))
    {   
        this->MessageBox(_T("�޷�����ExcelӦ�ã�")); 
        return;  
    }   
    books = app.get_Workbooks(); 
    //��Excel������pathnameΪExcel���·����  
    
    book = books.Add(covOptional);  
    sheets = book.get_Worksheets();  
    sheet = sheets.get_Item(COleVariant((short)1));  //�������Ϊ��A��1���ͣ�B��1����������Ԫ�� 
    range = sheet.get_Range(COleVariant(_T("A1")), COleVariant(_T("B1")));  //���õ�Ԫ������ΪHello Exce
    range.put_Value2(COleVariant(_T("Hello Excel")));  //ѡ�����У������ÿ��Ϊ����Ӧ 
    cols = range.get_EntireColumn();  
    cols.AutoFit();  
    //��������Ϊ���� 
    font = range.get_Font();  
    font.put_Bold(COleVariant((short)TRUE)); 
    //�������Ϊ��C��2����Ԫ�� 
    range = sheet.get_Range(COleVariant(_T("C2")), COleVariant(_T("C2"))); 
    //���ù�ʽ��=RAND()*100000��
    range.put_Formula(COleVariant(_T("=2*100000"))); 
    //�������ָ�ʽΪ������  
    range.put_NumberFormat(COleVariant(_T("$0.00"))); 
    //ѡ�����У������ÿ��Ϊ����Ӧ  
    cols = range.get_EntireColumn(); 
    cols.AutoFit(); 
    //��ʾExcel��
    app.put_Visible(TRUE);   
    app.put_UserControl(TRUE);
}


void CCTTTSTestDlg::OnBnClickedReadExcel()
{
	this->PrintLog(_T("===============��ʼ��ȡEXCEL============"));
	// TODO: �ڴ���ӿؼ�֪ͨ����������
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
		this->MessageBox(_T("��������Ϣ"));
		return;
	}

	fileColumnStartNum = atoi(sFileColumnStartNum);
	fileColumnEndNum = atoi(sFileColumnEndNum);
	contentColumnStartNum = atoi(sFileColumnStartNum);
	contentColumnEndNum = atoi(sFileColumnEndNum);
	
	this->PrintLog(_T("��ȡ�������"));

	//����
    COleVariant covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR); 
    if (!app.CreateDispatch(_T("Excel.Application")))
    {   
        this->MessageBox(_T("�޷�����ExcelӦ�ã�")); 
        return;  
    }
    books = app.get_Workbooks(); 
    //��Excel������pathnameΪExcel���·����  
    lpDisp = books.Open(sExcelPath,covOptional ,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional);
    book.AttachDispatch(lpDisp); 
    sheets = book.get_Worksheets(); 
    sheet = sheets.get_Item(COleVariant((short)atoi(sExcelSheetNum))); 

	this->PrintLog(_T("���ļ����"));
    
	//�������Ϊ��A��1���ĵ�Ԫ�� ,��ȡ��д��-----�ļ���=�ļ�λ��
	for(int i = fileColumnStartNum;i <= fileColumnEndNum;i++)
	{
		sLog.Format("<<ѭ����ʼi=%d����ʼ%d������%d>>",i,fileColumnStartNum,fileColumnEndNum);
		this->PrintLog(sLog);
		CString sFileName = _T("");
		CString sFileContent = _T("");
		CString sFileNameResult = _T("");
		CString sFileContentResult = _T("");
		//��ȡ�ļ���
		sFileName.Format("%s%d",sFileColumn,i);
		sLog.Format("��ȡ��Ԫ��Ϊ��%s",sFileName);
		this->PrintLog(sLog);
		range = sheet.get_Range(COleVariant(sFileName) ,COleVariant(sFileName));  
		//��õ�Ԫ������� 
		COleVariant rValue;
		rValue =   COleVariant(range.get_Value2()); 
		//ת���ɿ��ַ�  
		rValue.ChangeType(VT_BSTR); 
		//ת����ʽ������� 
		sFileNameResult = CString(rValue.bstrVal);
		sLog.Format("�ļ���Ϊ��%s",sFileNameResult);
		this->PrintLog(sLog);
		
		//��ȡ�ļ�����
		sFileContent.Format("%s%d",sContentColumn,i);
		sLog.Format("��ȡ��Ԫ��Ϊ��%s",sFileContent);
		this->PrintLog(sLog);
		range = sheet.get_Range(COleVariant(sFileContent) ,COleVariant(sFileContent));  
		//��õ�Ԫ������� 
		//COleVariant rValue;
		rValue =   COleVariant(range.get_Value2()); 
		//ת���ɿ��ַ�  
		rValue.ChangeType(VT_BSTR); 
		//ת����ʽ������� 
		sFileContentResult = CString(rValue.bstrVal);
		sLog.Format("�ļ���Ϊ��%s",sFileContentResult);
		this->PrintLog(sLog);


		//this->MessageBox(CString(rValue.bstrVal));
		/*CString value;
		value = CString(rValue.bstrVal);*/
		::WritePrivateProfileStringA("VALUE",sFileNameResult,sFileContentResult,"D:\\tts.ini");
		sLog.Format("<<ѭ������i=%d����ʼ%d������%d>>",i,fileColumnStartNum,fileColumnEndNum);
		this->PrintLog(sLog);
		
	}

	this->PrintLog(_T("=============ȫ�����============"));

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
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	CString sDllName = "";
	GetDlgItem(IDC_EDIT_DLL)->GetWindowText(sDllName);

	HMODULE hSApi = NULL;
	hSApi = LoadLibrary(_T(sDllName));
	if(hSApi == NULL) 
	{
		AfxMessageBox("����dll����");
		return;
	}


	this->PrintLog(_T("===============��ʼ��ȡEXCEL============"));
	
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
		this->MessageBox(_T("��������Ϣ"));
		return;
	}

	fileColumnStartNum = atoi(sFileColumnStartNum);
	fileColumnEndNum = atoi(sFileColumnEndNum);
	contentColumnStartNum = atoi(sFileColumnStartNum);
	contentColumnEndNum = atoi(sFileColumnEndNum);
	
	this->PrintLog(_T("��ȡ�������"));

	//����
    COleVariant covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR); 
    if (!app.CreateDispatch(_T("Excel.Application")))
    {   
        this->MessageBox(_T("�޷�����ExcelӦ�ã�")); 
        return;  
    }
    books = app.get_Workbooks(); 
    //��Excel������pathnameΪExcel���·����  
    lpDisp = books.Open(sExcelPath,covOptional ,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional);
    book.AttachDispatch(lpDisp); 
    sheets = book.get_Worksheets(); 
    sheet = sheets.get_Item(COleVariant((short)atoi(sExcelSheetNum))); 

	this->PrintLog(_T("���ļ����"));
    
	//�������Ϊ��A��1���ĵ�Ԫ�� ,��ȡ��д��-----�ļ���=�ļ�λ��
	for(int i = fileColumnStartNum;i <= fileColumnEndNum;i++)
	{
		sLog.Format("<<ѭ����ʼi=%d����ʼ%d������%d>>",i,fileColumnStartNum,fileColumnEndNum);
		this->PrintLog(sLog);
		CString sFileName = _T("");
		CString sFileContent = _T("");
		CString sFileNameResult = _T("");
		CString sFileContentResult = _T("");
		//��ȡ�ļ���
		sFileName.Format("%s%d",sFileColumn,i);
		sLog.Format("��ȡ��Ԫ��Ϊ��%s",sFileName);
		this->PrintLog(sLog);
		range = sheet.get_Range(COleVariant(sFileName) ,COleVariant(sFileName));  
		//��õ�Ԫ������� 
		COleVariant rValue;
		rValue =   COleVariant(range.get_Value2()); 
		//ת���ɿ��ַ�  
		rValue.ChangeType(VT_BSTR); 
		//ת����ʽ������� 
		sFileNameResult = CString(rValue.bstrVal);
		sLog.Format("�ļ���Ϊ��%s",sFileNameResult);
		this->PrintLog(sLog);
		
		//��ȡ�ļ�����
		sFileContent.Format("%s%d",sContentColumn,i);
		sLog.Format("��ȡ��Ԫ��Ϊ��%s",sFileContent);
		this->PrintLog(sLog);
		range = sheet.get_Range(COleVariant(sFileContent) ,COleVariant(sFileContent));  
		//��õ�Ԫ������� 
		//COleVariant rValue;
		rValue =   COleVariant(range.get_Value2()); 
		//ת���ɿ��ַ�  
		rValue.ChangeType(VT_BSTR); 
		//ת����ʽ������� 
		sFileContentResult = CString(rValue.bstrVal);
		sLog.Format("�ļ���Ϊ��%s",sFileContentResult);
		this->PrintLog(sLog);

		sFileNameResult = _T("F:\\bjzx\\") + sFileNameResult;

		//this->MessageBox(CString(rValue.bstrVal));
		/*CString value;
		value = CString(rValue.bstrVal);*/
		


		this->PrintLog(_T("===========��ʼ�ϳ�������������========"));
		if(sDllName.CompareNoCase("TTSClient.dll")==0)
		{
			String2AudioFile ProcS2F = (String2AudioFile)GetProcAddress(hSApi, "String2AudioFile"); 
			if(NULL != ProcS2F)
			{

				//GetDlgItem(IDC_EDIT_CONTENT)->GetWindowText(sContent);
				int nRet = (ProcS2F) ((LPSTR)(LPCSTR)sFileContentResult,(LPSTR)(LPCSTR)sFileNameResult); 
				//sRet.Format("���÷���%d", nRet);
				//GetDlgItem(IDC_EDIT_IP)->SetWindowText(sRet);
			} 
		} else if(sDllName.CompareNoCase("TTStranslate.dll")==0) {
		} 

		this->PrintLog(_T("===========�ϳɽ���========="));

		::WritePrivateProfileStringA("VALUE",sFileNameResult,sFileContentResult,"D:\\tts.ini");
		sLog.Format("<<ѭ������i=%d����ʼ%d������%d>>",i,fileColumnStartNum,fileColumnEndNum);
		this->PrintLog(sLog);

		
	}

	this->PrintLog(_T("=============ȫ�����============"));

	
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

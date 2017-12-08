// CTTTSTestDlg.h : header file
//

#include "CApplication.h"
#include "CWorkbook.h"
#include "CWorkbooks.h"
#include "CWorksheet.h"
#include "CWorksheets.h"
#include "CRange.h"
#include "CFont0.h"
#include "afxwin.h"

#if !defined(AFX_CTTTSTESTDLG_H__BC4C4037_D216_4F0C_AE77_2940C84E5C33__INCLUDED_)
#define AFX_CTTTSTESTDLG_H__BC4C4037_D216_4F0C_AE77_2940C84E5C33__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

/////////////////////////////////////////////////////////////////////////////
// CCTTTSTestDlg dialog

class CCTTTSTestDlg : public CDialog
{
// Construction
public:
	CCTTTSTestDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	//{{AFX_DATA(CCTTTSTestDlg)
	enum { IDD = IDD_CTTTSTEST_DIALOG };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CCTTTSTestDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

	//定义接口类变量 
	CApplication app;
	CWorkbook book;
	CWorkbooks books;
	CWorksheet sheet;
	CWorksheets sheets;
	CRange range;
	CFont0 font;
	CRange cols; 
	LPDISPATCH lpDisp;
	

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CCTTTSTestDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnBtnTranse();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedReadExcel();


public:
	CEdit m_editLog;

	void PrintLog(CString sLog)
	{
		CString oldLog;

		int LineCount = m_editLog.GetLineCount();


		//if(LineCount > 50)
		//{
		//	m_editLog.SetWindowText("");
		//}


		m_editLog.GetWindowText(oldLog);
		CString tmp = oldLog + "\r\n" + sLog;

		m_editLog.SetWindowText(tmp);

		int cnt = m_editLog.GetLineCount();
		m_editLog.LineScroll(cnt, 0);
	}
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_CTTTSTESTDLG_H__BC4C4037_D216_4F0C_AE77_2940C84E5C33__INCLUDED_)

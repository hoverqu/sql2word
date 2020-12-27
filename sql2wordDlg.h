// sql2wordDlg.h : header file
//

#if !defined(AFX_SQL2WORDDLG_H__411584DA_58D5_4DD8_8EF5_18DD4F6841F4__INCLUDED_)
#define AFX_SQL2WORDDLG_H__411584DA_58D5_4DD8_8EF5_18DD4F6841F4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

/////////////////////////////////////////////////////////////////////////////
// CSql2wordDlg dialog

class CSql2wordDlg : public CDialog
{
// Construction
public:
	CSql2wordDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	//{{AFX_DATA(CSql2wordDlg)
	enum { IDD = IDD_SQL2WORD_DIALOG };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CSql2wordDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CSql2wordDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnButton1();
	afx_msg void OnButton2();
	afx_msg void OnButton3();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_SQL2WORDDLG_H__411584DA_58D5_4DD8_8EF5_18DD4F6841F4__INCLUDED_)

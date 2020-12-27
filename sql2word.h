// sql2word.h : main header file for the SQL2WORD application
//

#if !defined(AFX_SQL2WORD_H__91FF1F3A_7900_48B9_9463_1B0988C03368__INCLUDED_)
#define AFX_SQL2WORD_H__91FF1F3A_7900_48B9_9463_1B0988C03368__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols
#include "Sql2word_i.h"

/////////////////////////////////////////////////////////////////////////////
// CSql2wordApp:
// See sql2word.cpp for the implementation of this class
//

class CSql2wordApp : public CWinApp
{
public:
	CSql2wordApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CSql2wordApp)
	public:
	virtual BOOL InitInstance();
		virtual int ExitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CSql2wordApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
private:
	BOOL m_bATLInited;
private:
	BOOL InitATL();
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_SQL2WORD_H__91FF1F3A_7900_48B9_9463_1B0988C03368__INCLUDED_)

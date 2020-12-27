// sql2wordDlg.cpp : implementation file
//

#include "stdafx.h"
#include "sql2word.h"
#include "sql2wordDlg.h"
#include "CommDlg.h"

#include"comdef.h"
#include"atlbase.h"
//#include"msword.h"
#include"WordHandle.h"
#include "sqlstruct.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

struct SqlDataBase * Paraser(CString SQLstr,struct SqlDataBase *);
int GetTableCount(CString SQLstr);
int IsTextUTF8(char* str,ULONGLONG length);


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
// CSql2wordDlg dialog

CSql2wordDlg::CSql2wordDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CSql2wordDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CSql2wordDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CSql2wordDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CSql2wordDlg)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CSql2wordDlg, CDialog)
	//{{AFX_MSG_MAP(CSql2wordDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON1, OnButton1)
	//ON_BN_CLICKED(IDOK, OnButton2)
	ON_BN_CLICKED(IDC_BUTTON3, OnButton3)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CSql2wordDlg message handlers

BOOL CSql2wordDlg::OnInitDialog()
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

void CSql2wordDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

void CSql2wordDlg::OnPaint() 
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
HCURSOR CSql2wordDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

void CSql2wordDlg::OnButton1() 
{
//���ļ��Ի���װ��sql�ļ�
int TableCount=0,i=0;
int c;
LPWSTR lpWideCharStr;
CString str1,str2;
CFileDialog dlg (TRUE, NULL,NULL, OFN_HIDEREADONLY|OFN_OVERWRITEPROMPT|OFN_ALLOWMULTISELECT,"All Files (*.sql)|*.sql| |", AfxGetMainWnd());
if (dlg.DoModal () == IDOK)
 {
  CStdioFile F;

  F.Open(dlg.GetPathName(),CFile::modeRead|CFile::typeText);
  //��ʾ�ļ�����
  CWnd* pWnd = GetDlgItem(IDC_STATIC1);
  CString strTemp("�ļ����أ�"+dlg.GetPathName());   
  pWnd->SetWindowText(_T(strTemp)); 
  //��ʾ������
  //��ʾ����ļ�����
  CWnd* pResultWnd = GetDlgItem(IDC_STATIC3);
  CString strReTemp("���н�����ļ����غ��ַ���ת���У���ȴ�...");   
  pResultWnd->SetWindowText(_T(strReTemp));
  ((CButton*)GetDlgItem(IDC_BUTTON1))->EnableWindow(FALSE); //���ð�ť
  while(F.ReadString(str1))
  {
   str2+=str1;
   str2+="\r\n";
  }  
  if(IsTextUTF8(str2.GetBuffer(0),str2.GetLength()))
  {
     c= MultiByteToWideChar(CP_UTF8,0,str2.GetBuffer(0),-1,NULL,0);
     lpWideCharStr = new wchar_t[c + 1];
	 wprintf(L"",lpWideCharStr); 
     MultiByteToWideChar (CP_UTF8,0,str2.GetBuffer(0),str2.GetLength(),lpWideCharStr,c);
  }
  SetDlgItemText(IDC_EDIT1,CString(lpWideCharStr));
  F.Close();
  pResultWnd = GetDlgItem(IDC_STATIC3);
  strReTemp=_T("���н�����ļ��Ѿ����أ�����ת��WORD�ļ�...");   
  pResultWnd->SetWindowText(_T(strReTemp));
 }
 ((CButton*)GetDlgItem(IDC_BUTTON1))->EnableWindow(TRUE); //���ð�ť

}

int IsTextUTF8(char* str,ULONGLONG length)
{
	int i;
	DWORD nBytes=0;//UFT8����1-6���ֽڱ���,ASCII��һ���ֽ�
	UCHAR chr;
	BOOL bAllAscii=TRUE; //���ȫ������ASCII, ˵������UTF-8 
	for(i=0;i<length;i++)
	{
		chr= *(str+i);
		
		// �ж��Ƿ�ASCII����,�������,˵���п�����UTF-8,ASCII��7λ����,����һ���ֽڴ�,���λ���Ϊ0,o0xxxxxxx
		if( (chr&0x80) != 0 )
			bAllAscii= FALSE;
 
		if(nBytes==0) //�������ASCII��,Ӧ���Ƕ��ֽڷ�,�����ֽ���
		{
			if(chr>=0x80)
			{
				if(chr>=0xFC&&chr<=0xFD)
					nBytes=6;
				else if(chr>=0xF8)
					nBytes=5;
				else if(chr>=0xF0)
					nBytes=4;
				else if(chr>=0xE0)
					nBytes=3;
				else if(chr>=0xC0)
					nBytes=2;
				else
				{
					return FALSE;
				}
				nBytes--;
			}
		}
		else //���ֽڷ��ķ����ֽ�,ӦΪ 10xxxxxx
		{
			if( (chr&0xC0) != 0x80 )
			{
				return FALSE;
			}
			nBytes--;
		}
	}
 
	if( nBytes > 0 ) //Υ������
	{
		return FALSE;
	}
	
	if( bAllAscii ) //���ȫ������ASCII, ˵������UTF-8
	{
		return FALSE;
	}
	
	return TRUE;
}
int IsUTF8File(CString src)
{
	CString Utf8CharSetStrSrc("character_set_client = utf8");  
	if(src.Find(Utf8CharSetStrSrc))
	{
		return 1;
	}
	else
	{
		return 0;
	}
}
void SplitCString(const CString& _cstr, const CString& _flag, CStringArray& _resultArray)
{
	CString strSrc(_cstr);
 
	CStringArray& strResult = _resultArray;
	CString strLeft = _T("");
 
	int nPos = strSrc.Find(_flag);
	while(0 <= nPos)
	{
		strLeft = strSrc.Left(nPos);
		if (!strLeft.IsEmpty())
		{
			strResult.Add(strLeft);
		}
		strSrc = strSrc.Right(strSrc.GetLength() - nPos - 1);
		nPos = strSrc.Find(_flag);
	}
 
	if (!strSrc.IsEmpty()) {
		strResult.Add(strSrc);
	}
}
CString GetPrimaryKey(CString src)
{
   CString PrimaryKeyStrSrc("PRIMARY KEY");  
   CString LeftBrackets( _T("(") );
   CString RightBrackets( _T(")") );
   CString SingleQuotes( _T("'") );   
   CStringArray _resultStatement;

   SplitCString(src,_T(" "),_resultStatement);
   return  ((CString)((CString)_resultStatement.GetAt(2).Replace(LeftBrackets,"")).Replace(RightBrackets,"")).Replace(SingleQuotes,"");
   
}
CString GetEngine(CString src)
{
	CStringArray _resultStatement;
	CString EngineStrSrc("ENGINE");
	SplitCString(src,_T(" "),_resultStatement);
	return  (CString)((CString)(_resultStatement.GetAt(0)).Replace(EngineStrSrc,"")).Replace(_T("="),"");
}
int GetTableCount(CString SQLstr)
{
   int   size=0,i=0;
   int   resultsize=0;
   CStringArray _resultAllStatement;
   CStringArray _resultAllTable;
   CString CreateStrSrc("CREATE TABLE");
   //���հ������ȷֿ�����������һ�������
   SplitCString(SQLstr,_T(";"),_resultAllStatement);
   size=_resultAllStatement.GetSize();   //�ó���
   for(i=0;i<size;i++)
   {      //����������ľ��Ǵ������
	   if(_resultAllStatement[i].Find(CreateStrSrc)>=0)
	   {
		   _resultAllTable.Add(_resultAllStatement[i]);
	   }
   }
   resultsize=_resultAllTable.GetSize();   //��ĸ���

   return resultsize;
}
//////����SQL�ļ�
struct SqlDataBase * Paraser(CString SQLstr,struct SqlDataBase *DB)
{
   int   size=0,i=0,j=0,n=0,m=0;
   int   resultsize=0;
   int   FieldSize=0;
   int   itemSize=0;
   //�ַ�������
   CString CreateStrSrc("CREATE TABLE");
   CString EngineStrSrc("ENGINE");
   CString PrimaryKeyStrSrc("PRIMARY KEY");
   CString CommentStrSrc("COMMENT");
   CString DefaultStrSrc("DEFAULT");
   CString SingleQuotes( _T("`") );   
   CString WSingleQuotes( _T("'") );   
   //�������
   CString LeftBrackets( _T("(") );
   CStringArray _resultAllStatement;
   CStringArray _resultAllTable;
   CStringArray _resultTable;
   CStringArray _resultField;
   SqlDataBase *PreTable;
   SqlDataBase *CurTable;
   SQLField *SQLf;
   SQLField *preSQLf;//,*curSQLf;
   CString CurStr("");
   //���հ������ȷֿ�����������һ�������
   SplitCString(SQLstr,_T(";"),_resultAllStatement);
   size=_resultAllStatement.GetSize();   //�ó���
   for(i=0;i<size;i++)
   {      //����������ľ��Ǵ������
	   if(_resultAllStatement[i].Find(CreateStrSrc)>=0)
	   {
		   _resultAllTable.Add(_resultAllStatement[i]);
	   }
   }
   resultsize=_resultAllTable.GetSize();   //��ĸ���   
   for(i=0;i<resultsize;i++)
   {
	   //����һ������ڴ�
	  	SQLTable *SQLt;
	    SQLt = (SQLTable *)malloc(sizeof(SQLTable));
	    if(NULL == SQLt)
		{
		   printf("%s\n", "SQLt malloc failed");
		   return 0;
		}
		memset(SQLt,0,sizeof(SQLTable));
	  //��װ����������
      _resultTable.RemoveAll();
      SplitCString(_resultAllTable[i],_T("\r\n"),_resultTable);
	  FieldSize=_resultTable.GetSize();   //һ������ֶ�����
	  //����ֶθ������ټ�ȥ3���������֣����棬�ؼ��֣�
      SQLt->FieldCount=FieldSize-3;
	  //�ֶβ��������ڴ�
      struct SqlField *TableField;
      TableField = (SqlField *)malloc(sizeof(SqlField)*FieldSize);
	  if(NULL == TableField)
		{
		   printf("%s\n", "TableField malloc failed");
		   return 0;
		}
	  memset(TableField,0,sizeof(SqlField)*FieldSize);
      //SQLt->TableField=TableField;
	  for(int j=0;j<FieldSize;j++)
	  {	
		 CurStr= _resultTable[j];
        //�ҵ�������
		if(_resultTable[j].Find(CreateStrSrc)>=0)
		{
			  //�滻ȥ������Ҫ���ַ����ͷ��ţ���ʣ�����֡�
			   CString TableName=_resultTable.GetAt(j);
			   TableName.Replace(CreateStrSrc,"");
			   TableName.Replace(LeftBrackets,"");
			   TableName.Replace(SingleQuotes,"");
			   TableName.TrimLeft();
			   TableName.TrimRight();
               memcpy(SQLt->TableName,TableName.GetBuffer(0),TableName.GetLength());
			   continue;
		 }
		 if(_resultTable[j].Find(EngineStrSrc)>=0)
		 {//ʹ�õ�����
			  memcpy(SQLt->TableEngine,GetEngine(_resultTable[j].GetBuffer(0)),GetEngine(_resultTable[j].GetBuffer(0)).GetLength());
			  continue;
		 }
		 if(_resultTable[j].Find(PrimaryKeyStrSrc)>=0)
		 {//�ؼ���.GetAt(j+1).
             memcpy(SQLt->TableKey,GetPrimaryKey(_resultTable[j].GetBuffer(0)),GetPrimaryKey(_resultTable[j].GetBuffer(0)).GetLength());
			 continue;
		 }
		 else
      {        
		//�����ֶνṹ�ڴ�
	  	SQLf = (SQLField *)malloc(sizeof(SQLField));
	    if(NULL == SQLf)
		{
		     printf("%s\n", "SQLf malloc failed");
		     return 0;
		}
		memset(SQLf,0,sizeof(SQLField));
        SQLf->NextField=NULL;
        //�ֶ��ڻ��ָ���   
        _resultField.RemoveAll();
		SplitCString(_resultTable[j],_T(" "),_resultField);
        itemSize=_resultField.GetSize(); 
		if(itemSize>2)
        {
			    //�ֶδ���               
			    //��ȡ�ֶ���
			    CString strFieldName=_resultField.GetAt(1); 
				strFieldName.Replace(SingleQuotes,"");
			    //��ȡ�ֶ����ͺͳ���
			    CString strType=_resultField.GetAt(2);    
                memcpy(SQLf->FieldName,strFieldName.GetBuffer(0),strFieldName.GetLength());
			    memcpy(SQLf->FieldType,strType.GetBuffer(0),strType.GetLength());
			  }
			  for(n=0;n<itemSize;n++)
			  {			   
			  //��ȡ��ע
                if((_resultField[n].CompareNoCase(CommentStrSrc)==0)&&((n+1)<=itemSize))
				 {
				      CString strFieldComment=_resultField.GetAt(n+1);
					  strFieldComment.Replace(SingleQuotes,"");
					  strFieldComment.Replace(WSingleQuotes,"");
					  if(TABLE_FIELD_COMMENTLEN>strFieldComment.GetLength())
					  {
						  memcpy(SQLf->FieldComment,strFieldComment.GetBuffer(0),strFieldComment.GetLength());
					  }
					  else{
						  memcpy(SQLf->FieldComment,strFieldComment.GetBuffer(0),TABLE_FIELD_COMMENTLEN-1);
					  }
				  }
              //��ȡȱʡֵ
				 if((_resultField[n].CompareNoCase(DefaultStrSrc)==0)&&((n+1)<=itemSize))
				 {
				   	CString strFieldDefault=_resultField.GetAt(n+1);
					memcpy(SQLf->FieldDefault,strFieldDefault.GetBuffer(0),strFieldDefault.GetLength());
				 }
			  }//end һ���ֶ��ڲ�
		  }//end һ���ֶ�
	  //���ֶμӵ�������
	  if(SQLt->TableField==NULL)
	  {
		  SQLt->TableField=SQLf;
          SQLt->TableField->NextField=NULL;
          memcpy(SQLt->TableField->FieldName,SQLf->FieldName,TABLE_FIELD_NAMELEN);
          memcpy(SQLt->TableField->FieldType,SQLf->FieldType,TABLE_FIELD_TYPELEN);
		  preSQLf=SQLf;
          
	  }else
	  {
		 preSQLf->NextField=SQLf;
         SQLf->NextField=NULL;
         memcpy(preSQLf->NextField->FieldName,SQLf->FieldName,TABLE_FIELD_NAMELEN);
         memcpy(preSQLf->NextField->FieldType,SQLf->FieldType,TABLE_FIELD_TYPELEN);
		 preSQLf=SQLf;
	  }     

	  }//end һ����
	  //��ӵ�����������
      //����һ��database
	  CurTable = (SqlDataBase *)malloc(sizeof(SqlDataBase));
      if(NULL == CurTable)
	  {
	     printf("%s\n", "CurTable malloc failed");
	     return 0;
	  }
	  memset(CurTable,0,sizeof(SqlDataBase));
      CurTable->NextTable=NULL;
      memcpy(&CurTable->Table,SQLt,sizeof(SQLTable));
	  if(i==0)
	  {
          DB=CurTable;          		  
	  }else
	  {
          PreTable->NextTable=CurTable;
	  }	  	  
	  PreTable=CurTable;	      
   }//end һ���ļ�
   //���ظ�������
   return DB;
}
//���ñ��ı�ͷ
void SetWordTableTitle(WordHandle wordHandle)
{
	wordHandle.WriteCell(1, 1, _T("�ֶ�"));
	wordHandle.WriteCell(1, 2, _T("�ֶ�����"));
	wordHandle.WriteCell(1, 3, _T("ȱʡֵ"));
	wordHandle.WriteCell(1, 4, _T("��ע"));
}
//�����ĵ��ı���
void SetWordTitle(WordHandle wordHandle)
{
	wordHandle.AddParagraph();
	wordHandle.SetParagraphFormat(wdAlignParagraphCenter);
	wordHandle.SetFont(_T("����"), 18);
	wordHandle.WriteText(_T("���ݿ���Ʊ���"));
}
//���ö��䣬�ĵ�����ʱ��
void SetWordGenTime(WordHandle wordHandle)
{
    wordHandle.AddParagraph();
	wordHandle.AddParagraph();
	wordHandle.SetFont(_T("����"), 10);
	wordHandle.SetParagraphFormat(wdAlignParagraphRight);
	wordHandle.WriteText(_T("����ʱ�䣺"));
	wordHandle.WriteText(CTime::GetCurrentTime().Format(_T("%Y-%m-%d %H:%M:%S")));

}
LPCTSTR Char2LPCTSTR(char src[])
{
   //����CString������������еı���ת��ʵ��
   CString cstr = src;
   LPCTSTR pStr = LPCTSTR(cstr);
   return pStr;

}
CString GetFilePath()
{
	char fileName[MAX_PATH];
	GetModuleFileName(NULL, fileName, MAX_PATH);
	char dir[260];
	char dirver[100];
	_splitpath(fileName, dirver, dir, NULL, NULL);
	CString strAppPath = dirver;
	strAppPath += dir;
	return strAppPath;

}
void GenWordTitle(WordHandle wordHandle)
{   
	wordHandle.CreateWord();	
	//��ȡӦ�õ�ǰDebug·��
	//wordHandle.AddPicture(GetFilePath()+"\\logo.bmp");
	//�����ĵ��ı���
	SetWordTitle(wordHandle);
	//���ö��䣬�ĵ�����ʱ��
	SetWordGenTime(wordHandle);
}
CString GenWordFileName()
{
	SYSTEMTIME st;
	CString strDate,strTime,filename;
	GetLocalTime(&st);
    strDate.Format("%4d_%2d_%2d_%2d_%2d_%2d",st.wYear,st.wMonth,st.wDay,st.wHour,st.wMinute,st.wSecond);    
    filename.Format("%s\\���ݿ�%s.docx",GetFilePath().GetBuffer(0),strDate.GetBuffer(0));
	return filename;

}
void SaveWordFile(WordHandle wordHandle)
{
    wordHandle.CloseWordSave(GenWordFileName());
 }
void GenFromSQLToWord(struct SqlDataBase *DB)
{
	struct SqlDataBase *Head=NULL;
	int TableCount = 1;
	int FieldCount = 0;
	int i = 2;
	CString TableSerial;
	WordHandle wordHandle;
	wordHandle.CreateWord();	
	//��ȡӦ�õ�ǰ·��
	//wordHandle.AddPicture(GetFilePath()+"\\logo.bmp");
	//�����ĵ��ı���
	SetWordTitle(wordHandle);
	//���ö��䣬�ĵ�����ʱ��
	SetWordGenTime(wordHandle);
    //GenWordTitle(wordHandle);

    Head=DB;
    while(Head!=NULL)
	{  
	   //���ñ�ͷ
	   wordHandle.AddParagraph();
	   wordHandle.AddParagraph();		  
	   wordHandle.SetParagraphFormat(wdAlignParagraphLeft);	   
	   TableSerial.Format(_T("��%d��"),TableCount);
	   TableCount++;
	   wordHandle.WriteText(TableSerial);
	   wordHandle.WriteText(_T(Head->Table.TableName));
       wordHandle.WriteText(_T(Head->Table.TableEngine));              
	   //����һ�����
       //�������ֶ�����һ
       FieldCount=Head->Table.FieldCount;
       FieldCount=FieldCount+1;
	   wordHandle.AddParagraph();	
	   wordHandle.CreateTable(FieldCount,4);
	  	//����������12
	   wordHandle.SetTableFont(1, 1, FieldCount, 4, _T("����"), 12);
	   wordHandle.SetTableFormat(1, 1, FieldCount, 4, wdAlignParagraphCenter, wdCellAlignVerticalCenter);
	   wordHandle.SetTablePadding();
	   //���ñ���ͷ����һ�У�
	   wordHandle.WriteCell(1, 1, _T("�ֶ�"));
	   wordHandle.WriteCell(1, 2, _T("�ֶ�����"));
	   wordHandle.WriteCell(1, 3, _T("ȱʡֵ"));
	   wordHandle.WriteCell(1, 4, _T("��ע"));
       //SetWordTableTitle(wordHandle);
       //��������
	   struct SqlField *CurField; 
       if(Head->Table.TableField!=NULL)
	   {
          CurField=Head->Table.TableField;
	   }
	   while(CurField!=NULL)
	   { 
		  wordHandle.WriteCell(i, 1, _T(LPCTSTR(CurField->FieldName)));
	      wordHandle.WriteCell(i, 2, _T(LPCTSTR(CurField->FieldType)));  
	      wordHandle.WriteCell(i, 3, _T(LPCTSTR(CurField->FieldDefault)));
		  wordHandle.WriteCell(i, 4, _T(LPCTSTR(CurField->FieldComment)));		  
		  i++;
		  CurField=CurField->NextField;

	   }
       Head=Head->NextTable;
	   wordHandle.AddParagraph();	
	   i = 2;
	}
 SaveWordFile(wordHandle);  
 //wordHandle.ShowWord(TRUE); //���ÿɼ� 
}
void FreeField(struct SqlField *TableField)
{
	int freecount=0;
	struct SqlField *CurField=NULL;    
    while(TableField!=NULL)
	{
	   CurField=TableField->NextField;
	   free(TableField);       
	   TableField=CurField;
	   freecount++;
	}
}
void FreeDB(struct SqlDataBase *Head)
{
	int tablecount=0;
	struct SqlDataBase *CurDB=NULL;   
	while(Head!=NULL)
	{
       CurDB=Head->NextTable;
       FreeField(Head->Table.TableField);
	   free(Head);
	   Head = CurDB;
	   tablecount++;       
	}

}
void CSql2wordDlg::OnButton3() 
{
	struct SqlDataBase *Head=NULL;   
    CString SQLstr;
	//��ȡSQL����
	GetDlgItemText(IDC_EDIT1,SQLstr);
	if(SQLstr.GetLength()==0)
		return;

	((CButton*)GetDlgItem(IDC_BUTTON3))->EnableWindow(FALSE); //���ð�ť
     //��ʾ���
    CWnd* pResultWnd = GetDlgItem(IDC_STATIC3);
    CString strReTemp("���н�����ļ������У���ȴ�...");   
    pResultWnd->SetWindowText(_T(strReTemp));
    //��������
    Head = Paraser(SQLstr,Head);
    if(Head==NULL)
	{
	    printf("The SQL no create table");
	}else
	{
	 printf("The SQL create table as follow");
	}
	//����word
	GenFromSQLToWord(Head);
	//�ͷ��ڴ�
	FreeDB(Head);
   //��ʾ����ļ�����
   CWnd* pWnd = GetDlgItem(IDC_STATIC3);
   CString strTemp("���н�����ļ�Ϊ"+GenWordFileName());   
   pWnd->SetWindowText(_T(strTemp)); 
   ((CButton*)GetDlgItem(IDC_BUTTON3))->EnableWindow(TRUE); //���ð�ť
}

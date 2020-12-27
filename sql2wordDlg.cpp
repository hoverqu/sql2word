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
//打开文件对话框，装载sql文件
int TableCount=0,i=0;
int c;
LPWSTR lpWideCharStr;
CString str1,str2;
CFileDialog dlg (TRUE, NULL,NULL, OFN_HIDEREADONLY|OFN_OVERWRITEPROMPT|OFN_ALLOWMULTISELECT,"All Files (*.sql)|*.sql| |", AfxGetMainWnd());
if (dlg.DoModal () == IDOK)
 {
  CStdioFile F;

  F.Open(dlg.GetPathName(),CFile::modeRead|CFile::typeText);
  //显示文件名字
  CWnd* pWnd = GetDlgItem(IDC_STATIC1);
  CString strTemp("文件加载："+dlg.GetPathName());   
  pWnd->SetWindowText(_T(strTemp)); 
  //显示加载中
  //显示结果文件名字
  CWnd* pResultWnd = GetDlgItem(IDC_STATIC3);
  CString strReTemp("运行结果：文件加载和字符集转换中，请等待...");   
  pResultWnd->SetWindowText(_T(strReTemp));
  ((CButton*)GetDlgItem(IDC_BUTTON1))->EnableWindow(FALSE); //禁用按钮
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
  strReTemp=_T("运行结果：文件已经加载，请点击转换WORD文件...");   
  pResultWnd->SetWindowText(_T(strReTemp));
 }
 ((CButton*)GetDlgItem(IDC_BUTTON1))->EnableWindow(TRUE); //启用按钮

}

int IsTextUTF8(char* str,ULONGLONG length)
{
	int i;
	DWORD nBytes=0;//UFT8可用1-6个字节编码,ASCII用一个字节
	UCHAR chr;
	BOOL bAllAscii=TRUE; //如果全部都是ASCII, 说明不是UTF-8 
	for(i=0;i<length;i++)
	{
		chr= *(str+i);
		
		// 判断是否ASCII编码,如果不是,说明有可能是UTF-8,ASCII用7位编码,但用一个字节存,最高位标记为0,o0xxxxxxx
		if( (chr&0x80) != 0 )
			bAllAscii= FALSE;
 
		if(nBytes==0) //如果不是ASCII码,应该是多字节符,计算字节数
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
		else //多字节符的非首字节,应为 10xxxxxx
		{
			if( (chr&0xC0) != 0x80 )
			{
				return FALSE;
			}
			nBytes--;
		}
	}
 
	if( nBytes > 0 ) //违返规则
	{
		return FALSE;
	}
	
	if( bAllAscii ) //如果全部都是ASCII, 说明不是UTF-8
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
   //按照包含；先分开，基本上是一条长语句
   SplitCString(SQLstr,_T(";"),_resultAllStatement);
   size=_resultAllStatement.GetSize();   //得长度
   for(i=0;i<size;i++)
   {      //包含创建表的就是创建表的
	   if(_resultAllStatement[i].Find(CreateStrSrc)>=0)
	   {
		   _resultAllTable.Add(_resultAllStatement[i]);
	   }
   }
   resultsize=_resultAllTable.GetSize();   //表的个数

   return resultsize;
}
//////解析SQL文件
struct SqlDataBase * Paraser(CString SQLstr,struct SqlDataBase *DB)
{
   int   size=0,i=0,j=0,n=0,m=0;
   int   resultsize=0;
   int   FieldSize=0;
   int   itemSize=0;
   //字符串定义
   CString CreateStrSrc("CREATE TABLE");
   CString EngineStrSrc("ENGINE");
   CString PrimaryKeyStrSrc("PRIMARY KEY");
   CString CommentStrSrc("COMMENT");
   CString DefaultStrSrc("DEFAULT");
   CString SingleQuotes( _T("`") );   
   CString WSingleQuotes( _T("'") );   
   //左边括号
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
   //按照包含；先分开，基本上是一条长语句
   SplitCString(SQLstr,_T(";"),_resultAllStatement);
   size=_resultAllStatement.GetSize();   //得长度
   for(i=0;i<size;i++)
   {      //包含创建表的就是创建表的
	   if(_resultAllStatement[i].Find(CreateStrSrc)>=0)
	   {
		   _resultAllTable.Add(_resultAllStatement[i]);
	   }
   }
   resultsize=_resultAllTable.GetSize();   //表的个数   
   for(i=0;i<resultsize;i++)
   {
	   //申请一个表的内存
	  	SQLTable *SQLt;
	    SQLt = (SQLTable *)malloc(sizeof(SQLTable));
	    if(NULL == SQLt)
		{
		   printf("%s\n", "SQLt malloc failed");
		   return 0;
		}
		memset(SQLt,0,sizeof(SQLTable));
	  //安装分行来划分
      _resultTable.RemoveAll();
      SplitCString(_resultAllTable[i],_T("\r\n"),_resultTable);
	  FieldSize=_resultTable.GetSize();   //一个表的字段数量
	  //表的字段个数是再减去3个（表名字，引擎，关键字）
      SQLt->FieldCount=FieldSize-3;
	  //字段部分申请内存
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
        //找到表名称
		if(_resultTable[j].Find(CreateStrSrc)>=0)
		{
			  //替换去掉不需要的字符串和符号，就剩下名字。
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
		 {//使用的引擎
			  memcpy(SQLt->TableEngine,GetEngine(_resultTable[j].GetBuffer(0)),GetEngine(_resultTable[j].GetBuffer(0)).GetLength());
			  continue;
		 }
		 if(_resultTable[j].Find(PrimaryKeyStrSrc)>=0)
		 {//关键字.GetAt(j+1).
             memcpy(SQLt->TableKey,GetPrimaryKey(_resultTable[j].GetBuffer(0)),GetPrimaryKey(_resultTable[j].GetBuffer(0)).GetLength());
			 continue;
		 }
		 else
      {        
		//申请字段结构内存
	  	SQLf = (SQLField *)malloc(sizeof(SQLField));
	    if(NULL == SQLf)
		{
		     printf("%s\n", "SQLf malloc failed");
		     return 0;
		}
		memset(SQLf,0,sizeof(SQLField));
        SQLf->NextField=NULL;
        //字段内划分各项   
        _resultField.RemoveAll();
		SplitCString(_resultTable[j],_T(" "),_resultField);
        itemSize=_resultField.GetSize(); 
		if(itemSize>2)
        {
			    //字段处理               
			    //获取字段名
			    CString strFieldName=_resultField.GetAt(1); 
				strFieldName.Replace(SingleQuotes,"");
			    //获取字段类型和长度
			    CString strType=_resultField.GetAt(2);    
                memcpy(SQLf->FieldName,strFieldName.GetBuffer(0),strFieldName.GetLength());
			    memcpy(SQLf->FieldType,strType.GetBuffer(0),strType.GetLength());
			  }
			  for(n=0;n<itemSize;n++)
			  {			   
			  //获取备注
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
              //获取缺省值
				 if((_resultField[n].CompareNoCase(DefaultStrSrc)==0)&&((n+1)<=itemSize))
				 {
				   	CString strFieldDefault=_resultField.GetAt(n+1);
					memcpy(SQLf->FieldDefault,strFieldDefault.GetBuffer(0),strFieldDefault.GetLength());
				 }
			  }//end 一个字段内部
		  }//end 一个字段
	  //把字段加到表里面
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

	  }//end 一个表
	  //添加到表数组里面
      //申请一个database
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
   }//end 一个文件
   //返回各表数组
   return DB;
}
//设置表格的表头
void SetWordTableTitle(WordHandle wordHandle)
{
	wordHandle.WriteCell(1, 1, _T("字段"));
	wordHandle.WriteCell(1, 2, _T("字段类型"));
	wordHandle.WriteCell(1, 3, _T("缺省值"));
	wordHandle.WriteCell(1, 4, _T("备注"));
}
//设置文档的标题
void SetWordTitle(WordHandle wordHandle)
{
	wordHandle.AddParagraph();
	wordHandle.SetParagraphFormat(wdAlignParagraphCenter);
	wordHandle.SetFont(_T("黑体"), 18);
	wordHandle.WriteText(_T("数据库设计报告"));
}
//设置段落，文档生成时间
void SetWordGenTime(WordHandle wordHandle)
{
    wordHandle.AddParagraph();
	wordHandle.AddParagraph();
	wordHandle.SetFont(_T("宋体"), 10);
	wordHandle.SetParagraphFormat(wdAlignParagraphRight);
	wordHandle.WriteText(_T("生成时间："));
	wordHandle.WriteText(CTime::GetCurrentTime().Format(_T("%Y-%m-%d %H:%M:%S")));

}
LPCTSTR Char2LPCTSTR(char src[])
{
   //利用CString的运算符重载中的编码转换实现
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
	//获取应用当前Debug路径
	//wordHandle.AddPicture(GetFilePath()+"\\logo.bmp");
	//设置文档的标题
	SetWordTitle(wordHandle);
	//设置段落，文档生成时间
	SetWordGenTime(wordHandle);
}
CString GenWordFileName()
{
	SYSTEMTIME st;
	CString strDate,strTime,filename;
	GetLocalTime(&st);
    strDate.Format("%4d_%2d_%2d_%2d_%2d_%2d",st.wYear,st.wMonth,st.wDay,st.wHour,st.wMinute,st.wSecond);    
    filename.Format("%s\\数据库%s.docx",GetFilePath().GetBuffer(0),strDate.GetBuffer(0));
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
	//获取应用当前路径
	//wordHandle.AddPicture(GetFilePath()+"\\logo.bmp");
	//设置文档的标题
	SetWordTitle(wordHandle);
	//设置段落，文档生成时间
	SetWordGenTime(wordHandle);
    //GenWordTitle(wordHandle);

    Head=DB;
    while(Head!=NULL)
	{  
	   //设置表头
	   wordHandle.AddParagraph();
	   wordHandle.AddParagraph();		  
	   wordHandle.SetParagraphFormat(wdAlignParagraphLeft);	   
	   TableSerial.Format(_T("表%d："),TableCount);
	   TableCount++;
	   wordHandle.WriteText(TableSerial);
	   wordHandle.WriteText(_T(Head->Table.TableName));
       wordHandle.WriteText(_T(Head->Table.TableEngine));              
	   //创建一个表格
       //行数是字段数加一
       FieldCount=Head->Table.FieldCount;
       FieldCount=FieldCount+1;
	   wordHandle.AddParagraph();	
	   wordHandle.CreateTable(FieldCount,4);
	  	//字体是宋体12
	   wordHandle.SetTableFont(1, 1, FieldCount, 4, _T("宋体"), 12);
	   wordHandle.SetTableFormat(1, 1, FieldCount, 4, wdAlignParagraphCenter, wdCellAlignVerticalCenter);
	   wordHandle.SetTablePadding();
	   //设置表格表头（第一行）
	   wordHandle.WriteCell(1, 1, _T("字段"));
	   wordHandle.WriteCell(1, 2, _T("字段类型"));
	   wordHandle.WriteCell(1, 3, _T("缺省值"));
	   wordHandle.WriteCell(1, 4, _T("备注"));
       //SetWordTableTitle(wordHandle);
       //设置内容
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
 //wordHandle.ShowWord(TRUE); //设置可见 
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
	//获取SQL内容
	GetDlgItemText(IDC_EDIT1,SQLstr);
	if(SQLstr.GetLength()==0)
		return;

	((CButton*)GetDlgItem(IDC_BUTTON3))->EnableWindow(FALSE); //禁用按钮
     //显示结果
    CWnd* pResultWnd = GetDlgItem(IDC_STATIC3);
    CString strReTemp("运行结果：文件生成中，请等待...");   
    pResultWnd->SetWindowText(_T(strReTemp));
    //解析内容
    Head = Paraser(SQLstr,Head);
    if(Head==NULL)
	{
	    printf("The SQL no create table");
	}else
	{
	 printf("The SQL create table as follow");
	}
	//生成word
	GenFromSQLToWord(Head);
	//释放内存
	FreeDB(Head);
   //显示结果文件名字
   CWnd* pWnd = GetDlgItem(IDC_STATIC3);
   CString strTemp("运行结果：文件为"+GenWordFileName());   
   pWnd->SetWindowText(_T(strTemp)); 
   ((CButton*)GetDlgItem(IDC_BUTTON3))->EnableWindow(TRUE); //启用按钮
}

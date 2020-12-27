// WordHandle.cpp: implementation of the WordHandle class.
//
//////////////////////////////////////////////////////////////////////
 
#include "stdafx.h"
#include "WordHandle.h"
 
#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif
 
//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////
 
WordHandle::WordHandle(){
	OleInitialize(NULL);    
	//CoInitialize(NULL);
	ComVariantTrue=CComVariant(true);
	ComVariantFalse=CComVariant(false);
	vTrue=COleVariant((short)TRUE);
	vFalse=COleVariant ((short)FALSE);
	tpl=CComVariant(_T(""));
	DocType=CComVariant(0);
	NewTemplate=CComVariant(false);
	vOptional=COleVariant((long)DISP_E_PARAMNOTFOUND,VT_ERROR);
 
}
 
WordHandle::~WordHandle(){
	wordtable.ReleaseDispatch();
	wordtables.ReleaseDispatch();
	wordsel.ReleaseDispatch();
	worddocs.ReleaseDispatch();
	worddoc.ReleaseDispatch();
	paragraphformat.ReleaseDispatch();
	wordapp.ReleaseDispatch();
	wordFt.ReleaseDispatch();
	cell.ReleaseDispatch();
	range.ReleaseDispatch();
	ishaps.ReleaseDispatch();
	ishap.ReleaseDispatch();
	CoUninitialize();
}
 
BOOL WordHandle::CreateApp(){
	if (FALSE == wordapp.CreateDispatch("Word.Application"))	{
		MessageBox(NULL,"Application创建失败!","", MB_OK|MB_ICONWARNING);
		return FALSE;
	}
	//    m_wdApp.SetVisible(TRUE);
	return TRUE;
 
}
 
BOOL WordHandle::CreateDocumtent(){
	if (!wordapp.m_lpDispatch){
       MessageBox(NULL,"Application为空,Documents创建失败!", "错误提示",MB_OK|MB_ICONWARNING);
       return FALSE;
   }
   worddocs.AttachDispatch(wordapp.GetDocuments());
   if (!worddocs.m_lpDispatch){
       MessageBox(NULL,"Documents创建失败!", "错误提示",MB_OK|MB_ICONWARNING);
       return FALSE;
   }
   CComVariant Template(_T(""));    
   CComVariant NewTemplate(false),DocumentType(0),Visible;
   worddocs.Add(&Template,&NewTemplate,&DocumentType,&Visible);    
 
   worddocs = wordapp.GetActiveDocument();
   if (!worddocs.m_lpDispatch){
       MessageBox(NULL,"Document获取失败!", "",MB_OK|MB_ICONWARNING);
       return FALSE;
   }
   wordsel = wordapp.GetSelection();
   if (!wordsel.m_lpDispatch){
       MessageBox(NULL,"Select获取失败!", "",MB_OK|MB_ICONWARNING);
       return FALSE;
   }    
   return TRUE;
 
}
 
void WordHandle::SetParagraphFormat(int nAlignment)
{
	if (!wordsel.m_lpDispatch){
		return ;
	}
	paragraphformat = wordsel.GetParagraphFormat();
	paragraphformat.SetAlignment(nAlignment);
	wordsel.SetParagraphFormat(paragraphformat);
}
 
BOOL WordHandle::CreateTable(int nRow, int nColumn){
	if (!wordapp.m_lpDispatch){
		return FALSE;
	}
 
	worddoc = wordapp.GetActiveDocument();
	wordtables = worddoc.GetTables();
	VARIANT vtDefault, vtAuto;
	vtDefault.vt = VT_I4;
	vtAuto.vt = VT_I4;
	vtDefault.intVal = 1;
	vtAuto.intVal = 0;
	wordtables.Add(wordsel.GetRange(), nRow, nColumn, &vtDefault, &vtAuto);    
	wordtable = wordtables.Item(wordtables.GetCount());
	
	VARIANT vtstyle;
	vtstyle.vt = VT_BSTR;
	_bstr_t bstr = "网格型";
	vtstyle.bstrVal = bstr;
	if (wordtable.GetStyle().bstrVal == vtstyle.bstrVal){
		wordtable.SetStyle(&vtstyle);
		wordtable.SetApplyStyleFirstColumn(TRUE);
		wordtable.SetApplyStyleHeadingRows(TRUE);
		wordtable.SetApplyStyleLastColumn(TRUE);
		wordtable.SetApplyStyleLastRow(TRUE);        
	}
	return TRUE;
}
 
void WordHandle::WriteCell(int nRow, int nColumne, LPCTSTR szText){
	if (wordtable!=NULL){
		cell = wordtable.Cell(nRow,nColumne);
		cell.Select();
		wordsel.TypeText(szText);		
	}	
}
 
//将数字num以lpszFormat的形式写入（nRow，nColumne）中
void WordHandle::WriteCell(int nRow, int nColumne, LPCTSTR lpszFormat, float num){ 
	if(wordtable != NULL){
		cell = wordtable.Cell(nRow, nColumne);
		cell.Select();
		CString str;
		str.Format(lpszFormat, num);
		wordsel.TypeText(str);
	}
}
 
void WordHandle::ShowWord(BOOL show){
	if (!wordapp.m_lpDispatch){
		return ;
	}
	wordapp.SetVisible(show);
}
 
BOOL WordHandle::CreateWord(){
	if (FALSE == CreateApp()){
		return FALSE;
	}
	return CreateDocumtent();
}
 
//输出文字
void WordHandle::WriteText(LPCTSTR szText){
	if (!wordsel.m_lpDispatch){
		return ;
	}
	wordsel.TypeText(szText);
	wordtable=NULL;
}
 
//添加新段落
void WordHandle::AddParagraph(){
	if (!wordsel.m_lpDispatch){
		return ;
	}
	wordsel.EndKey(COleVariant((short)6),COleVariant(short(0)));  //定位到全文末尾
	wordsel.TypeParagraph();//新的段落，也就是回车换行
	wordtable=NULL;
}
 
//设置字体样式、字体大小、字体颜色、背景色
void WordHandle::SetFont(LPCTSTR szFontName, float fSize, long lFontColor, long lBackColor){
	if (!wordsel.m_lpDispatch)	{
		MessageBox(NULL,"Select为空，字体设置失败!","错误提示", MB_OK|MB_ICONWARNING);
		return;
	}
	wordsel.SetText("F");
	wordFt = wordsel.GetFont();
	wordFt.SetSize(fSize);
	wordFt.SetName(szFontName);
	wordFt.SetColor(lFontColor);
	wordsel.SetFont(wordFt);
	range = wordsel.GetRange();
	range.SetHighlightColorIndex(lBackColor);
}
 
//设置粗体、斜体、下划线
void WordHandle::SetFont(BOOL bBold, BOOL bItalic, BOOL bUnderLine){
	if (!wordsel.m_lpDispatch){
		MessageBox(NULL,"Select为空，字体设置失败!", "错误提示",MB_OK|MB_ICONWARNING);
		return;
	}
	wordsel.SetText("F");
	wordFt = wordsel.GetFont();
	wordFt.SetBold(bBold);
	wordFt.SetItalic(bItalic);
	wordFt.SetUnderline(bUnderLine);
	wordsel.SetFont(wordFt);
}
 
//设置单元格(nRow, nColumn)字体样式、字体大小、字体颜色、背景色
void WordHandle::SetTableFont(int nRow, int nColumn, LPCTSTR szFontName, float fSize, long lFontColor, long lBackColor){
	cell = wordtable.Cell(nRow, nColumn);
	cell.Select();
	wordFt = wordsel.GetFont();
	wordFt.SetName(szFontName);
	wordFt.SetSize(fSize);
	wordFt.SetColor(lFontColor);
	wordsel.SetFont(wordFt);
	range = wordsel.GetRange();
	range.SetHighlightColorIndex(lBackColor);
}
 
//设置从(beginRow, beginColumn)到(endRow, endColumn)所围区域单元格的字体样式、字体大小、字体颜色、背景色
void WordHandle::SetTableFont(int beginRow, int beginColumn, int endRow, int endColumn, LPCTSTR szFontName, float fSize, long lFontColor, long lBackColor){
	wordtable.SetBottomPadding(10.0);
	int nRow, nColumn;
	for(nRow = beginRow; nRow <= endRow; nRow++){
		for(nColumn = beginColumn; nColumn <= endColumn; nColumn++){
			SetTableFont(nRow, nColumn, szFontName, fSize, lFontColor, lBackColor);
		}
	}
}
 
//设置单元格(nRow, nColumn)粗体、斜体、下划线
void WordHandle::SetTableFont(int nRow, int nColumn, BOOL bBold, BOOL bItalic, BOOL bUnderLine){
	cell = wordtable.Cell(nRow, nColumn);
	cell.Select();
	wordFt = wordsel.GetFont();
	wordFt.SetBold(bBold);
	wordFt.SetItalic(bItalic);
	wordFt.SetUnderline(bUnderLine);
	wordsel.SetFont(wordFt);
}
 
//设置表格单元格(nRow, nColumn)的水平、垂直对齐方式
void WordHandle::SetTableFormat(int nRow, int nColumn, long horizontalAlignment, long verticalAlignment){
	cell = wordtable.Cell(nRow, nColumn);
	cell.Select();
	cell.SetVerticalAlignment(verticalAlignment);
	paragraphformat = wordsel.GetParagraphFormat();
	paragraphformat.SetAlignment(horizontalAlignment);
	wordsel.SetParagraphFormat(paragraphformat);
}
 
//设置表格从(beginRow, beginColumn)到(endRow, endRow)区域内水平、垂直对齐方式
void WordHandle::SetTableFormat(int beginRow, int beginColumn, int endRow, int endColumn, long horizontalAlignment, long verticalAlignment){
	int nRow, nColumn;
	for( nRow = beginRow; nRow <= endRow; nRow++){
		for( nColumn = beginColumn; nColumn <= endColumn; nColumn++){
			SetTableFormat(nRow, nColumn, horizontalAlignment, verticalAlignment);
		}
	}
}
 
//设置表格单元格的上下左右内边距填充
void WordHandle::SetTablePadding(float topPadding, float bottomPadding, float leftPadding, float rightPadding){
	wordtable.SetTopPadding(topPadding);
	wordtable.SetBottomPadding(bottomPadding);
	wordtable.SetLeftPadding(leftPadding);
	wordtable.SetRightPadding(rightPadding);
}
 
//合并单元格
BOOL WordHandle::MergeCell(int  cell1row, int cell1col, int cell2row, int cell2col){
	Cell cell1=wordtable.Cell(cell1row,cell1col);
	Cell cell2=wordtable.Cell(cell2row,cell2col);
	cell1.Merge(cell2);
 
	return TRUE;
}
 
//创建图表
void WordHandle::CreateChart(){
	ishaps=	wordsel.GetInlineShapes();
	CComVariant classt;
	classt=CComVariant("MSGraph.Chart.8");
	ishap=ishaps.AddOLEObject(&classt,vOptional,vFalse,vFalse,vOptional,vOptional,vOptional,vOptional);
	AddParagraph();
}
 
//添加图片
void WordHandle::AddPicture(LPCTSTR picname){
	
	ishaps=	wordsel.GetInlineShapes();	
	ishap=ishaps.AddPicture(picname,vFalse,vTrue,vOptional);   
 
	AddParagraph();
}
 
void WordHandle::AddTime(CTime time){
}
 
void WordHandle::CloseWord(){
	if (!wordapp.m_lpDispatch){
		return ;
	}
	SaveChanges=ComVariantFalse;
	wordapp.Quit(&SaveChanges,&OriginalFormat,&RouteDocument);
}
 
void WordHandle::CloseWordSave(LPCTSTR wordname){
	if (!wordapp.m_lpDispatch){
		return ;
	}
	if (!worddoc.m_lpDispatch){
		return ;
	}
    worddoc.SaveAs(&CComVariant(wordname), 
		&CComVariant((short)0),
		vFalse, &CComVariant(""), vTrue, &CComVariant(""),
		vFalse, vFalse, vFalse, vFalse, vFalse,
		vFalse, vFalse, vFalse, vFalse, vFalse);
	SaveChanges=ComVariantTrue;
    wordapp.Quit(&SaveChanges,&OriginalFormat,&RouteDocument);
}
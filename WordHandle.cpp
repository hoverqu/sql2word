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
		MessageBox(NULL,"Application����ʧ��!","", MB_OK|MB_ICONWARNING);
		return FALSE;
	}
	//    m_wdApp.SetVisible(TRUE);
	return TRUE;
 
}
 
BOOL WordHandle::CreateDocumtent(){
	if (!wordapp.m_lpDispatch){
       MessageBox(NULL,"ApplicationΪ��,Documents����ʧ��!", "������ʾ",MB_OK|MB_ICONWARNING);
       return FALSE;
   }
   worddocs.AttachDispatch(wordapp.GetDocuments());
   if (!worddocs.m_lpDispatch){
       MessageBox(NULL,"Documents����ʧ��!", "������ʾ",MB_OK|MB_ICONWARNING);
       return FALSE;
   }
   CComVariant Template(_T(""));    
   CComVariant NewTemplate(false),DocumentType(0),Visible;
   worddocs.Add(&Template,&NewTemplate,&DocumentType,&Visible);    
 
   worddocs = wordapp.GetActiveDocument();
   if (!worddocs.m_lpDispatch){
       MessageBox(NULL,"Document��ȡʧ��!", "",MB_OK|MB_ICONWARNING);
       return FALSE;
   }
   wordsel = wordapp.GetSelection();
   if (!wordsel.m_lpDispatch){
       MessageBox(NULL,"Select��ȡʧ��!", "",MB_OK|MB_ICONWARNING);
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
	_bstr_t bstr = "������";
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
 
//������num��lpszFormat����ʽд�루nRow��nColumne����
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
 
//�������
void WordHandle::WriteText(LPCTSTR szText){
	if (!wordsel.m_lpDispatch){
		return ;
	}
	wordsel.TypeText(szText);
	wordtable=NULL;
}
 
//����¶���
void WordHandle::AddParagraph(){
	if (!wordsel.m_lpDispatch){
		return ;
	}
	wordsel.EndKey(COleVariant((short)6),COleVariant(short(0)));  //��λ��ȫ��ĩβ
	wordsel.TypeParagraph();//�µĶ��䣬Ҳ���ǻس�����
	wordtable=NULL;
}
 
//����������ʽ�������С��������ɫ������ɫ
void WordHandle::SetFont(LPCTSTR szFontName, float fSize, long lFontColor, long lBackColor){
	if (!wordsel.m_lpDispatch)	{
		MessageBox(NULL,"SelectΪ�գ���������ʧ��!","������ʾ", MB_OK|MB_ICONWARNING);
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
 
//���ô��塢б�塢�»���
void WordHandle::SetFont(BOOL bBold, BOOL bItalic, BOOL bUnderLine){
	if (!wordsel.m_lpDispatch){
		MessageBox(NULL,"SelectΪ�գ���������ʧ��!", "������ʾ",MB_OK|MB_ICONWARNING);
		return;
	}
	wordsel.SetText("F");
	wordFt = wordsel.GetFont();
	wordFt.SetBold(bBold);
	wordFt.SetItalic(bItalic);
	wordFt.SetUnderline(bUnderLine);
	wordsel.SetFont(wordFt);
}
 
//���õ�Ԫ��(nRow, nColumn)������ʽ�������С��������ɫ������ɫ
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
 
//���ô�(beginRow, beginColumn)��(endRow, endColumn)��Χ����Ԫ���������ʽ�������С��������ɫ������ɫ
void WordHandle::SetTableFont(int beginRow, int beginColumn, int endRow, int endColumn, LPCTSTR szFontName, float fSize, long lFontColor, long lBackColor){
	wordtable.SetBottomPadding(10.0);
	int nRow, nColumn;
	for(nRow = beginRow; nRow <= endRow; nRow++){
		for(nColumn = beginColumn; nColumn <= endColumn; nColumn++){
			SetTableFont(nRow, nColumn, szFontName, fSize, lFontColor, lBackColor);
		}
	}
}
 
//���õ�Ԫ��(nRow, nColumn)���塢б�塢�»���
void WordHandle::SetTableFont(int nRow, int nColumn, BOOL bBold, BOOL bItalic, BOOL bUnderLine){
	cell = wordtable.Cell(nRow, nColumn);
	cell.Select();
	wordFt = wordsel.GetFont();
	wordFt.SetBold(bBold);
	wordFt.SetItalic(bItalic);
	wordFt.SetUnderline(bUnderLine);
	wordsel.SetFont(wordFt);
}
 
//���ñ��Ԫ��(nRow, nColumn)��ˮƽ����ֱ���뷽ʽ
void WordHandle::SetTableFormat(int nRow, int nColumn, long horizontalAlignment, long verticalAlignment){
	cell = wordtable.Cell(nRow, nColumn);
	cell.Select();
	cell.SetVerticalAlignment(verticalAlignment);
	paragraphformat = wordsel.GetParagraphFormat();
	paragraphformat.SetAlignment(horizontalAlignment);
	wordsel.SetParagraphFormat(paragraphformat);
}
 
//���ñ���(beginRow, beginColumn)��(endRow, endRow)������ˮƽ����ֱ���뷽ʽ
void WordHandle::SetTableFormat(int beginRow, int beginColumn, int endRow, int endColumn, long horizontalAlignment, long verticalAlignment){
	int nRow, nColumn;
	for( nRow = beginRow; nRow <= endRow; nRow++){
		for( nColumn = beginColumn; nColumn <= endColumn; nColumn++){
			SetTableFormat(nRow, nColumn, horizontalAlignment, verticalAlignment);
		}
	}
}
 
//���ñ��Ԫ������������ڱ߾����
void WordHandle::SetTablePadding(float topPadding, float bottomPadding, float leftPadding, float rightPadding){
	wordtable.SetTopPadding(topPadding);
	wordtable.SetBottomPadding(bottomPadding);
	wordtable.SetLeftPadding(leftPadding);
	wordtable.SetRightPadding(rightPadding);
}
 
//�ϲ���Ԫ��
BOOL WordHandle::MergeCell(int  cell1row, int cell1col, int cell2row, int cell2col){
	Cell cell1=wordtable.Cell(cell1row,cell1col);
	Cell cell2=wordtable.Cell(cell2row,cell2col);
	cell1.Merge(cell2);
 
	return TRUE;
}
 
//����ͼ��
void WordHandle::CreateChart(){
	ishaps=	wordsel.GetInlineShapes();
	CComVariant classt;
	classt=CComVariant("MSGraph.Chart.8");
	ishap=ishaps.AddOLEObject(&classt,vOptional,vFalse,vFalse,vOptional,vOptional,vOptional,vOptional);
	AddParagraph();
}
 
//���ͼƬ
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
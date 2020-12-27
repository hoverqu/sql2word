// WordHandle.h: interface for the WordHandle class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_WORDHANDLE_H__7F5096E5_F09B_4DA1_9FC2_28FA2D2C1573__INCLUDED_)
#define AFX_WORDHANDLE_H__7F5096E5_F09B_4DA1_9FC2_28FA2D2C1573__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


//段落对齐方式
enum wdAlignParagraphAlignment {
	    wdAlignParagraphLeft = 0, //左对齐
		wdAlignParagraphCenter = 1, //居中
		wdAlignParagraphRight = 2, //右对齐
		wdAlignParagraphJustify = 3 //两端对齐
}; 

//单元格垂直对齐方式
enum WdCellVerticalAlignment{
	    wdCellAlignVerticalTop = 0, //文字与单元格上框线对齐
		wdCellAlignVerticalCenter = 1, //文字与单元格中心对齐
		wdCellAlignVerticalBottom = 3 //文字与单元格底边框线对齐
};

#include "msword.h"
#include  <atlbase.h>
#include "comdef.h"
#include <afxdisp.h> //COlVirant 


class WordHandle  
{
private:
	_Application wordapp;
	COleVariant vTrue, vFalse, vOptional;
	Documents worddocs;
	CComVariant tpl, Visble, DocType, NewTemplate;
	Selection wordsel;
	_Document worddoc;
	Tables wordtables;
	Table wordtable;
	_ParagraphFormat paragraphformat;
	Cell cell;
	Cells cells;
	_Font wordFt;
	Range range;
	CComVariant SaveChanges,OriginalFormat,RouteDocument;
	CComVariant ComVariantTrue,ComVariantFalse;
	InlineShapes ishaps;
	InlineShape ishap;
	
	
public:
	void AddPicture(LPCTSTR picname);
	void AddTime(CTime time);
	void CreateChart();
	BOOL MergeCell(int  cell1row,int cell1col,int  cell2row,int cell2col);
	void CloseWordSave(LPCTSTR wordname);  
	void CloseWord();
	void SetTableFormat(int nRow, int nColumn, long horizontalAlignment, long verticalAlignment);
	void SetTableFormat(int beginRow, int beginColumn, int endRow, int endColumn, long horizontalAlignment = wdAlignParagraphCenter, long verticalAlignment = 
		wdCellAlignVerticalCenter);
	void SetTablePadding(float topPadding = 4.0, float bottomPadding = 4.0, float leftPadding = 4.0, float rightPadding = 4.0);
	void SetTableFont(int nRow, int nColumn, BOOL bBold, BOOL bItalic = FALSE, BOOL bUnderLine = FALSE);
	void SetTableFont(int nRow, int nColumn, LPCTSTR szFontName = "宋体", float fSize=9, long lFontColor=0, long lBackColor = 0);
	void SetTableFont(int beginRow, int beginColumn, int endRow, int endColumn, LPCTSTR szFontName = "宋体", float fSize=9, long lFontColor=0, long lBackColor = 
		
		0);
	void SetFont(BOOL bBold, BOOL bItalic = FALSE, BOOL bUnderLine = FALSE );
	void SetFont(LPCTSTR szFontName ,float fSize = 9, long lFontColor = 0, long lBackColor=0);
	void AddParagraph();
	void WriteText(LPCTSTR szText);
	BOOL CreateWord();
	void ShowWord(BOOL show);
	void WriteCell(int nRow, int nColumne, LPCTSTR szText);
	void WriteCell(int nRow, int nColumne, LPCTSTR lpszFormat, float num);
	BOOL CreateTable(int nRow, int nColumn);
	void SetParagraphFormat(int nAlignment);
	BOOL CreateDocumtent();
	BOOL CreateApp();
	WordHandle();
	virtual ~WordHandle();
	
};

#endif // !defined(AFX_WORDHANDLE_H__7F5096E5_F09B_4DA1_9FC2_28FA2D2C1573__INCLUDED_)
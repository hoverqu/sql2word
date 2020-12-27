// SQLField.h: interface for the WordHandle class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_SQLFIELD_H__7F5096E5_F09B_4DA1_9FC2_28FA2D2C1573__INCLUDED_)
#define AFX_SQLFIELD_H__7F5096E5_F09B_4DA1_9FC2_28FA2D2C1573__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


class SQLField  
{
   //SQL◊÷∂Œ Ù–‘
private:
   CString FieldName;
   CString FieldType;
   CString FieldComment;
   
public:
	void SetFieldName(CString Fd);
	void SetFieldComment(CString Fc);
	void SetFieldType(CString Ft);
	
 

}

#endif // !defined(AFX_SQLFIELD_H__7F5096E5_F09B_4DA1_9FC2_28FA2D2C1573__INCLUDED_)
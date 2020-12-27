// SQLTable.h: interface for the WordHandle class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_SQLTABLE_H__7F5096E5_F09B_4DA1_9FC2_28FA2D2C1573__INCLUDED_)
#define AFX_SQLTABLE_H__7F5096E5_F09B_4DA1_9FC2_28FA2D2C1573__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


class SQLTable  
{
   //SQL±Ì Ù–‘
private:
   CString TableName;
   int FieldCount;
   Fields   FieldList[];
public:
	void SetTableName();
	void SetFieldCount(int Count);
    int GetFieldCount();
	void AddField();
 

}

#endif // !defined(AFX_SQLTABLE_H__7F5096E5_F09B_4DA1_9FC2_28FA2D2C1573__INCLUDED_)
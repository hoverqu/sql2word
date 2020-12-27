/////sqlstruct.h
#ifndef SQLSTRUCT
#define SQLSTRUCT
#define TABLE_NAMELEN   64
#define TABLE_INDEXSLEN   256
#define TABLE_ENGINELEN  32
#define TABLE_CHARSETLEN  32
#define TABLE_FIELD_NAMELEN   64
#define TABLE_FIELD_NAMELEN   64
#define TABLE_FIELD_DEFAULTLEN   64
#define TABLE_FIELD_COMMENTLEN   128
#define TABLE_FIELD_TYPELEN   32
typedef struct SqlField  
{
	 char FieldName[TABLE_FIELD_NAMELEN];//字段名
     char FieldType[TABLE_FIELD_TYPELEN];//字段类型
     int  FieldLenth;//字段长度
     int  FieldDecimalsLenth;//字段小数长度
     int  IsNull;
	 int  Iskey;
	 char FieldDefault[TABLE_FIELD_DEFAULTLEN];//字段缺省
	 char FieldComment[TABLE_FIELD_COMMENTLEN];//字段注释
     struct SqlField  *NextField;
}SQLField;
typedef struct SqlTable  
{
	char TableName[TABLE_NAMELEN];
    char TableEngine[TABLE_ENGINELEN];
    char TableCharSet[TABLE_CHARSETLEN];
    char TableKey[TABLE_NAMELEN];
    char TableIndexs[TABLE_INDEXSLEN];
	int  FieldCount;
    struct SqlField *TableField;  //指向指针结构数组
}SQLTable;
//数据库
typedef struct SqlDataBase  
{	
	struct SqlTable Table;
    struct SqlDataBase *NextTable;

}DataBase;

#endif
/////
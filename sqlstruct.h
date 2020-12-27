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
	 char FieldName[TABLE_FIELD_NAMELEN];//�ֶ���
     char FieldType[TABLE_FIELD_TYPELEN];//�ֶ�����
     int  FieldLenth;//�ֶγ���
     int  FieldDecimalsLenth;//�ֶ�С������
     int  IsNull;
	 int  Iskey;
	 char FieldDefault[TABLE_FIELD_DEFAULTLEN];//�ֶ�ȱʡ
	 char FieldComment[TABLE_FIELD_COMMENTLEN];//�ֶ�ע��
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
    struct SqlField *TableField;  //ָ��ָ��ṹ����
}SQLTable;
//���ݿ�
typedef struct SqlDataBase  
{	
	struct SqlTable Table;
    struct SqlDataBase *NextTable;

}DataBase;

#endif
/////
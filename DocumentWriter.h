/*
  ������ ��� �������� ������ ������ � ��������� Word, Excel

  ��� ������ � Word
  ExportToWordTemplate(params, ds1, ds2);
  params - ��������� ��� �������� � Word, � ��� ����� ���� � �������, ���� � ����������,
           ����� ����� �������� ������ ds1 � ds2
  ds1 - ��� ������������� ����� ������ ��� ����� MERGEFIELD � DOCVARIABLE � ��������� �����.
  ds2 - ��� ������������ (������� ������ ������ ������)
        ��� ������������� (���� ������� ����� � params) ����� ������,
        ������������ ��� ����� DOCVARIABLE
  ������� ������:
  //1. ���������� �������� ����� MERGEFIELD � DOCVARIABLE � ��������� �����
  ����������:
  1. ����� ����� � ds1 � ds2 �� ������ ���������, � ��������� ������ ��� ������ ��������� ����� �� ds1

*/

#ifndef DOCUMENTWRITER_H
#define DOCUMENTWRITER_H

#include "Ora.hpp"
#include "OraDataTypeMap.hpp"
//#include "odacutils.h"
//#include "math.h"
#include "WordUtil.h"
#include "ExcelUtil.h"
#include "DbfUtil.h"
#include <vector>
#include "taskutils.h"


/* ����� ��� �������� ����������� ������ ��������*/
class TDocumentWriterResult
{
public:
    std::vector<String> resultFiles;
    void __fastcall addResultFile(String filename);
    void __fastcall appendResultFiles(std::vector<String> filenames);
    void __fastcall clear();
    int __fastcall resultFileCount();
};

/* Dbf */
class TDbaseExportParams {
public:
    String resultFilename;
    TOraQuery *srcDataSet;

    String id;          // �������������
    //String label;   // �� ������������
    //bool fDefault;  // �� ������������
    bool fDisableUnassignedFields;    // ��������� ��������������� ���� (�� ������������ required �� ������ �����?)
    TDbfFieldList Fields;    // ������ ����� ��� �������� � ���� DBF
};


/* Excel parameters */
// ��������� ��� �������� ���������� ���� (�������) MS Excel
class TExcelField    // ��� �������� ������� ����� � Excel
{
public:
    TExcelField();
    ~TExcelField();
    String format;      // ������ ������ � Excel
    //String name;        // ��� ����
    String descr;
    String fieldName; // ��������. ����� �������� �� name
    bool visible;

    //int title_rows;       // ������ ��������� � �������
    int width;              // ������ �������
    int bwraptext_head;     // ���� �������� �� ������ � ����� �������
    int bwraptext_body;     // ���� �������� �� ����� � ���� �������
};

TExcelField::TExcelField() :
    format("@"),
    fieldName(""),
    descr(""),
    width(-1),
    bwraptext_head(-1),
    bwraptext_body(-1),
    visible(true)
{
}
TExcelField::~TExcelField()
{
}



/**/
class TExcelTableDataSet
{
public:
    TExcelTableDataSet(TDataSet* dataSet_, String tableName_, String fieldNamePrefix_):
        dataSet(dataSet_),
        tableName(tableName_),
        fieldNamePrefix(fieldNamePrefix_)

    {};
    String tableName;
    TDataSet* dataSet;
    String fieldNamePrefix;
};

class TExcelSingleDataSet
{
public:
    TExcelSingleDataSet(TDataSet* dataSet_, String fieldNamePrefix_):
        dataSet(dataSet_),
        fieldNamePrefix(fieldNamePrefix_)

    {};
    TDataSet* dataSet;
    String fieldNamePrefix;
};

class TExcelVtArray
{
public:
    TExcelVtArray(Variant vtArray_, String tableName_):
        vtArray(vtArray_),
        tableName(tableName_)
    {};
    Variant vtArray;
    String tableName;
};


typedef std::vector<TExcelField> TExcelFieldList;

// ��������� ��� �������� ���������� �������� � MS Excel
class TExcelExportParams  {
public:
    String id;
    String label;
    //bool fDefault;
    Variant templateDocument;       // ������ ���������
    String templateFilename;        // ��� ����� ������� Excel  (���� �� ����� templateDocument)
    String resultFilename;

    // ��� ��������������� �������� �������
    String title_label;             // ������ - ��������� � �������� ��������� � ������ Excel (��������� � ��������� ���������)
    int title_height;               // ������ ��������� � �������  (��������� � ��������� ���������)
    TExcelFieldList Fields; // ������ ����� ��� �������� � ���� MS Excel
    String table_range_name;        // ��� ��������� ��������� ����� (��� ������ � ������)
    bool fUnbounded;                // ���� ����, ��� �������� table_range_name ����� ��������, � ������������ � ����������� ������� � ��������� ������
    String link_field_left;
    String link_field_right;


    // ��� ������ � ���������
    void addTableDataSet(TDataSet* dataSet, const String& tableName, const String& fieldNamePrefix = "");
    void addSingleDataSet(TDataSet* dataSet, const String& fieldNamePrefix = "");
    std::vector<TExcelTableDataSet> tableDs;   // ��������� ������ ��� ���������� ������
    std::vector<TExcelSingleDataSet> singleDs;   // ��������� ������ ��� ���������� ��������� �����

    void addTableVtArray(Variant vtArray, const String& tableName);
    std::vector<TExcelVtArray> tableVtArray;   // ��������� ������ ��� ���������� ��������� �����

    Variant findTableVtArray(const String& tableName);
};

/* ��������� DataSet � ������ ��� ������ */
void TExcelExportParams::addTableDataSet(TDataSet* dataSet, const String& tableName, const String& fieldNamePrefix)
{
    tableDs.push_back(TExcelTableDataSet(dataSet, tableName, fieldNamePrefix));
}

/* ��������� DataSet � ������ ��� ��������� �����*/
void TExcelExportParams::addSingleDataSet(TDataSet* dataSet, const String& fieldNamePrefix)
{
    singleDs.push_back( TExcelSingleDataSet(dataSet, fieldNamePrefix) );
}

/* ��������� ������ Variant */
void TExcelExportParams::addTableVtArray(Variant vtArray, const String& tableName)
{
    tableVtArray.push_back( TExcelVtArray(vtArray, tableName) );
}

/* ������� ������� �� ����� */
Variant TExcelExportParams::findTableVtArray(const String& tableName)
{
    for (std::vector<TExcelVtArray>::iterator it = tableVtArray.begin(); it != tableVtArray.end(); it++ )
    {
        if (it->tableName == tableName)
        {
            return it->vtArray;
        }
    }
    return Variant();
}



/* Word */

typedef TDataSet* TWordMergeDataSet;

class TWordSingleDataSet
{
public:
    TWordSingleDataSet(TDataSet* dataSet_, String fieldNamePrefix_):
        dataSet(dataSet_),
        fieldNamePrefix(fieldNamePrefix_)

    {};
    TDataSet* dataSet;
    String fieldNamePrefix;

};


class TWordTableDataSet
{
public:
    TWordTableDataSet(TDataSet* dataSet_, int tableIndex_, String fieldNamePrefix_):
        dataSet(dataSet_),
        tableIndex(tableIndex_),
        fieldNamePrefix(fieldNamePrefix_)

    {};
    int tableIndex;
    TDataSet* dataSet;
    String fieldNamePrefix;
};

// ��������� ��� �������� ���������� �������� � MS Word
class TWordExportParams
{
public:
    String id;          // ������������� ����������
    AnsiString label;
    //bool fDefault;
    //AnsiString template_name;   // ��� ����� ������� MS Word
    //int page_per_doc;           // ���������� ������� �� �������� MS Word
    //AnsiString filter_main_field;      // ��� ���� �� ��������� ������� ��� ��������� �� ��������� ���� word_filter_sec_field
    //AnsiString filter_sec_field;       // ��� ���� �� ���������������� ������� (��. word_filter_main_field)
    //AnsiString filter_infix_sec_field; // ��� ���� �� ���������������� �������, �������� �������� ����� ������������ � ����� ��������������� �����





    String templateFilename;   // ������ ��� ����� ������� MS Word
    String resultFilename;   //
    String imageFilesDirectory;   // ������� � ������� ������������, ����������� � ���� [img]
    int pagePerDocument;           // ���������� ������� �� �������� MS Word

    /* DataSets links*/
    String filter_main_field;      // ��� ���� �� ��������� ������� ��� ��������� �� ��������� ���� word_filter_sec_field
    String filter_sec_field;       // ��� ���� �� ���������������� ������� (��. word_filter_main_field)
    //String filter_infix_sec_field; // ��� ���� �� ���������������� �������, �������� �������� ����� ������������ � ����� ��������������� �����

    //int tableIndex
    std::vector<TWordMergeDataSet> mergeDs;   // ��������� ������ ��� �������
    std::vector<TWordTableDataSet> tableDs;   // ��������� ������ ��� ���������� ������
    std::vector<TWordSingleDataSet> singleTextDs;   // ��������� ������ ��� ���������� ����� DOCVariable (�������)
    std::vector<TWordSingleDataSet> singleImageDs;  // ��������� ������ ��� ���������� ����� DOCVariable (������������)
    std::vector<TWordSingleDataSet> formtextDs;  // ��������� ������ ��� ���������� ����� FORMTEXT

    void addMergeDataSet(TDataSet* dataSet);
    void addTableDataSet(TDataSet* dataSet, int tableIndex, const String& fieldNamePrefix = "");
    void addSingleTextDataSet(TDataSet* dataSet, const String& fieldNamePrefix = "");
    void addSingleImageDataSet(TDataSet* dataSet, const String& fieldNamePrefix = "");
    void addFormtextDataSet(TDataSet* dataSet, const String& fieldNamePrefix = "");
};

void TWordExportParams::addMergeDataSet(TDataSet* dataSet)
{
    mergeDs.push_back(dataSet);
}

void TWordExportParams::addTableDataSet(TDataSet* dataSet, int tableIndex, const String& fieldNamePrefix)
{
    tableDs.push_back(TWordTableDataSet(dataSet, tableIndex, fieldNamePrefix));
}

void TWordExportParams::addSingleTextDataSet(TDataSet* dataSet, const String& fieldNamePrefix)
{
    singleTextDs.push_back( TWordSingleDataSet(dataSet, fieldNamePrefix) );
}

void TWordExportParams::addSingleImageDataSet(TDataSet* dataSet, const String& fieldNamePrefix)
{
    singleImageDs.push_back( TWordSingleDataSet(dataSet, fieldNamePrefix) );
}

void TWordExportParams::addFormtextDataSet(TDataSet* dataSet, const String& fieldNamePrefix)
{
    formtextDs.push_back( TWordSingleDataSet(dataSet, fieldNamePrefix) );
}



class TDocumentWriter
{
private:

public:
    TDocumentWriterResult _result;

    void __fastcall ExportToWordTemplate_old(const TWordExportParams* wordExportParams, TDataSet *QueryMerge, TDataSet *QueryFormFields);  // ���������� ������ Word �� ���� �������
    void __fastcall ExportToExcelTemplate(TExcelExportParams* excelExportParams);

    void __fastcall ExportToWordTemplate(TWordExportParams* wordExportParams);  // ���������� ������ Word �� ���� �������

    void __fastcall ExportToExcel(TExcelExportParams* excelExportParams);
    void __fastcall ExportToDbf(TDbaseExportParams* dbaseExportParams);

    Variant __fastcall CreateExcelTemplate(TExcelExportParams* excelExportParams);  

    //void __fastcall ExportToExcel(TOraQuery *OraQuery); // ���������� ������ Excel
    //void __fastcall ExportToExcelTemplate(TOraQuery *QueryTable, TOraQuery *QueryFields);
    //void __fastcall ExportToDBF(TOraQuery *OraQuery);   // ���������� DBF-�����
};

#endif


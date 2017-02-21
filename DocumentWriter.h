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
#include "MSWordWorks.h"
#include "MSExcelWorks.h"
#include <vector>


class TDocumentWriterResult
{
public:
    std::vector<String> resultFiles;
    void __fastcall addResultFile(String filename);
    void __fastcall appendResultFiles(std::vector<String> filenames);
    void __fastcall clear();
};


// ��������� ��� �������� ���������� ���� (�������) DBASE
typedef struct {    // ��� �������� ��������� dbf-�����
    String type;    // ��� fieldtype is a single character [C,D,F,L,M,N]
    String name;    // ��� ���� (�� 10 ��������).
    int length;         // ����� ����
    int decimals;       // ����� ���������� �����
    // Character 	 1-255
    // Date	  8
    // Logical	  1
    // Memo	  10
    // Numeric	1-30
    // Decimals is 0 for non-numeric, number of decimals for numeric.
} DBASEFIELD;


/* Excel parameters */
// ��������� ��� �������� ���������� ���� (�������) MS Excel
typedef struct {    // ��� �������� ������� ����� � Excel
    AnsiString format;      // ������ ������ � Excel
    AnsiString name;        // ��� ����
    //int title_rows;       // ������ ��������� � �������
    int width;              // ������ �������
    int bwraptext;          // ���� ������� �� ������
} EXCELFIELD;


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



// ��������� ��� �������� ���������� �������� � MS Excel
class TExcelExportParams  {
public:
    String id;
    String label;
    //bool fDefault;
    String templateFilename;       // ��� ����� ������� Excel
    String resultFilename;
    AnsiString title_label;         // ������ - ��������� � �������� ��������� � ������ Excel (��������� � ��������� ���������)

    int title_height;               // ������ ��������� � �������  (��������� � ��������� ���������)
    std::vector<EXCELFIELD> Fields;     // ������ ����� ��� �������� � ���� MS Excel
    String table_range_name;        // ��� ��������� ��������� ����� (��� ������ � ������)
    bool fUnbounded;                    // ���� ����, ��� �������� table_range_name ����� ��������, � ������������ � ����������� ������� � ��������� ������

    void addTableDataSet(TDataSet* dataSet, const String& tableName, const String& fieldNamePrefix = "");
    void addSingleDataSet(TDataSet* dataSet, const String& fieldNamePrefix = "");
    std::vector<TExcelTableDataSet> tableDs;   // ��������� ������ ��� ���������� ������
    std::vector<TExcelSingleDataSet> singleDs;   // ��������� ������ ��� ���������� ��������� �����
};

void TExcelExportParams::addTableDataSet(TDataSet* dataSet, const String& tableName, const String& fieldNamePrefix)
{
    tableDs.push_back(TExcelTableDataSet(dataSet, tableName, fieldNamePrefix));
}

void TExcelExportParams::addSingleDataSet(TDataSet* dataSet, const String& fieldNamePrefix)
{
    singleDs.push_back( TExcelSingleDataSet(dataSet, fieldNamePrefix) );
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

    void addMergeDataSet(TDataSet* dataSet);
    void addTableDataSet(TDataSet* dataSet, int tableIndex, const String& fieldNamePrefix = "");
    void addSingleTextDataSet(TDataSet* dataSet, const String& fieldNamePrefix = "");
    void addSingleImageDataSet(TDataSet* dataSet, const String& fieldNamePrefix = "");
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



// ��������� ��� �������� ���������� �������� � DBF
typedef struct {    // ��� �������� ������� ����� � Excel
    String id;
    String label;
    //bool fDefault;
    bool fAllowUnassignedFields;
    std::vector<DBASEFIELD> Fields;    // ������ ����� ��� ������� � ���� DBF
} EXPORT_PARAMS_DBASE;


class TDocumentWriter
{
private:

public:
    TDocumentWriterResult _result;

    void __fastcall ExportToWordTemplate_old(const TWordExportParams* wordExportParams, TDataSet *QueryMerge, TDataSet *QueryFormFields);  // ���������� ������ Word �� ���� �������
    void __fastcall ExportToExcelTemplate(TExcelExportParams* excelExportParams);

    void __fastcall ExportToWordTemplate(TWordExportParams* wordExportParams);  // ���������� ������ Word �� ���� �������

    //void __fastcall ExportToExcel(TOraQuery *OraQuery); // ���������� ������ Excel
    //void __fastcall ExportToExcelTemplate(TOraQuery *QueryTable, TOraQuery *QueryFields);
    //void __fastcall ExportToDBF(TOraQuery *OraQuery);   // ���������� DBF-�����
};

#endif


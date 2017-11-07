/*
  Модуль для удобства вывода данных в документы Word, Excel

  При выводе в Word
  ExportToWordTemplate(params, ds1, ds2);
  params - параметры для экспорта в Word, в том числе путь к шаблону, путь к результату,
           связь между наборами данных ds1 и ds2
  ds1 - это многострочный набор данных для полей MERGEFIELD и DOCVARIABLE в табличной части.
  ds2 - это однострочный (берется только первая строка)
        или многострочный (если указана связь в params) набор данных,
        используется для полей DOCVARIABLE
  Принцип работы:
  //1. Выясняется набличие полей MERGEFIELD и DOCVARIABLE в табличной части
  Примечания:
  1. Имена полей в ds1 и ds2 не должны совпадать, в противном случае при замене приоритет будет за ds1

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


/* Класс для хранения результатов работы объектов*/
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

    String id;          // Идентификатор
    //String label;   // Не используется
    //bool fDefault;  // Не используется
    bool fDisableUnassignedFields;    // Разрешить неприкрепленные поля (мб использовать required из списка полей?)
    TDbfFieldList Fields;    // Список полей для экспорта в файл DBF
};


/* Excel parameters */
// Структура для хранения параметров поля (столбца) MS Excel
class TExcelField    // Для описания формата ячеек в Excel
{
public:
    TExcelField();
    ~TExcelField();
    String format;      // Формат ячейки в Excel
    //String name;        // Имя поля
    String descr;
    String fieldName; // Временно. потом поменять на name
    bool visible;

    //int title_rows;       // Высота заголовка в строках
    int width;              // Ширина столбца
    int bwraptext_head;     // Флаг переноса по словам в шапке таблицы
    int bwraptext_body;     // Флаг переноса по слова в теле таблицы
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

// Структура для хранения параметров экспорта в MS Excel
class TExcelExportParams  {
public:
    String id;
    String label;
    //bool fDefault;
    Variant templateDocument;       // Шаблон документа
    String templateFilename;        // Имя файла шаблона Excel  (если не задан templateDocument)
    String resultFilename;

    // Для автоматического создания шаблона
    String title_label;             // Строка - выводимая в качестве заголовка в отчете Excel (перенести в отдельную структуру)
    int title_height;               // Высота заголовка в строках  (перенести в отдельную структуру)
    TExcelFieldList Fields; // Список полей для экспорта в файл MS Excel
    String table_range_name;        // Имя диапазона табличной части (при выводе в шаблон)
    bool fUnbounded;                // Флаг того, что диапазон table_range_name будет увеличен, в соответствии с количеством записей в источнике данных
    String link_field_left;
    String link_field_right;


    // Для работы с шаблонами
    void addTableDataSet(TDataSet* dataSet, const String& tableName, const String& fieldNamePrefix = "");
    void addSingleDataSet(TDataSet* dataSet, const String& fieldNamePrefix = "");
    std::vector<TExcelTableDataSet> tableDs;   // Источники данных для заполнения таблиц
    std::vector<TExcelSingleDataSet> singleDs;   // Источники данных для заполнения отдельных полей

    void addTableVtArray(Variant vtArray, const String& tableName);
    std::vector<TExcelVtArray> tableVtArray;   // Источники данных для заполнения отдельных полей

    Variant findTableVtArray(const String& tableName);
};

/* Добавляет DataSet в список для Таблиц */
void TExcelExportParams::addTableDataSet(TDataSet* dataSet, const String& tableName, const String& fieldNamePrefix)
{
    tableDs.push_back(TExcelTableDataSet(dataSet, tableName, fieldNamePrefix));
}

/* Добавляет DataSet в список для отдельных полей*/
void TExcelExportParams::addSingleDataSet(TDataSet* dataSet, const String& fieldNamePrefix)
{
    singleDs.push_back( TExcelSingleDataSet(dataSet, fieldNamePrefix) );
}

/* Добавляет массив Variant */
void TExcelExportParams::addTableVtArray(Variant vtArray, const String& tableName)
{
    tableVtArray.push_back( TExcelVtArray(vtArray, tableName) );
}

/* Находит таблицу по имени */
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

// Структура для хранения параметров экспорта в MS Word
class TWordExportParams
{
public:
    String id;          // Идентификатор параметров
    AnsiString label;
    //bool fDefault;
    //AnsiString template_name;   // Имя файла шаблона MS Word
    //int page_per_doc;           // Количество страниц на документ MS Word
    //AnsiString filter_main_field;      // Имя поля из основного запроса для сравнения со значением поля word_filter_sec_field
    //AnsiString filter_sec_field;       // Имя поля из вспомогательного запроса (см. word_filter_main_field)
    //AnsiString filter_infix_sec_field; // Имя поля из вспомогательного запроса, значение которого будет присоединено к имени результирующего файла





    String templateFilename;   // Полное имя файла шаблона MS Word
    String resultFilename;   //
    String imageFilesDirectory;   // Каталог с файлами изображаений, вставляемых в поля [img]
    int pagePerDocument;           // Количество страниц на документ MS Word

    /* DataSets links*/
    String filter_main_field;      // Имя поля из основного запроса для сравнения со значением поля word_filter_sec_field
    String filter_sec_field;       // Имя поля из вспомогательного запроса (см. word_filter_main_field)
    //String filter_infix_sec_field; // Имя поля из вспомогательного запроса, значение которого будет присоединено к имени результирующего файла

    //int tableIndex
    std::vector<TWordMergeDataSet> mergeDs;   // Источники данных для слияния
    std::vector<TWordTableDataSet> tableDs;   // Источники данных для заполнения таблиц
    std::vector<TWordSingleDataSet> singleTextDs;   // Источники данных для заполнения полей DOCVariable (текстом)
    std::vector<TWordSingleDataSet> singleImageDs;  // Источники данных для заполнения полей DOCVariable (изображением)
    std::vector<TWordSingleDataSet> formtextDs;  // Источники данных для заполнения полей FORMTEXT

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

    void __fastcall ExportToWordTemplate_old(const TWordExportParams* wordExportParams, TDataSet *QueryMerge, TDataSet *QueryFormFields);  // Заполнение отчета Word на базе шаблона
    void __fastcall ExportToExcelTemplate(TExcelExportParams* excelExportParams);

    void __fastcall ExportToWordTemplate(TWordExportParams* wordExportParams);  // Заполнение отчета Word на базе шаблона

    void __fastcall ExportToExcel(TExcelExportParams* excelExportParams);
    void __fastcall ExportToDbf(TDbaseExportParams* dbaseExportParams);

    Variant __fastcall CreateExcelTemplate(TExcelExportParams* excelExportParams);  

    //void __fastcall ExportToExcel(TOraQuery *OraQuery); // Заполнение отчета Excel
    //void __fastcall ExportToExcelTemplate(TOraQuery *QueryTable, TOraQuery *QueryFields);
    //void __fastcall ExportToDBF(TOraQuery *OraQuery);   // Заполнение DBF-файла
};

#endif


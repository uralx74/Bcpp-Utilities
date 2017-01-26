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


// Структура для хранения параметров поля (столбца) DBASE
typedef struct {    // Для описания структуры dbf-файла
    String type;    // Тип fieldtype is a single character [C,D,F,L,M,N]
    String name;    // Имя поля (до 10 символов).
    int length;         // Длина поля
    int decimals;       // Длина десятичной части
    // Character 	 1-255
    // Date	  8
    // Logical	  1
    // Memo	  10
    // Numeric	1-30
    // Decimals is 0 for non-numeric, number of decimals for numeric.
} DBASEFIELD;


/* Excel parameters */
// Структура для хранения параметров поля (столбца) MS Excel
typedef struct {    // Для описания формата ячеек в Excel
    AnsiString format;      // Формат ячейки в Excel
    AnsiString name;        // Имя поля
    //int title_rows;       // Высота заголовка в строках
    int width;              // Ширина столбца
    int bwraptext;          // Флаг перенос по словам
} EXCELFIELD;

// Структура для хранения параметров экспорта в MS Excel
class TExcelExportParams  {
public:
    String id;
    String label;
    //bool fDefault;
    String templateFilename;       // Имя файла шаблона Excel
    String resultFilename;
    AnsiString title_label;         // Строка - выводимая в качестве заголовка в отчете Excel (перенести в отдельную структуру)
    int title_height;               // Высота заголовка в строках  (перенести в отдельную структуру)
    std::vector<EXCELFIELD> Fields;     // Список полей для экспорта в файл MS Excel
    String table_range_name;        // Имя диапазона табличной части (при выводе в шаблон)
    bool fUnbounded;                    // Флаг того, что диапазон table_range_name будет увеличен, в соответствии с количеством записей в источнике данных
};

/* Word */
// Структура для хранения параметров экспорта в MS Word
class TWordExportParams
{
public:
    String templateFilename;   // Полное имя файла шаблона MS Word
    String resultFilename;   // Полное имя файла шаблона MS Word
    String imageFilesDirectory;   // Каталог с файлами изображаений, вставляемых в поля [img]
    int pagePerDocument;           // Количество страниц на документ MS Word

    /* DataSets links*/
    String filter_main_field;      // Имя поля из основного запроса для сравнения со значением поля word_filter_sec_field
    String filter_sec_field;       // Имя поля из вспомогательного запроса (см. word_filter_main_field)
    //String filter_infix_sec_field; // Имя поля из вспомогательного запроса, значение которого будет присоединено к имени результирующего файла
};

// Структура для хранения параметров экспорта в DBF
typedef struct {    // Для описания формата ячеек в Excel
    String id;
    String label;
    //bool fDefault;
    bool fAllowUnassignedFields;
    std::vector<DBASEFIELD> Fields;    // Список полей для экспрта в файл DBF
} EXPORT_PARAMS_DBASE;


class TDocumentWriter
{
private:

public:
    TDocumentWriterResult result;

    void __fastcall ExportToWordTemplate(const TWordExportParams* wordExportParams, TDataSet *QueryMerge, TDataSet *QueryFormFields);  // Заполнение отчета Word на базе шаблона
    void __fastcall ExportToExcelTemplate(const TExcelExportParams* excelExportParams, TDataSet* QueryTable, TDataSet* QueryFields);

    //void __fastcall ExportToExcel(TOraQuery *OraQuery); // Заполнение отчета Excel
    //void __fastcall ExportToExcelTemplate(TOraQuery *QueryTable, TOraQuery *QueryFields);
    //void __fastcall ExportToDBF(TOraQuery *OraQuery);   // Заполнение DBF-файла
};

#endif


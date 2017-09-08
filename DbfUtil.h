//---------------------------------------------------------------------------
#ifndef DBFUTIL_H
#define DBFUTIL_H

/*******************************************************************************
	Класс для работы с Dbf
    Версия от 08.09.2017
    В.С.Овчинников

*******************************************************************************/

#include "system.hpp"
#include <utilcls.h>
#include "Comobj.hpp"
#include "sysutils.hpp"
#include "taskutils.h"
#include "Db.hpp"
#include "dbf.hpp"
#include "Dbf_Lang.hpp"


typedef std::pair<int, int> TLinkFields;


// Класс для хранения параметров поля (столбца) DBASE
class TDbfField
{    // Для описания структуры dbf-файла
public:
    char type;    // Тип fieldtype is a single character [C,D,F,L,M,N]
    String name;    // Имя поля (до 10 символов).
    int length;         // Длина поля
    int decimals;       // Длина десятичной части
    // Character 	 1-255
    // Date	  8
    // Logical	  1
    // Memo	  10
    // Numeric	1-30
    // Decimals is 0 for non-numeric, number of decimals for numeric.
};

/* Список полей */
typedef std::vector<TDbfField> TDbfFieldList;

/* Класс для работы с Dbf-файлами */
class TDbfUtil
{
private:

public:
    TDbf* __fastcall CreateDbf(String filename, TDbfFieldList* dbfFieldList);
    std::vector<TLinkFields> assignDataSet(TDataSet* srcDataSet, TDataSet* dstDataSet);

    void __fastcall WriteToDbf(TDataSet* srcTable, TDataSet* dstTable);

protected:

};

//---------------------------------------------------------------------------
#endif

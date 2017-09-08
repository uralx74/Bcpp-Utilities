//---------------------------------------------------------------------------
#ifndef DBFUTIL_H
#define DBFUTIL_H

/*******************************************************************************
	����� ��� ������ � Dbf
    ������ �� 08.09.2017
    �.�.����������

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


// ����� ��� �������� ���������� ���� (�������) DBASE
class TDbfField
{    // ��� �������� ��������� dbf-�����
public:
    char type;    // ��� fieldtype is a single character [C,D,F,L,M,N]
    String name;    // ��� ���� (�� 10 ��������).
    int length;         // ����� ����
    int decimals;       // ����� ���������� �����
    // Character 	 1-255
    // Date	  8
    // Logical	  1
    // Memo	  10
    // Numeric	1-30
    // Decimals is 0 for non-numeric, number of decimals for numeric.
};

/* ������ ����� */
typedef std::vector<TDbfField> TDbfFieldList;

/* ����� ��� ������ � Dbf-������� */
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

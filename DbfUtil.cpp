#include "DbfUtil.h"

using namespace strtools;


TDbf* __fastcall TDbfUtil::CreateDbf(String filename, TDbfFieldList* dbfFieldList)
{
    // ������� dbf-���� ����������
    TDbf* pTable = new TDbf(NULL);

    //pTableDst->TableLevel = 7; // required for AutoInc field
    pTable->TableLevel = 4;
    pTable->LanguageID = DbfLangId_RUS_866;

    pTable->TableName = ExtractFileName(filename);
    pTable->FilePathFull = ExtractFilePath(filename);


    // ������� ����������� ����� ������� �� ����������
    TDbfFieldDefs* TempFieldDefs = new TDbfFieldDefs(NULL);

    /*if (TempFieldDefs == NULL)
    {
        String msg = "Can't create storage.";
        throw Exception(msg);
    }*/

    for(TDbfFieldList::iterator it = dbfFieldList->begin(); it != dbfFieldList->end(); it++ )
    {
        TDbfFieldDef* TempFieldDef = TempFieldDefs->AddFieldDef();
        //TempFieldDef->FieldName = (*it).dbfFieldList->name;
        TempFieldDef->FieldName = it->name;
        //TempFieldDef->Required = true;
        //TempFieldDef->FieldType = Field->type;    // Use FieldType if Field->Type is TFieldType else use NativeFieldType

        TempFieldDef->NativeFieldType = it->type;
        if (TempFieldDef->NativeFieldType == '\0')
        {
            delete pTable;
            String msg = "��� ���� " + it->name + " ������ �� �����";
            throw Exception(msg);
        }

        TempFieldDef->Size = it->length;
        TempFieldDef->Precision = it->decimals;
    }

    // ?
    if (TempFieldDefs->Count == 0)
    {
        delete pTable;
        String msg = "�� ������� ��������� �������� �����.";
        throw Exception(msg);
    }

    pTable->CreateTableEx(TempFieldDefs);
    delete TempFieldDefs;

    return pTable;


}

/* ���������� ������ ��� �������� ����� �� dataSet � table
*/
std::vector<TLinkFields> TDbfUtil::assignDataSet(TDataSet* srcDataSet, TDataSet* dstDataSet)
{
    std::vector<TLinkFields> result;
    result.reserve(dstDataSet->FieldCount);

    int n = dstDataSet->FieldCount;

    // ���� �� ����� � ���������
    for (int i = 1; i <= n; i++ )
    {
        TField* fieldOfDataSet = srcDataSet->Fields->FindField(dstDataSet->Fields->FieldByNumber(i)->DisplayName);
        if ( fieldOfDataSet != NULL)
        {
            result.push_back(std::make_pair(fieldOfDataSet->FieldNo, i));
        }
    }

    return result;
}


/* ����������� ������ �� ������ DataSet � ������ */
void __fastcall TDbfUtil::WriteToDbf(TDataSet* srcTable, TDataSet* dstTable)
{
    // ������������ ����
    std::vector<TLinkFields> links = assignDataSet(srcTable, dstTable);

    // ������ ������ � �������
    while ( !srcTable->Eof )
    {
        dstTable->Append();

        for (std::vector<TLinkFields>::iterator it = links.begin(); it != links.end(); it++)
        {
            dstTable->Fields->FieldByNumber(it->second)->Value = srcTable->Fields->FieldByNumber(it->first)->Value;
        }
        srcTable->Next();  // ��������� � ��������� ������
    }
    dstTable->Post();
}

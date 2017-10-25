#include "DocumentWriter.h"


/* */
void __fastcall TDocumentWriterResult::clear()
{
    resultFiles.clear();
}

void __fastcall TDocumentWriterResult::addResultFile(String filename)
{
    resultFiles.push_back(filename);
}

void __fastcall TDocumentWriterResult::appendResultFiles(std::vector<String> filenames)
{
    resultFiles.insert(resultFiles.end(), filenames.begin(), filenames.end());
}


/*
 ���������� ������� MS Word
 QueryMerge - �������� ������, ������������ � �������� ��������� ������ ��� �������
 QueryFormFields - ��������������� ������, ������������ � �������� ��������� ������
 ��� ������ ����� FormFields � ������� MS Word. ����� ���� NULL.
 ���� QueryFormFields == NULL, �� ����������� ������ �������.
 ���� � ���������� wordExportParams �� ������ �����, �� �� QueryFormFields ������������ ������ ������� ������.
 �����:
   1. ������� ����� �������� ��������� ������� � ������������ DataSet.
   2. ������� ����� �������� �������� Filter � ������������ DataSet.
*/
void __fastcall TDocumentWriter::ExportToWordTemplate(TWordExportParams* wordExportParams)
// TDataSet *QueryMerge, TDataSet *QueryFormFields)
{
    //CoInitialize(NULL);
    _result.clear();


    //String TemplateFullName = AppPath + param_word.template_name; // ���������� ���� � �����-�������
    //String SavePath = ExtractFilePath(wordExportParams->resultFileDirectory);         // ���� ��� ���������� �����������
    //String ResultFileNamePrefix = ExtractFileName(DstFileName);     // ������� ����� �����-����������

    //std::vector<String> formFields;    // ������ � ������� ������ - �����������

    /*if (QueryMerge->RecordCount == 0)
    {
        return;
    }   */


    MSWordWorks msword;// = new MSWordWorks();
    Variant wordDocument;   // ������

    try
    {
        msword.OpenWord();

        #ifdef _DEBUG
        msword.SetVisible(true);
        msword.SetDisplayAlerts(true);
        #endif
        #ifndef _DEBUG
        msword.OptimizePerformance(true);
        #endif



        wordDocument =  msword.OpenDocument(wordExportParams->templateFilename, false);
    }
    catch (Exception &e)
    {
        /*
        switch (e)
        {
        case 1:
        }
        String msg = "��������� ������� ��������� ���������� Microsoft Word."
            "\n����������, ���������� � ���������� ��������������.\n" + e.Message;

        msword.CloseApplication();
        VarClear(Document);

        String msg = "��������� ������� ������ " + wordExportParams->templateFilename +
            "\n����������, ���������� � ���������� ��������������.\n" + e.Message;
        throw Exception(msg);*/
        return;

    }


    // ����� �� ��������� QueryFormFields->RecordCount ?
    // ����������������� 2017-02-15
    //bool bFilterExist = wordExportParams->filter_main_field != "" && wordExportParams->filter_sec_field != "";    // ���� � ���������� ����� ������, �� �������, ��� ���������� ������


    /* � ������ ������� ������ ��������� ���� - ����������� */
    for (std::vector<TWordSingleDataSet>::iterator ds = wordExportParams->singleImageDs.begin(); ds != wordExportParams->singleImageDs.end(); ds++ )
    {
        msword.ReplaceImageVariables(wordDocument, (*ds).dataSet, (*ds).fieldNamePrefix);
    }

    /* ������ ��������� ���� - �����*/
    for (std::vector<TWordSingleDataSet>::iterator ds = wordExportParams->singleTextDs.begin(); ds != wordExportParams->singleTextDs.end(); ds++ )
    {
        msword.ReplaceVariablesAll(wordDocument, (*ds).dataSet, DFT_DOCVARIABLE, (*ds).fieldNamePrefix); 
    }

    /* ������ ��������� ���� - FORMTEXT */
    for (std::vector<TWordSingleDataSet>::iterator ds = wordExportParams->formtextDs.begin(); ds != wordExportParams->formtextDs.end(); ds++ )
    {
        msword.ReplaceFormFields(wordDocument, (*ds).dataSet/*, (*ds).fieldNamePrefix*/);
    }


    /* ����� ��������� �������, ���� ��� ������� ���� */
    for (std::vector<TWordTableDataSet>::iterator ds = wordExportParams->tableDs.begin(); ds != wordExportParams->tableDs.end(); ds++ )
    {
        Variant table = msword.GetTableByIndex(wordDocument, (*ds).tableIndex);
        msword.writeDataSetToTable(table, (*ds).dataSet, (*ds).fieldNamePrefix);
    }

    // ��������� �������� � ������������
    //UpdateFieldsHeadersAndFooters
    msword.UpdateAllFieldsFast(wordDocument);  // �� ����� �� ��������� 2017-08-24

    /* � ������� ������ ������� */
    for (std::vector<TWordMergeDataSet>::iterator ds = wordExportParams->mergeDs.begin(); ds != wordExportParams->mergeDs.end(); ds++ )
    {
        std::vector<String> vResults;
        vResults = msword.ExportToWordFields(*ds, wordDocument, wordExportParams->resultFilename, wordExportParams->pagePerDocument);
        _result.appendResultFiles(vResults);
    }



    // ���� �� ���� �������, �� ��������� ������� �������� (����� ��� ����� � ����������� ����������� � ��������� �������)
    if (wordExportParams->mergeDs.size() == 0)
    {
        //2017-08-23
        //msword.UpdateFields(wordDocument.OlePropertyGet("Sections").OlePropertyGet("Fields"),64);
        //msword.UpdateFields(wordDocument.OlePropertyGet("Fields"),64);
        //msword.UpdateAllFields(wordDocument);

        // ����������� ��� ���� � ��������
        //msword.UnlinkFields(wordDocument.OlePropertyGet("Fields"));
        msword.UnlinkAllFields(wordDocument);

        // ���������
        msword.SaveAsDocument(wordDocument, wordExportParams->resultFilename + ".doc");
    }

    if (!VarIsEmpty(wordDocument))      // ���� ������ ������
    {
        msword.CloseDocument(wordDocument);
        VarClear(wordDocument);
    }

    msword.CloseApplication();




    /*if ()
    {
        // ���� ����� ���� ������, �� ������ ������ �������
        // ������� ��������� Word � ��������
        if (QueryMerge->RecordCount > 0)
        {

            std::vector<AnsiString> vNew;
            vNew = msword.ExportToWordFields(QueryMerge, Document, wordExportParams->resultFilename, wordExportParams->pagePerDocument);
            result.appendResultFiles(vNew);
        }
    }
    else
    {
        // ���� ������ ��� �������, ��:
        // 1. ���� ����� ������ � ����� ������ ������ ��������� �������
        // 2. ����������� �������� � FormFields-���� � �������
        // 3. ������ �������
        //int n_doc = 0;  // ���������� ����� ��������� ������� (������������ � ����� ������ �����������)
        //int nPadLength = IntToStr(QueryFormFields->RecordCount).Length();
 * /
        String oldFilter = QueryMerge->Filter;

        while ( !QueryFormFields->Eof )
        {

            if ( VarIsEmpty(Document) )           // ���� ������ �� ������, ��������� ��� (��������� �� ������ ���� �����)
            {
                Document = msword.OpenDocument(wordExportParams->templateFilename, false);
            }

            if ( bFilterExist )  // ���� �� ��������������� ������� ������ 1 ������, �� ��������� ������
            {
                try
                {
                    String sFilter = wordExportParams->filter_main_field + "='" + QueryFormFields->FieldByName(wordExportParams->filter_sec_field)->AsString + "'";
                    if (oldFilter != "")
                    {
                        sFilter = " AND " + sFilter;
                    }

                    //QueryMerge->Filtered = false;
                    QueryMerge->Filter = oldFilter + sFilter;
                    QueryMerge->Filtered = true;
                }
                catch ( Exception &e )
                {
                    QueryMerge->Filtered = false;
                    //String msg = "��������� ������������ ���������� ������� � ���������� �������� ��� ���������� � ���������� ��������������.\n" + e.Message;
                    //throw Exception(msg);
                    break;
                }

            }

            if (QueryMerge->RecordCount != 0)         // ���� ��� �������, �� ��������� ��� �����
            {

                //������ ����� FormFields
                msword.ReplaceFormFields(Document, QueryFormFields);

                // �������
                std::vector<String> vNew;   // ���������� ��� ���������� ����������� �������

                try
                {
                    vNew = msword.ExportToWordFields(QueryMerge, Document, wordExportParams->resultFilename, wordExportParams->pagePerDocument);
                }
                catch (Exception &e)
                {
                    / *_threadStatus = WM_THREAD_ERROR_IN_PROCESS;
                    _threadMessage = "� �������� ������� ��������� � ���������� ������ ��������� ������."
                        "\n���������� � ���������� ��������������."
                        "\n" + e.Message; * /
                    break;
                }

                result.appendResultFiles(vNew);
                vNew.clear();
            }



            #ifndef _DEBUG
            msword.CloseDocument(Document);
            VarClear(Document);
            #endif

            if ( bFilterExist )
            {
                QueryFormFields->Next();
            }
            else
            {
                // ���� ������ �� ����������, ����� ������� �� �����
                break;
            }
        }
    }

    if (!VarIsEmpty(Document))      // ���� ������ ������
    {
        #ifndef DEBUG
        msword.CloseDocument(Document);
        #endif
    }             */
}
























/*
 ���������� ������� MS Word
 QueryMerge - �������� ������, ������������ � �������� ��������� ������ ��� �������
 QueryFormFields - ��������������� ������, ������������ � �������� ��������� ������
 ��� ������ ����� FormFields � ������� MS Word. ����� ���� NULL.
 ���� QueryFormFields == NULL, �� ����������� ������ �������.
 ���� � ���������� wordExportParams �� ������ �����, �� �� QueryFormFields ������������ ������ ������� ������.
 �����:
   1. ������� ����� �������� ��������� ������� � ������������ DataSet.
   2. ������� ����� �������� �������� Filter � ������������ DataSet.
*/
void __fastcall TDocumentWriter::ExportToWordTemplate_old(const TWordExportParams* wordExportParams, TDataSet *QueryMerge, TDataSet *QueryFormFields)
{
    CoInitialize(NULL);
    _result.clear();


    //String TemplateFullName = AppPath + param_word.template_name; // ���������� ���� � �����-�������
    //String SavePath = ExtractFilePath(wordExportParams->resultFileDirectory);         // ���� ��� ���������� �����������
    //String ResultFileNamePrefix = ExtractFileName(DstFileName);     // ������� ����� �����-����������

    //std::vector<String> formFields;    // ������ � ������� ������ - �����������

    if (QueryMerge->RecordCount == 0)
    {
        return;
    }


    MSWordWorks msword;// = new MSWordWorks();
    Variant Document;   // ������

    try
    {
        msword.OpenWord();

        #ifdef _DEBUG
        msword.SetVisible(true);
        msword.SetDisplayAlerts(true);
        #endif

        Document =  msword.OpenDocument(wordExportParams->templateFilename, false);
    }
    catch (Exception &e)
    {
        /*
        switch (e)
        {
        case 1:
        }
        String msg = "��������� ������� ��������� ���������� Microsoft Word."
            "\n����������, ���������� � ���������� ��������������.\n" + e.Message;

        msword.CloseApplication();
        VarClear(Document);

        String msg = "��������� ������� ������ " + wordExportParams->templateFilename +
            "\n����������, ���������� � ���������� ��������������.\n" + e.Message;
        throw Exception(msg);*/
        return;

    }



    // ����� �� ��������� QueryFormFields->RecordCount ?
    bool bFilterExist = wordExportParams->filter_main_field != "" && wordExportParams->filter_sec_field != "";    // ���� � ���������� ����� ������, �� �������, ��� ���������� ������


    if (QueryFormFields == NULL)
    {
        // ���� ����� ���� ������, �� ������ ������ �������
        // ������� ��������� Word � ��������
        if (QueryMerge->RecordCount > 0)
        {

            std::vector<AnsiString> vNew;
            vNew = msword.ExportToWordFields(QueryMerge, Document, wordExportParams->resultFilename, wordExportParams->pagePerDocument);
            _result.appendResultFiles(vNew);
        }
    }
    else
    {
        // ���� ������ ��� �������, ��:
        // 1. ���� ����� ������ � ����� ������ ������ ��������� �������
        // 2. ����������� �������� � FormFields-���� � �������
        // 3. ������ �������
        //int n_doc = 0;  // ���������� ����� ��������� ������� (������������ � ����� ������ �����������)
        //int nPadLength = IntToStr(QueryFormFields->RecordCount).Length();

        String oldFilter = QueryMerge->Filter;

        while ( !QueryFormFields->Eof )
        {

            if ( VarIsEmpty(Document) )           // ���� ������ �� ������, ��������� ��� (��������� �� ������ ���� �����)
            {
                Document = msword.OpenDocument(wordExportParams->templateFilename, false);
            }

            if ( bFilterExist )  // ���� �� ��������������� ������� ������ 1 ������, �� ��������� ������
            {
                try
                {
                    String sFilter = wordExportParams->filter_main_field + "='" + QueryFormFields->FieldByName(wordExportParams->filter_sec_field)->AsString + "'";
                    if (oldFilter != "")
                    {
                        sFilter = " AND " + sFilter;
                    }

                    //QueryMerge->Filtered = false;
                    QueryMerge->Filter = oldFilter + sFilter;
                    QueryMerge->Filtered = true;
                }
                catch ( Exception &e )
                {
                    QueryMerge->Filtered = false;
                    //String msg = "��������� ������������ ���������� ������� � ���������� �������� ��� ���������� � ���������� ��������������.\n" + e.Message;
                    //throw Exception(msg);
                    break;
                }

            }

            if (QueryMerge->RecordCount != 0)         // ���� ��� �������, �� ��������� ��� �����
            {

                //������ ����� FormFields
                msword.ReplaceFormFields(Document, QueryFormFields);

                // �������
                std::vector<String> vNew;   // ���������� ��� ���������� ����������� �������

                try
                {
                    vNew = msword.ExportToWordFields(QueryMerge, Document, wordExportParams->resultFilename, wordExportParams->pagePerDocument);
                }
                catch (Exception &e)
                {
                    /*_threadStatus = WM_THREAD_ERROR_IN_PROCESS;
                    _threadMessage = "� �������� ������� ��������� � ���������� ������ ��������� ������."
                        "\n���������� � ���������� ��������������."
                        "\n" + e.Message; */
                    break;
                }

                _result.appendResultFiles(vNew);
                vNew.clear();
            }



            #ifndef _DEBUG
            msword.CloseDocument(Document);
            VarClear(Document);
            #endif

            if ( bFilterExist )
            {
                QueryFormFields->Next();
            }
            else
            {
                // ���� ������ �� ����������, ����� ������� �� �����
                break;
            }
        }
    }

    if (!VarIsEmpty(Document))      // ���� ������ ������
    {
        #ifndef _DEBUG
        msword.CloseDocument(Document);
        VarClear(Document);
        #endif
    }
    msword.CloseApplication();

    CoUninitialize();
}


/*
   ��������������� ������� ��� ��������������� �������� �������
   �� ����������
*/
Variant __fastcall TDocumentWriter::CreateExcelTemplate(TExcelExportParams* excelExportParams)
{


    //CoInitialize(NULL);
    //String TemplateFullName = excelExportParams->templateFilename; // ���������� ���� � �����-�������

    // ��������� ������ MS Excel
    MSExcelWorks msexcel;
    Variant Workbook;
    Variant Worksheet1;     // �������� ����
    Variant Worksheet2;     // ���� ��� ������ �������

    try
    {
        msexcel.OpenApplication();
        Workbook = msexcel.OpenDocument();
        Worksheet1 = msexcel.GetSheet(Workbook, 1);

        #ifdef _DEBUG
        msexcel.SetVisible(true);
        #endif

        if (msexcel.GetSheetsCount(Workbook) > 1)       // ���� � ����� ������ ������ �����, �� �������� ������ ����
        {
            Worksheet2 = msexcel.GetSheet(Workbook, 2);
        }
        else
        {
            Worksheet2 = msexcel.AddSheet(Workbook);    // ����� ��������� ����� ����
        }
    }
    catch (Exception &e)
    {
        try
        {
            msexcel.CloseApplication();
        }
        catch (...)
        {
        }
        String msg = "������ ��� �������� �������.\n���������� � ���������� ��������������.";
        throw Exception(msg);
    }

    msexcel.SetActiveWorksheet(Worksheet1); // ������������� �������� ���� 1�

    Variant range;

    /*Variant range = msexcel.GetRange(Worksheet1, 1,1,1,1);
    msexcel.AddName(Worksheet1,"doc_title", range);

    range = msexcel.GetRange(Worksheet1, 2,1,1,1);
    msexcel.AddName(Worksheet1,"info", range);

    range = msexcel.GetRange(Worksheet1, 3,1,1,1);
    msexcel.AddName(Worksheet1,"head_table", range); */


    TDataSet* tableDs = excelExportParams->tableDs[0].dataSet;


	int RecCount = tableDs->RecordCount;

    // ���������� ���������� �����
    int FieldCount = tableDs->FieldCount;

    Variant data_head;

    int ExcelFieldCount = excelExportParams->Fields.size();

    std::vector<TCellFormat> cf_tablebody;
    std::vector<TCellFormat> cf_tablehead;

    if (ExcelFieldCount < FieldCount)   // ���������� ���� ���� ���� � ExcelFields  (���� ��������� �������� �� ������)
    {
        data_head = vartools::CreateVariantArray(1, FieldCount);         // ����� �������

        excelExportParams->Fields.clear();

        // ��������� ����� �������
        for (int j = 1; j <= FieldCount; j++ )  		// ���������� ��� ����
        {
            TField* field = tableDs->Fields->FieldByNumber(j);
            // ������ ������ �������� � ������� Excel
            String sCellFormat;

            //data_head.PutElement(field->DisplayName.c_str(), 1, j);
            data_head.PutElement(Variant(field->DisplayName), 1, j);  // 2017-09-14
            //data_head.PutElement(Variant(field->Index), 1, j);


            switch (field->DataType) {  // ����� ������������ � ��������� (�������� ������� � ��.)
            case ftString:
                sCellFormat = "@";
                break;
            case ftTime:
                sCellFormat = "��:��:��";
                break;
            case ftDate:
                sCellFormat = "��.��.����";
                break;
            case ftDateTime:
                sCellFormat = "��.��.����";
                break;
            case ftCurrency: case ftFloat:
                sCellFormat = "0.00";
                break;
            case ftSmallint: case ftInteger: case ftLargeint:
                sCellFormat = "0";
                break;
            default:
                sCellFormat = "@";
            }

            TExcelField ef;

            ef.name = field->DisplayName;
            ef.format = sCellFormat;
            excelExportParams->Fields.push_back(ef);
        }
        ExcelFieldCount = excelExportParams->Fields.size();    // ��������� ���-�� ����� ���������� ��������
    }


    try      // ����������� ������ �����, ������������ ����� �������, ����������� ���� ������
    {
        data_head = vartools::CreateVariantArray(1, ExcelFieldCount);            // ����� �������

        TCellFormat cf_tablehead_tmp;   // ��������� ������ ��� �������������� ������ ����� �������
        TCellFormat cf_tablebody_tmp;   // ��������� ������ ��� �������������� ������ ���� �������


        for (unsigned int j = 0; j < ExcelFieldCount; j++)
        {
            data_head.PutElement(excelExportParams->Fields[j].name.c_str(), 1, j+1);

            cf_tablehead_tmp.BorderStyle = TCellFormat::xlContinuous;
            cf_tablehead_tmp.FontStyle = cf_tablehead_tmp.FontStyle << TCellFormat::fsBold;
            cf_tablehead_tmp.Width = excelExportParams->Fields[j].width;
            cf_tablehead.push_back(cf_tablehead_tmp);
            cf_tablehead_tmp.bWrapText = excelExportParams->Fields[j].bwraptext_head;


            //cf_tablebody_tmp.DataFormat = excelExportParams->Fields[j].format;
            cf_tablebody_tmp.DataFormat = excelExportParams->Fields[j].format;
            cf_tablebody_tmp.bWrapText = excelExportParams->Fields[j].bwraptext_body;
            cf_tablebody_tmp.BorderStyle = TCellFormat::xlContinuous;


            cf_tablebody.push_back(cf_tablebody_tmp);

            //String t1 = excelExportParams->Fields[j].format;
            //String t2 = excelExportParams->Fields[j].name;
        }
    }
    catch (Exception &e)
    {
        VarClear(data_head);
        //VarClear(data_body);

        throw Exception(e.Message);
    }


    // ������ - ��������� ���������
    Variant range_doc_title = msexcel.GetRange(Worksheet1, 1, 1, 1, 1);
    msexcel.AddName(Workbook, "report_title", range_doc_title);
    TCellFormat cf_doc_title;   // ��������� ������ ��� �������������� ������ ���� �������
    cf_doc_title.FontStyle = cf_doc_title.FontStyle << TCellFormat::fsBold;
    cf_doc_title.bSetFontColor = true;
    cf_doc_title.FontColor = clBlack;
    msexcel.SetRangeFormat(range_doc_title, cf_doc_title);


    // ������ - ����� �������� ������
    Variant range_cre_dttm = msexcel.GetRange(Worksheet1, 2, 1, 1, 1);
    msexcel.AddName(Workbook, "report_cre_dttm", range_cre_dttm);
    TCellFormat cf_cre_dttm;   // ��������� ������ ��� �������������� ������ ���� �������
    cf_cre_dttm.bSetFontColor = true;
    cf_cre_dttm.FontColor = clRed;
    //cf_cre_dttm.BorderStyle = TCellFormat::xlContinuous;
    //cf_cre_dttm.FontStyle = cf_cre_dttm.FontStyle << TCellFormat::fsBold;
    //cf_cre_dttm.Width = excelExportParams->Fields[j].width;
    //cf_cre_dttm.push_back(cf_tablehead_tmp);
    //cf_cre_dttm.bWrapText = excelExportParams->Fields[j].bwraptext_head;
    msexcel.SetRangeFormat(range_cre_dttm, cf_cre_dttm);

     // ������ ����������
    //excelExportParams->

    int currentRowNumber = 3;
    if ( !VarIsEmpty(excelExportParams->findTableVtArray("report_parameters")) )  // ���� ���� ������ ��� ������� report_parameters
    {
        Variant range_parameters = msexcel.GetRange(Worksheet1, 3, 1, 1, FieldCount);
        msexcel.AddName(Workbook, "report_parameters", range_parameters);
        TCellFormat cf_parameters;   // ��������� ������ ��� �������������� ������ ���� �������
        cf_parameters.bSetFontColor = true;
        cf_parameters.FontColor = clBlue;
        //cf_parameters.FontStyle = cf_parameters.FontStyle << TCellFormat::fsBold;
        msexcel.SetRangeFormat(range_parameters, cf_parameters);
        currentRowNumber++;
    }


    // ����� �������
    Variant range_tablehead = msexcel.GetRange(Worksheet1, currentRowNumber++, 1, 1, FieldCount);
    msexcel.AddName(Workbook, "table_head", range_tablehead);
    msexcel.WriteTableToRange(range_tablehead, data_head, 1, 1, false);
    //Variant range_tablehead = msexcel.WriteTable(Worksheet1, data_head, 3 + visible_param_count, 1);
     msexcel.SetRangeColumnsFormat(range_tablehead, cf_tablehead);

    Variant range_tablebody = msexcel.GetRange(Worksheet1, currentRowNumber, 1, 1, FieldCount);

    msexcel.AddName(Workbook,"table_body", range_tablebody);
    //msexcel.SetRangeColumnsFormat(range_tablebody, df_body);
    msexcel.SetRangeColumnsFormat(range_tablebody, cf_tablebody);


    // ����� ����� ��� ������ �������
    for (int i = 1; i <= FieldCount; i++)
    {
        TField* field = tableDs->Fields->FieldByNumber(i);
        Variant cell = msexcel.GetRangeFromRange(range_tablebody, 1, i, 1, 1);
        try
        {
            msexcel.AddName(Workbook, Variant("table_column_" + field->DisplayName), cell);  // 2017-09-11
        }
        catch (Exception &e)
        {
            //throw Exception("test " + field->DisplayName + " " + IntToStr(field->Index) + " " + field->DisplayLabel + " " + field->FullName);
            VarClear(data_head);
            msexcel.CloseWorkbook(Workbook);
            msexcel.CloseApplication();
            throw Exception("������ ��� �������� �������. ������������ ������������ ��� ��� ������ ������� \"" + field->DisplayName + "\"");
        }
    }


    // ������� ��� ������ �������
    Variant range_query_text = msexcel.GetRange(Worksheet2, 1, 1, 1, 1);
    msexcel.AddName(Workbook, "report_query_text", range_query_text);

    return Workbook;
}

//---------------------------------------------------------------------------
// ������������ ������ MS EXCEL
void __fastcall TDocumentWriter::ExportToExcel(TExcelExportParams* excelExportParams)
{
   _result.clear();

    Variant Workbook;

    MSExcelWorks msexcel;

    //if (excelExportParams->templateFilename == "")
    //{
        // ������� ������
        try
        {
            Workbook = CreateExcelTemplate(excelExportParams);
        }
        catch(Exception &e)
        {
            throw Exception(e.Message);  // 2017-09-14 ��� ����� ���� �� ����� ���������� ��� ������� �� ����������
        }
        msexcel.AssignDocument(Workbook);   // ������������ � �������
    /*}
    else
    {
        // ��� ��������� ������ �� �����
        Workbook = msexcel.OpenDocument(excelExportParams->templateFilename);

    }  */

    Variant Worksheet1 = msexcel.GetSheet(Workbook, 1);

    Variant range_doc_title = msexcel.GetRangeByName(Worksheet1, "report_title");
    msexcel.WriteToRange(range_doc_title, excelExportParams->title_label);

    Variant range_cre_dttm = msexcel.GetRangeByName(Worksheet1, "report_cre_dttm");
    msexcel.WriteToRange(range_cre_dttm, "�� ��������� ��: " + TDateTime::CurrentDateTime());

    // ������������� ��������� �������� (��������� ��������-������)
    excelExportParams->templateDocument = Workbook;

    // ������� ������
    ExportToExcelTemplate(excelExportParams);


    // ���������� �������������� ��������� ������
    Variant range_table_body = msexcel.GetRangeByName(Worksheet1, "table_body");
    Variant range_table_head = msexcel.GetRangeByName(Worksheet1, "table_head");
    Variant rangeRowsCount = msexcel.GetRangeRowsCount(range_table_body);
    Variant rangeColumnsCount = msexcel.GetRangeColumnsCount(range_table_body);
    Variant range_table_all = msexcel.GetRangeFromRange(range_table_head, 1, 1, rangeRowsCount + 1, rangeColumnsCount);
    msexcel.AddName(Worksheet1, "table_all", range_table_all);


    Variant range_tablebody = msexcel.GetRangeByName(Worksheet1, "table_body");
    msexcel.SetAutoFilter(range_table_all);   // �������� ����������

    // ����������� ������ �����
    // msexcel.SetColumnsAutofit(range_table_all);  // ������ ����� �� �����������
    for (int i=1; i<= rangeColumnsCount; i++)
    {
        int width = excelExportParams->Fields[i-1].width;
        msexcel.SetColumnWidth(range_table_all, i, width);
    }


    // ������������ ������
    /*VarClear(data_head);
    data_head = NULL;

    VarClear(data_body);
    data_body = NULL;*/

    /*if(false)
    {
        throw Exception("test");
    }  */


}



//---------------------------------------------------------------------------
// ���������� Excel ����� � �������������� ������� xlt
void __fastcall TDocumentWriter::ExportToExcelTemplate(TExcelExportParams* excelExportParams)
{
    String TemplateFullName = excelExportParams->templateFilename; // ���������� ���� � �����-�������

    // ��������� ������ MS Excel
    MSExcelWorks msexcel;

    if (VarIsClear(excelExportParams->templateDocument))
    {
        try
        {
            msexcel.OpenApplication();
            excelExportParams->templateDocument = msexcel.OpenDocument(TemplateFullName);
        }
        catch (Exception &e)
        {
            try
            {
                msexcel.CloseApplication();
            }
            catch (...)
            {
            }
            String msg = "������ ��� �������� �����-������� " + TemplateFullName + ".\n���������� � ���������� ��������������.";
            throw Exception(msg);
        }
    }
    else
    {
        msexcel.AssignDocument(excelExportParams->templateDocument);
    }

    Variant Workbook = excelExportParams->templateDocument;
    Variant Worksheet = msexcel.GetSheet(Workbook, 1);


    /* ������� Variant array */
    for (std::vector<TExcelVtArray>::iterator vtArrayIt = excelExportParams->tableVtArray.begin(); vtArrayIt != excelExportParams->tableVtArray.end(); vtArrayIt++ )
    {
        //if ( VarIsEmpty((*vtArrayIt).vtArray) ) //2017-09-11
        //{
        //    continue;
        //}
        Variant table =  msexcel.GetRangeByNameGlobal(Workbook, (*vtArrayIt).tableName);
        Variant new_range = msexcel.WriteTableToRange(table, (*vtArrayIt).vtArray, 1, 1, true);
        msexcel.ChangeNamedRange(table, new_range);
    }

    /* ������� ��������� ��������� ���� */
    for (std::vector<TExcelSingleDataSet>::iterator ds = excelExportParams->singleDs.begin(); ds != excelExportParams->singleDs.end(); ds++ )
    {
        msexcel.writeDataSetToSingleRange(Workbook, (*ds).dataSet, (*ds).fieldNamePrefix);
    }

    /* ����� ��������� �������, ���� ��� ������� ���� */
    for (std::vector<TExcelTableDataSet>::iterator ds = excelExportParams->tableDs.begin(); ds != excelExportParams->tableDs.end(); ds++ )
    {
        Variant table =  msexcel.GetRangeByName(Worksheet, (*ds).tableName);
        Variant new_range = msexcel.writeDataSetToTableRange(table, (*ds).dataSet, (*ds).fieldNamePrefix);

        msexcel.ChangeNamedRange(table, new_range);
    }

    if (excelExportParams->resultFilename != "")
    {
        msexcel.SaveDocument(Workbook, excelExportParams->resultFilename + ".xlsx");

        _result.addResultFile(excelExportParams->resultFilename + ".xlsx"); // 2017-09-08 ���������

        VarClear(Worksheet);
        msexcel.CloseApplication();
    }
    else
    {
        msexcel.SetVisible(true);
    }
}

/* ���������� DBF-�����
    2017-09-08 ���������
*/
void __fastcall TDocumentWriter::ExportToDbf(TDbaseExportParams* dbaseExportParams)
{
    _result.clear();

    TDbfUtil dbfUtil;   // �������� ������ ��� ������ � Dbf ������

    // ���� ���� allowunassigned = false
    if (dbaseExportParams->Fields.size() > dbaseExportParams->srcDataSet->FieldCount && dbaseExportParams->fDisableUnassignedFields)
    {
        String msg = "���������� ��������� ����� ��������� ���������� ����� � ��������� ������."
            "\n����������, ���������� � ���������� ��������������.";
        throw Exception(msg);
    }

    if (dbaseExportParams->Fields.size() == 0)
    {
        String msg = "�� ����� ������ ����� � ���������� ��������."
            "\n����������, ���������� � ���������� ��������������.";
        throw Exception(msg);
    }


    // ������� dbf-���� ����������
    TDbf* pTable = dbfUtil.CreateDbf(dbaseExportParams->resultFilename, &dbaseExportParams->Fields);

    // ��������� ��������� ���� �� ������
    pTable->Exclusive = true;
    try
    {
        pTable->Open();
    }
    catch (Exception &e)
    {
        delete pTable;
        throw Exception("�� ������� ������� ���� �� ������.\n" + e.Message);
    }


    // ���������� ������ �� �������� � ����
    try
    {
        dbfUtil.WriteToDbf(dbaseExportParams->srcDataSet, pTable);
        _result.addResultFile(dbaseExportParams->resultFilename);   // 2017-09-08 ���������
        pTable->Close();
    }
    catch(Exception &e)
    {
        delete pTable;
        throw Exception("�� ������� ��������� ����.\n" + e.Message );
    }

    delete pTable;
}














    //CoUninitialize();


    /*// ������� ������ ������ �����
    try
    {
        if (QueryFields != NULL)
        {
            msexcel.ExportToExcelFields(QueryFields, Worksheet);
        }
    }
    catch (Exception &e)
    {
        msexcel.CloseApplication();
        CoUninitialize();
        //String msg = e.Message;
        throw Exception(e);
    }

    // ����� ��������� ��������� �����
    try
    {
        if (QueryTable != NULL && excelExportParams->table_range_name != "") // ������ ���� ������ ��� ��������� �������� �����
        {
            msexcel.ExportToExcelTable(QueryTable, Worksheet, excelExportParams->table_range_name, excelExportParams->fUnbounded);
        }
    }
    catch (Exception &e)
    {
        try {
            msexcel.CloseApplication();
        }
        catch (...)
        {
        }
        CoUninitialize();
        //_threadMessage = e.Message;
        throw Exception(e);
    }

    if (excelExportParams->resultFilename == "")         // ������ ��������� ��������, ���� ��� �����-���������� �� ������
    {
        msexcel.SetVisible(Workbook);
    }
    else
    {                        // ����� ��������� � ����
        try
        {
            msexcel.SaveDocument(Workbook, excelExportParams->resultFilename);
            msexcel.CloseApplication();
            _result.addResultFile(excelExportParams->resultFilename);
        }
        catch (Exception &e)
        {
            try
            {
                msexcel.CloseApplication();
            }
            catch (...)
            {
            }
            CoUninitialize();
            String msg = "������ ��� ���������� ���������� � ���� " + excelExportParams->resultFilename + ".\n" + e.Message;
            throw Exception(msg);
        }
    }

    // � ���������� ������� ���������� �������� � MS Word
    // ����������� ���� ������ QueryFields � QueryTable
    */




/*
//---------------------------------------------------------------------------
// ���������� DBF-�����
// ���������� ��� ������� � ������������� ���������� TDbf
void __fastcall TDocumentWriter::ExportToDBF(TOraQuery *OraQuery)
{
    //TStringList* ListFields;
    //int n = this->param_dbase.Fields.size();
    //if (n > 0)    // ��������� ������ ����� ��� �������� � DBF ("���;���;�����;����� ������� �����")
    //{
    //    ListFields = new TStringList();
    //    for (int i = 0; i < n; i++)
    //    {
    //        ListFields->Add(param_dbase.Fields[i].name + ";" + param_dbase.Fields[i].type + ";"+ param_dbase.Fields[i].length + ";" + param_dbase.Fields[i].decimals);
    //    }
    //}
    //else
   // {
    //    _threadMessage = "�� ����� ������ ����� � ���������� ��������."
    //        "\n����������, ���������� � ���������� ��������������.";
    //    throw Exception(_threadMessage);
    //}

    // ��� ������� ������, � ����� � ���, ��� ��������� ���� ����� ���������� �������
    if (param_dbase.Fields.size() > OraQuery->FieldCount && !param_dbase.fAllowUnassignedFields)
    {
        _threadMessage = "���������� ��������� ����� ��������� ���������� ����� � ��������� ������."
            "\n����������, ���������� � ���������� ��������������.";
        throw Exception(_threadMessage);
    }

    if (param_dbase.Fields.size() == 0) {
        _threadMessage = "�� ����� ������ ����� � ���������� ��������."
            "\n����������, ���������� � ���������� ��������������.";
        throw Exception(_threadMessage);
    }

    // ������� dbf-���� ����������
    TDbf* pTable = new TDbf(NULL);

    //pTableDst->TableLevel = 7; // required for AutoInc field
    pTable->TableLevel = 4;
    pTable->LanguageID = DbfLangId_RUS_866;

    pTable->TableName = ExtractFileName(DstFileName);
    pTable->FilePathFull = ExtractFilePath(DstFileName);


    // ������� ����������� ����� ������� �� ����������
    TDbfFieldDefs* TempFieldDefs = new TDbfFieldDefs(NULL);

    if (TempFieldDefs == NULL) {
        _threadMessage = "Can't create storage.";
        throw Exception(_threadMessage);
    }

    for(std::vector<DBASEFIELD>::iterator it = param_dbase.Fields.begin(); it < param_dbase.Fields.end(); it++ )
    {
        TDbfFieldDef* TempFieldDef = TempFieldDefs->AddFieldDef();
        TempFieldDef->FieldName = it->name;
        //TempFieldDef->Required = true;
        //TempFieldDef->FieldType = Field->type;    // Use FieldType if Field->Type is TFieldType else use NativeFieldType
        TempFieldDef->NativeFieldType = it->type[1];
        TempFieldDef->Size = it->length;
        TempFieldDef->Precision = it->decimals;
    }

    if (TempFieldDefs->Count == 0)
    {
        delete pTable;
        _threadMessage = "�� ������� ��������� �������� �����.";
        throw Exception(_threadMessage);
    }

    pTable->CreateTableEx(TempFieldDefs);
    pTable->Exclusive = true;
    try
    {
        pTable->Open();
    }
    catch (Exception &e)
    {
        _threadMessage = e.Message;
    }

    // ������ ������ � �������
    try
    {
	    while ( !OraQuery->Eof )
        {
            pTable->Append();
		    for (int j = 1; j <= OraQuery->FieldCount; j++ )
            {
                pTable->Fields->FieldByNumber(j)->Value = OraQuery->Fields->FieldByNumber(j)->Value;
            }
            OraQuery->Next();  // ��������� � ��������� ������
	    }
        pTable->Post();
        pTable->Close();

        _resultFiles.push_back(DstFileName);

    }
    catch(Exception &e)
    {
        pTable->Close();

        delete TempFieldDefs;
        delete pTable;

        _threadMessage = e.Message;
        throw Exception(e);
    }

    delete TempFieldDefs;
    delete pTable;
}       */

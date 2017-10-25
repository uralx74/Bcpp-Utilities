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
 Заполнение шаблона MS Word
 QueryMerge - основной запрос, используется в качестве источника данных при слиянии
 QueryFormFields - вспомогательный запрос, используется в качестве источника данных
 при замене полей FormFields в шаблоне MS Word. Может быть NULL.
 Если QueryFormFields == NULL, то выполняется только слияние.
 Если в параметрах wordExportParams не задана связь, то из QueryFormFields используется только текущая строка.
 ВАЖНО:
   1. Функция может изменить положение курсора в передаваемых DataSet.
   2. Функция может изменить значение Filter в передаваемых DataSet.
*/
void __fastcall TDocumentWriter::ExportToWordTemplate(TWordExportParams* wordExportParams)
// TDataSet *QueryMerge, TDataSet *QueryFormFields)
{
    //CoInitialize(NULL);
    _result.clear();


    //String TemplateFullName = AppPath + param_word.template_name; // Абсолютный путь к файлу-шаблону
    //String SavePath = ExtractFilePath(wordExportParams->resultFileDirectory);         // Путь для сохранения результатов
    //String ResultFileNamePrefix = ExtractFileName(DstFileName);     // Префикс имени файла-результата

    //std::vector<String> formFields;    // Вектор с именами файлов - результатов

    /*if (QueryMerge->RecordCount == 0)
    {
        return;
    }   */


    MSWordWorks msword;// = new MSWordWorks();
    Variant wordDocument;   // Шаблон

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
        String msg = "Неудалось создать экземпляр приложения Microsoft Word."
            "\nПожалуйста, обратитесь к системному администратору.\n" + e.Message;

        msword.CloseApplication();
        VarClear(Document);

        String msg = "Неудалось открыть шаблон " + wordExportParams->templateFilename +
            "\nПожалуйста, обратитесь к системному администратору.\n" + e.Message;
        throw Exception(msg);*/
        return;

    }


    // Нужно ли учитывать QueryFormFields->RecordCount ?
    // закомментированно 2017-02-15
    //bool bFilterExist = wordExportParams->filter_main_field != "" && wordExportParams->filter_sec_field != "";    // Если в параметрах задан фильтр, то считаем, что установлен фильтр


    /* В первую очередь меняем одиночные поля - изображения */
    for (std::vector<TWordSingleDataSet>::iterator ds = wordExportParams->singleImageDs.begin(); ds != wordExportParams->singleImageDs.end(); ds++ )
    {
        msword.ReplaceImageVariables(wordDocument, (*ds).dataSet, (*ds).fieldNamePrefix);
    }

    /* Меняем одиночные поля - текст*/
    for (std::vector<TWordSingleDataSet>::iterator ds = wordExportParams->singleTextDs.begin(); ds != wordExportParams->singleTextDs.end(); ds++ )
    {
        msword.ReplaceVariablesAll(wordDocument, (*ds).dataSet, DFT_DOCVARIABLE, (*ds).fieldNamePrefix); 
    }

    /* Меняем одиночные поля - FORMTEXT */
    for (std::vector<TWordSingleDataSet>::iterator ds = wordExportParams->formtextDs.begin(); ds != wordExportParams->formtextDs.end(); ds++ )
    {
        msword.ReplaceFormFields(wordDocument, (*ds).dataSet/*, (*ds).fieldNamePrefix*/);
    }


    /* Затем заполняем таблицы, если они конечно есть */
    for (std::vector<TWordTableDataSet>::iterator ds = wordExportParams->tableDs.begin(); ds != wordExportParams->tableDs.end(); ds++ )
    {
        Variant table = msword.GetTableByIndex(wordDocument, (*ds).tableIndex);
        msword.writeDataSetToTable(table, (*ds).dataSet, (*ds).fieldNamePrefix);
    }

    // Обновляем значения в колонтитулах
    //UpdateFieldsHeadersAndFooters
    msword.UpdateAllFieldsFast(wordDocument);  // До конца не проверена 2017-08-24

    /* И наконец делаем слияние */
    for (std::vector<TWordMergeDataSet>::iterator ds = wordExportParams->mergeDs.begin(); ds != wordExportParams->mergeDs.end(); ds++ )
    {
        std::vector<String> vResults;
        vResults = msword.ExportToWordFields(*ds, wordDocument, wordExportParams->resultFilename, wordExportParams->pagePerDocument);
        _result.appendResultFiles(vResults);
    }



    // Если не было слияния, то сохраняем текущий документ (иначе при файлы с результатом сохраняются в процедуре слияния)
    if (wordExportParams->mergeDs.size() == 0)
    {
        //2017-08-23
        //msword.UpdateFields(wordDocument.OlePropertyGet("Sections").OlePropertyGet("Fields"),64);
        //msword.UpdateFields(wordDocument.OlePropertyGet("Fields"),64);
        //msword.UpdateAllFields(wordDocument);

        // Преобразуем все поля в значения
        //msword.UnlinkFields(wordDocument.OlePropertyGet("Fields"));
        msword.UnlinkAllFields(wordDocument);

        // Сохраняем
        msword.SaveAsDocument(wordDocument, wordExportParams->resultFilename + ".doc");
    }

    if (!VarIsEmpty(wordDocument))      // Если шаблон открыт
    {
        msword.CloseDocument(wordDocument);
        VarClear(wordDocument);
    }

    msword.CloseApplication();




    /*if ()
    {
        // Если задан один запрос, то делаем только слияние
        // Слияние документа Word с таблицей
        if (QueryMerge->RecordCount > 0)
        {

            std::vector<AnsiString> vNew;
            vNew = msword.ExportToWordFields(QueryMerge, Document, wordExportParams->resultFilename, wordExportParams->pagePerDocument);
            result.appendResultFiles(vNew);
        }
    }
    else
    {
        // Если задано два запроса, то:
        // 1. если задан фильтр в цикле задаем фильтр основному запросу
        // 2. подставляем значения в FormFields-поля в шаблоне
        // 3. делаем слияние
        //int n_doc = 0;  // Порядковый номер процедуры слияния (используется в имени файлов результатов)
        //int nPadLength = IntToStr(QueryFormFields->RecordCount).Length();
 * /
        String oldFilter = QueryMerge->Filter;

        while ( !QueryFormFields->Eof )
        {

            if ( VarIsEmpty(Document) )           // Если шаблон не открыт, открываем его (требуется на втором шаге цикла)
            {
                Document = msword.OpenDocument(wordExportParams->templateFilename, false);
            }

            if ( bFilterExist )  // Если во вспомогательном запросе больше 1 строки, то применяем фильтр
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
                    //String msg = "Проверьте корректность параметров фильтра в параметрах экспорта или обратитесь к системному администратору.\n" + e.Message;
                    //throw Exception(msg);
                    break;
                }

            }

            if (QueryMerge->RecordCount != 0)         // Если нет записей, то следующий шаг цикла
            {

                //Замена полей FormFields
                msword.ReplaceFormFields(Document, QueryFormFields);

                // Слияние
                std::vector<String> vNew;   // переменная для сохранения результатов слияния

                try
                {
                    vNew = msword.ExportToWordFields(QueryMerge, Document, wordExportParams->resultFilename, wordExportParams->pagePerDocument);
                }
                catch (Exception &e)
                {
                    / *_threadStatus = WM_THREAD_ERROR_IN_PROCESS;
                    _threadMessage = "В процессе слияния документа с источником данных произошла ошибка."
                        "\nОбратитесь к системному администратору."
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
                // Если фильтр не установлен, тогда выходим из цикла
                break;
            }
        }
    }

    if (!VarIsEmpty(Document))      // Если шаблон открыт
    {
        #ifndef DEBUG
        msword.CloseDocument(Document);
        #endif
    }             */
}
























/*
 Заполнение шаблона MS Word
 QueryMerge - основной запрос, используется в качестве источника данных при слиянии
 QueryFormFields - вспомогательный запрос, используется в качестве источника данных
 при замене полей FormFields в шаблоне MS Word. Может быть NULL.
 Если QueryFormFields == NULL, то выполняется только слияние.
 Если в параметрах wordExportParams не задана связь, то из QueryFormFields используется только текущая строка.
 ВАЖНО:
   1. Функция может изменить положение курсора в передаваемых DataSet.
   2. Функция может изменить значение Filter в передаваемых DataSet.
*/
void __fastcall TDocumentWriter::ExportToWordTemplate_old(const TWordExportParams* wordExportParams, TDataSet *QueryMerge, TDataSet *QueryFormFields)
{
    CoInitialize(NULL);
    _result.clear();


    //String TemplateFullName = AppPath + param_word.template_name; // Абсолютный путь к файлу-шаблону
    //String SavePath = ExtractFilePath(wordExportParams->resultFileDirectory);         // Путь для сохранения результатов
    //String ResultFileNamePrefix = ExtractFileName(DstFileName);     // Префикс имени файла-результата

    //std::vector<String> formFields;    // Вектор с именами файлов - результатов

    if (QueryMerge->RecordCount == 0)
    {
        return;
    }


    MSWordWorks msword;// = new MSWordWorks();
    Variant Document;   // Шаблон

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
        String msg = "Неудалось создать экземпляр приложения Microsoft Word."
            "\nПожалуйста, обратитесь к системному администратору.\n" + e.Message;

        msword.CloseApplication();
        VarClear(Document);

        String msg = "Неудалось открыть шаблон " + wordExportParams->templateFilename +
            "\nПожалуйста, обратитесь к системному администратору.\n" + e.Message;
        throw Exception(msg);*/
        return;

    }



    // Нужно ли учитывать QueryFormFields->RecordCount ?
    bool bFilterExist = wordExportParams->filter_main_field != "" && wordExportParams->filter_sec_field != "";    // Если в параметрах задан фильтр, то считаем, что установлен фильтр


    if (QueryFormFields == NULL)
    {
        // Если задан один запрос, то делаем только слияние
        // Слияние документа Word с таблицей
        if (QueryMerge->RecordCount > 0)
        {

            std::vector<AnsiString> vNew;
            vNew = msword.ExportToWordFields(QueryMerge, Document, wordExportParams->resultFilename, wordExportParams->pagePerDocument);
            _result.appendResultFiles(vNew);
        }
    }
    else
    {
        // Если задано два запроса, то:
        // 1. если задан фильтр в цикле задаем фильтр основному запросу
        // 2. подставляем значения в FormFields-поля в шаблоне
        // 3. делаем слияние
        //int n_doc = 0;  // Порядковый номер процедуры слияния (используется в имени файлов результатов)
        //int nPadLength = IntToStr(QueryFormFields->RecordCount).Length();

        String oldFilter = QueryMerge->Filter;

        while ( !QueryFormFields->Eof )
        {

            if ( VarIsEmpty(Document) )           // Если шаблон не открыт, открываем его (требуется на втором шаге цикла)
            {
                Document = msword.OpenDocument(wordExportParams->templateFilename, false);
            }

            if ( bFilterExist )  // Если во вспомогательном запросе больше 1 строки, то применяем фильтр
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
                    //String msg = "Проверьте корректность параметров фильтра в параметрах экспорта или обратитесь к системному администратору.\n" + e.Message;
                    //throw Exception(msg);
                    break;
                }

            }

            if (QueryMerge->RecordCount != 0)         // Если нет записей, то следующий шаг цикла
            {

                //Замена полей FormFields
                msword.ReplaceFormFields(Document, QueryFormFields);

                // Слияние
                std::vector<String> vNew;   // переменная для сохранения результатов слияния

                try
                {
                    vNew = msword.ExportToWordFields(QueryMerge, Document, wordExportParams->resultFilename, wordExportParams->pagePerDocument);
                }
                catch (Exception &e)
                {
                    /*_threadStatus = WM_THREAD_ERROR_IN_PROCESS;
                    _threadMessage = "В процессе слияния документа с источником данных произошла ошибка."
                        "\nОбратитесь к системному администратору."
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
                // Если фильтр не установлен, тогда выходим из цикла
                break;
            }
        }
    }

    if (!VarIsEmpty(Document))      // Если шаблон открыт
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
   Вспомогательная функция для автоматического создания шаблона
   из параметров
*/
Variant __fastcall TDocumentWriter::CreateExcelTemplate(TExcelExportParams* excelExportParams)
{


    //CoInitialize(NULL);
    //String TemplateFullName = excelExportParams->templateFilename; // Абсолютный путь к файлу-шаблону

    // Открываем шаблон MS Excel
    MSExcelWorks msexcel;
    Variant Workbook;
    Variant Worksheet1;     // Основной лист
    Variant Worksheet2;     // Лист для текста запроса

    try
    {
        msexcel.OpenApplication();
        Workbook = msexcel.OpenDocument();
        Worksheet1 = msexcel.GetSheet(Workbook, 1);

        #ifdef _DEBUG
        msexcel.SetVisible(true);
        #endif

        if (msexcel.GetSheetsCount(Workbook) > 1)       // Если в книге больше одного листа, то получаем второй лист
        {
            Worksheet2 = msexcel.GetSheet(Workbook, 2);
        }
        else
        {
            Worksheet2 = msexcel.AddSheet(Workbook);    // иначе добавляем новый лист
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
        String msg = "Ошибка при создании шаблона.\nОбратитесь к системному администратору.";
        throw Exception(msg);
    }

    msexcel.SetActiveWorksheet(Worksheet1); // Устанавливаем активный лист 1й

    Variant range;

    /*Variant range = msexcel.GetRange(Worksheet1, 1,1,1,1);
    msexcel.AddName(Worksheet1,"doc_title", range);

    range = msexcel.GetRange(Worksheet1, 2,1,1,1);
    msexcel.AddName(Worksheet1,"info", range);

    range = msexcel.GetRange(Worksheet1, 3,1,1,1);
    msexcel.AddName(Worksheet1,"head_table", range); */


    TDataSet* tableDs = excelExportParams->tableDs[0].dataSet;


	int RecCount = tableDs->RecordCount;

    // Определяем количество полей
    int FieldCount = tableDs->FieldCount;

    Variant data_head;

    int ExcelFieldCount = excelExportParams->Fields.size();

    std::vector<TCellFormat> cf_tablebody;
    std::vector<TCellFormat> cf_tablehead;

    if (ExcelFieldCount < FieldCount)   // Заполнение если есть поля в ExcelFields  (если параметры экспорта не заданы)
    {
        data_head = vartools::CreateVariantArray(1, FieldCount);         // Шапка таблицы

        excelExportParams->Fields.clear();

        // Формируем шапку таблицы
        for (int j = 1; j <= FieldCount; j++ )  		// Перебираем все поля
        {
            TField* field = tableDs->Fields->FieldByNumber(j);
            // Задаем формат столбцов в таблице Excel
            String sCellFormat;

            //data_head.PutElement(field->DisplayName.c_str(), 1, j);
            data_head.PutElement(Variant(field->DisplayName), 1, j);  // 2017-09-14
            //data_head.PutElement(Variant(field->Index), 1, j);


            switch (field->DataType) {  // Нужно тестирование и доработка (добавить форматы и тд.)
            case ftString:
                sCellFormat = "@";
                break;
            case ftTime:
                sCellFormat = "чч:мм:сс";
                break;
            case ftDate:
                sCellFormat = "ДД.ММ.ГГГГ";
                break;
            case ftDateTime:
                sCellFormat = "ДД.ММ.ГГГГ";
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
        ExcelFieldCount = excelExportParams->Fields.size();    // Обновляем кол-во полей параметров экспорта
    }


    try      // Определение списка полей, формирование шапки таблицы, определение типа данных
    {
        data_head = vartools::CreateVariantArray(1, ExcelFieldCount);            // Шапка таблицы

        TCellFormat cf_tablehead_tmp;   // Временный объект для форматирования ячейки шапки таблицы
        TCellFormat cf_tablebody_tmp;   // Временный объект для форматирования ячейки тела таблицы


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


    // Раздел - Заголовок документа
    Variant range_doc_title = msexcel.GetRange(Worksheet1, 1, 1, 1, 1);
    msexcel.AddName(Workbook, "report_title", range_doc_title);
    TCellFormat cf_doc_title;   // Временный объект для форматирования ячейки тела таблицы
    cf_doc_title.FontStyle = cf_doc_title.FontStyle << TCellFormat::fsBold;
    cf_doc_title.bSetFontColor = true;
    cf_doc_title.FontColor = clBlack;
    msexcel.SetRangeFormat(range_doc_title, cf_doc_title);


    // Раздел - Время создания отчета
    Variant range_cre_dttm = msexcel.GetRange(Worksheet1, 2, 1, 1, 1);
    msexcel.AddName(Workbook, "report_cre_dttm", range_cre_dttm);
    TCellFormat cf_cre_dttm;   // Временный объект для форматирования ячейки тела таблицы
    cf_cre_dttm.bSetFontColor = true;
    cf_cre_dttm.FontColor = clRed;
    //cf_cre_dttm.BorderStyle = TCellFormat::xlContinuous;
    //cf_cre_dttm.FontStyle = cf_cre_dttm.FontStyle << TCellFormat::fsBold;
    //cf_cre_dttm.Width = excelExportParams->Fields[j].width;
    //cf_cre_dttm.push_back(cf_tablehead_tmp);
    //cf_cre_dttm.bWrapText = excelExportParams->Fields[j].bwraptext_head;
    msexcel.SetRangeFormat(range_cre_dttm, cf_cre_dttm);

     // Раздел параметров
    //excelExportParams->

    int currentRowNumber = 3;
    if ( !VarIsEmpty(excelExportParams->findTableVtArray("report_parameters")) )  // Если есть данные для раздела report_parameters
    {
        Variant range_parameters = msexcel.GetRange(Worksheet1, 3, 1, 1, FieldCount);
        msexcel.AddName(Workbook, "report_parameters", range_parameters);
        TCellFormat cf_parameters;   // Временный объект для форматирования ячейки тела таблицы
        cf_parameters.bSetFontColor = true;
        cf_parameters.FontColor = clBlue;
        //cf_parameters.FontStyle = cf_parameters.FontStyle << TCellFormat::fsBold;
        msexcel.SetRangeFormat(range_parameters, cf_parameters);
        currentRowNumber++;
    }


    // Шапка таблицы
    Variant range_tablehead = msexcel.GetRange(Worksheet1, currentRowNumber++, 1, 1, FieldCount);
    msexcel.AddName(Workbook, "table_head", range_tablehead);
    msexcel.WriteTableToRange(range_tablehead, data_head, 1, 1, false);
    //Variant range_tablehead = msexcel.WriteTable(Worksheet1, data_head, 3 + visible_param_count, 1);
     msexcel.SetRangeColumnsFormat(range_tablehead, cf_tablehead);

    Variant range_tablebody = msexcel.GetRange(Worksheet1, currentRowNumber, 1, 1, FieldCount);

    msexcel.AddName(Workbook,"table_body", range_tablebody);
    //msexcel.SetRangeColumnsFormat(range_tablebody, df_body);
    msexcel.SetRangeColumnsFormat(range_tablebody, cf_tablebody);


    // Имена ячеек для вывода таблицы
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
            throw Exception("Ошибка при создании шаблона. Использовано недопустимое имя для ячейки таблицы \"" + field->DisplayName + "\"");
        }
    }


    // Таблица для текста запроса
    Variant range_query_text = msexcel.GetRange(Worksheet2, 1, 1, 1, 1);
    msexcel.AddName(Workbook, "report_query_text", range_query_text);

    return Workbook;
}

//---------------------------------------------------------------------------
// ФОРМИРОВАНИЕ ОТЧЕТА MS EXCEL
void __fastcall TDocumentWriter::ExportToExcel(TExcelExportParams* excelExportParams)
{
   _result.clear();

    Variant Workbook;

    MSExcelWorks msexcel;

    //if (excelExportParams->templateFilename == "")
    //{
        // Создаем шаблон
        try
        {
            Workbook = CreateExcelTemplate(excelExportParams);
        }
        catch(Exception &e)
        {
            throw Exception(e.Message);  // 2017-09-14 Без этого выше не ловит исключения при запуске из проводника
        }
        msexcel.AssignDocument(Workbook);   // Присоединяем к объекту
    /*}
    else
    {
        // или открываем шаблон из файла
        Workbook = msexcel.OpenDocument(excelExportParams->templateFilename);

    }  */

    Variant Worksheet1 = msexcel.GetSheet(Workbook, 1);

    Variant range_doc_title = msexcel.GetRangeByName(Worksheet1, "report_title");
    msexcel.WriteToRange(range_doc_title, excelExportParams->title_label);

    Variant range_cre_dttm = msexcel.GetRangeByName(Worksheet1, "report_cre_dttm");
    msexcel.WriteToRange(range_cre_dttm, "По состоянию на: " + TDateTime::CurrentDateTime());

    // Устанавливаем параметры экспорта (назначаем документ-шаблон)
    excelExportParams->templateDocument = Workbook;

    // Выводим данные
    ExportToExcelTemplate(excelExportParams);


    // Производим дополнительные настройки отчета
    Variant range_table_body = msexcel.GetRangeByName(Worksheet1, "table_body");
    Variant range_table_head = msexcel.GetRangeByName(Worksheet1, "table_head");
    Variant rangeRowsCount = msexcel.GetRangeRowsCount(range_table_body);
    Variant rangeColumnsCount = msexcel.GetRangeColumnsCount(range_table_body);
    Variant range_table_all = msexcel.GetRangeFromRange(range_table_head, 1, 1, rangeRowsCount + 1, rangeColumnsCount);
    msexcel.AddName(Worksheet1, "table_all", range_table_all);


    Variant range_tablebody = msexcel.GetRangeByName(Worksheet1, "table_body");
    msexcel.SetAutoFilter(range_table_all);   // Включаем автофильтр

    // Настраиваем ширину ячеек
    // msexcel.SetColumnsAutofit(range_table_all);  // Ширина ячеек по содержимому
    for (int i=1; i<= rangeColumnsCount; i++)
    {
        int width = excelExportParams->Fields[i-1].width;
        msexcel.SetColumnWidth(range_table_all, i, width);
    }


    // Освобождение памяти
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
// Заполнение Excel файла с использованием шаблона xlt
void __fastcall TDocumentWriter::ExportToExcelTemplate(TExcelExportParams* excelExportParams)
{
    String TemplateFullName = excelExportParams->templateFilename; // Абсолютный путь к файлу-шаблону

    // Открываем шаблон MS Excel
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
            String msg = "Ошибка при открытии файла-шаблона " + TemplateFullName + ".\nОбратитесь к системному администратору.";
            throw Exception(msg);
        }
    }
    else
    {
        msexcel.AssignDocument(excelExportParams->templateDocument);
    }

    Variant Workbook = excelExportParams->templateDocument;
    Variant Worksheet = msexcel.GetSheet(Workbook, 1);


    /* Выводим Variant array */
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

    /* Сначала заполняем отдельные поля */
    for (std::vector<TExcelSingleDataSet>::iterator ds = excelExportParams->singleDs.begin(); ds != excelExportParams->singleDs.end(); ds++ )
    {
        msexcel.writeDataSetToSingleRange(Workbook, (*ds).dataSet, (*ds).fieldNamePrefix);
    }

    /* Затем заполняем таблицы, если они конечно есть */
    for (std::vector<TExcelTableDataSet>::iterator ds = excelExportParams->tableDs.begin(); ds != excelExportParams->tableDs.end(); ds++ )
    {
        Variant table =  msexcel.GetRangeByName(Worksheet, (*ds).tableName);
        Variant new_range = msexcel.writeDataSetToTableRange(table, (*ds).dataSet, (*ds).fieldNamePrefix);

        msexcel.ChangeNamedRange(table, new_range);
    }

    if (excelExportParams->resultFilename != "")
    {
        msexcel.SaveDocument(Workbook, excelExportParams->resultFilename + ".xlsx");

        _result.addResultFile(excelExportParams->resultFilename + ".xlsx"); // 2017-09-08 Проверить

        VarClear(Worksheet);
        msexcel.CloseApplication();
    }
    else
    {
        msexcel.SetVisible(true);
    }
}

/* Заполнение DBF-файла
    2017-09-08 проверить
*/
void __fastcall TDocumentWriter::ExportToDbf(TDbaseExportParams* dbaseExportParams)
{
    _result.clear();

    TDbfUtil dbfUtil;   // Основной объект для работы с Dbf файлом

    // Если флаг allowunassigned = false
    if (dbaseExportParams->Fields.size() > dbaseExportParams->srcDataSet->FieldCount && dbaseExportParams->fDisableUnassignedFields)
    {
        String msg = "Количество требуемых полей превышает количество полей в источнике данных."
            "\nПожалуйста, обратитесь к системному администратору.";
        throw Exception(msg);
    }

    if (dbaseExportParams->Fields.size() == 0)
    {
        String msg = "Не задан список полей в параметрах экспорта."
            "\nПожалуйста, обратитесь к системному администратору.";
        throw Exception(msg);
    }


    // Создаем dbf-файл назначения
    TDbf* pTable = dbfUtil.CreateDbf(dbaseExportParams->resultFilename, &dbaseExportParams->Fields);

    // Открываем созданный файл на запись
    pTable->Exclusive = true;
    try
    {
        pTable->Open();
    }
    catch (Exception &e)
    {
        delete pTable;
        throw Exception("Не удалось открыть файл на запись.\n" + e.Message);
    }


    // Сохранение данных из датасета в файл
    try
    {
        dbfUtil.WriteToDbf(dbaseExportParams->srcDataSet, pTable);
        _result.addResultFile(dbaseExportParams->resultFilename);   // 2017-09-08 Проверить
        pTable->Close();
    }
    catch(Exception &e)
    {
        delete pTable;
        throw Exception("Не удалось заполнить файл.\n" + e.Message );
    }

    delete pTable;
}














    //CoUninitialize();


    /*// Сначала делаем замену полей
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

    // Затем вставляем табличную часть
    try
    {
        if (QueryTable != NULL && excelExportParams->table_range_name != "") // Должно быть задано имя диапазона таблично части
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

    if (excelExportParams->resultFilename == "")         // Просто открываем документ, если имя файла-результата не задано
    {
        msexcel.SetVisible(Workbook);
    }
    else
    {                        // иначе сохраняем в файл
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
            String msg = "Ошибка при сохранении результата в файл " + excelExportParams->resultFilename + ".\n" + e.Message;
            throw Exception(msg);
        }
    }

    // В дальнейшем сделать аналогично выгрузке в MS Word
    // обьединение двух таблиц QueryFields и QueryTable
    */




/*
//---------------------------------------------------------------------------
// Заполнение DBF-файла
// Переделать эту функцию с использование компонента TDbf
void __fastcall TDocumentWriter::ExportToDBF(TOraQuery *OraQuery)
{
    //TStringList* ListFields;
    //int n = this->param_dbase.Fields.size();
    //if (n > 0)    // Формируем список полей для экспорта в DBF ("Имя;Тип;Длина;Длина дробной части")
    //{
    //    ListFields = new TStringList();
    //    for (int i = 0; i < n; i++)
    //    {
    //        ListFields->Add(param_dbase.Fields[i].name + ";" + param_dbase.Fields[i].type + ";"+ param_dbase.Fields[i].length + ";" + param_dbase.Fields[i].decimals);
    //    }
    //}
    //else
   // {
    //    _threadMessage = "Не задан список полей в параметрах экспорта."
    //        "\nПожалуйста, обратитесь к системному администратору.";
    //    throw Exception(_threadMessage);
    //}

    // Это условие убрано, в связи с тем, что некоторые поля могут оставаться пустыми
    if (param_dbase.Fields.size() > OraQuery->FieldCount && !param_dbase.fAllowUnassignedFields)
    {
        _threadMessage = "Количество требуемых полей превышает количество полей в источнике данных."
            "\nПожалуйста, обратитесь к системному администратору.";
        throw Exception(_threadMessage);
    }

    if (param_dbase.Fields.size() == 0) {
        _threadMessage = "Не задан список полей в параметрах экспорта."
            "\nПожалуйста, обратитесь к системному администратору.";
        throw Exception(_threadMessage);
    }

    // Создаем dbf-файл назначения
    TDbf* pTable = new TDbf(NULL);

    //pTableDst->TableLevel = 7; // required for AutoInc field
    pTable->TableLevel = 4;
    pTable->LanguageID = DbfLangId_RUS_866;

    pTable->TableName = ExtractFileName(DstFileName);
    pTable->FilePathFull = ExtractFilePath(DstFileName);


    // Создаем определение полей таблицы из параметров
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
        _threadMessage = "Не удалось загрузить описание полей.";
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

    // Запись данных в таблицу
    try
    {
	    while ( !OraQuery->Eof )
        {
            pTable->Append();
		    for (int j = 1; j <= OraQuery->FieldCount; j++ )
            {
                pTable->Fields->FieldByNumber(j)->Value = OraQuery->Fields->FieldByNumber(j)->Value;
            }
            OraQuery->Next();  // Переходим к следующей записи
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

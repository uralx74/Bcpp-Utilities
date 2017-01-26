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
void __fastcall TDocumentWriter::ExportToWordTemplate(const TWordExportParams* wordExportParams, TDataSet *QueryMerge, TDataSet *QueryFormFields)
{
    CoInitialize(NULL);
    result.clear();


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

        #ifndef NDEBUG
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

                result.appendResultFiles(vNew);
                vNew.clear();
            }



            #ifndef DEBUG
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
        VarClear(Document);
        #endif
    }
    msword.CloseApplication();

    CoUninitialize();
}

/*
//---------------------------------------------------------------------------
// ФОРМИРОВАНИЕ ОТЧЕТА MS EXCEL
void __fastcall TDocumentWriter::ExportToExcel(TOraQuery *OraQuery)
{
    CoInitialize(NULL);

    bool fDone = false;

    // Определяем количество записей
    OraQuery->Last();
	int RecCount = OraQuery->RecordCount;

    // Определяем количество полей
    int FieldCount = OraQuery->FieldCount;

    Variant data_body;
    Variant data_head;
    DATAFORMAT df_body;
    df_body.reserve(FieldCount);

    int ExcelFieldCount = param_excel.Fields.size();

    try {     // Определение списка полей, формирование шапки таблицы, определение типа данных
        //data_body = CreateVariantArray(RecCount, FieldCount);  // Создаем массив для таблицы
        //data_head = CreateVariantArray(1, FieldCount);  // Шапка таблицы

        if (ExcelFieldCount >= FieldCount)   // Заполнение если есть поля в ExcelFields
        {
            data_body = CreateVariantArray(RecCount, ExcelFieldCount);     // Создаем массив для таблицы
            data_head = CreateVariantArray(1, ExcelFieldCount);            // Шапка таблицы

            for (unsigned int j = 0; j < ExcelFieldCount; j++)
            {
                data_head.PutElement(param_excel.Fields[j].name, 1, j+1);
                df_body.push_back(param_excel.Fields[j].format);
            }
        }
        else
        {
            data_body = CreateVariantArray(RecCount, FieldCount);  // Создаем массив для таблицы
            data_head = CreateVariantArray(1, FieldCount);         // Шапка таблицы

            // Формируем шапку таблицы
            for (int j = 1; j <= FieldCount; j++ )  		// Перебираем все поля
            {
                TField* field = OraQuery->Fields->FieldByNumber(j);
                // Задаем формат столбцов в таблице Excel
                AnsiString sCellFormat;

                data_head.PutElement(field->DisplayName, 1, j);
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
                df_body.push_back(sCellFormat);
	        }
        }
    }
    catch (Exception &e)
    {
        VarClear(data_head);
        VarClear(data_body);
        CoUninitialize();
        _threadMessage = e.Message;
        throw Exception(_threadMessage);
        //fDone = true;
    }

    MSExcelWorks msexcel;
    Variant Workbook;
    Variant Worksheet1;

    if (!fDone)           // Заполнение массива данных
    {
        AnsiString s = "";
        OraQuery->First();	// Переходим к первой записи (на всякий случай)
        VarArrayLock(data_body);
        int i = 1;          // Пропускаем шапку таблицы
	    while (!OraQuery->Eof)
        {
		    for (int j = 1; j <= FieldCount; j++ )
            {
        	    s = OraQuery->Fields->FieldByNumber(j)->AsString;
                data_body.PutElement(s, i, j);
            }
            OraQuery->Next();  // Переходим к следующей записи
            i++;
	    }
        VarArrayUnlock(data_body);
       try
       {
            msexcel.OpenApplication();
            Workbook = msexcel.OpenDocument();
        }
        catch (Exception &e)
        {
            VarClear(data_head);
            VarClear(data_body);
            CoUninitialize();
            _threadMessage = e.Message;
            throw Exception(_threadMessage);
        }
        Worksheet1 = msexcel.GetSheet(Workbook, 1);
    }


    if (!fDone && !VarIsEmpty(Worksheet1))  // Заполняем документ Excel
    {
    //if (!VarIsEmpty(Worksheet1)) { // Заполняем документ Excel
        TDateTime DateTime = TDateTime::CurrentDateTime();

        CELLFORMAT cf_body;
        CELLFORMAT cf_head;
        CELLFORMAT cf_title;
        CELLFORMAT cf_createtime;
        CELLFORMAT cf_sql;

        cf_body.BorderStyle = CELLFORMAT::xlContinuous;
        cf_head.BorderStyle = CELLFORMAT::xlContinuous;
        cf_head.FontStyle = cf_head.FontStyle << CELLFORMAT::fsBold;

        cf_head.bWrapText = false;

        cf_title.FontStyle = cf_title.FontStyle << CELLFORMAT::fsBold;
        cf_createtime.bSetFontColor = true;
        cf_createtime.FontColor = clRed;
        cf_sql.bWrapText = false;

        // Определяем формат данных
        //std::vector<MSExcelWorks::CELLFORMAT> formats;
        //formats = msexcel.GetDataFormat(ArrayDataBody, 1);
        //std::vector<AnsiString> DataFormat;
        //DataFormat = msexcel.GetDataFormat(ArrayDataBody, 1);
        //for (int i=0; i  < QueryParams.size(); i++) {
            //QueryParams[i].
        //}


        // Заполняем массив, со значеними параметров, заданными пользователем
        // Возможно в будущем сделать распознавание параметра с типом "separator",
        Variant data_parameters;
        int param_count = UserParams.size();  // Параметры отчета
        int visible_param_count = 0;
        for (int i=0; i <= param_count-1; i++)    // Подсчитываем кол-во отображаемых параметров
        {
            if ( UserParams[i]->isVisible() )
            {
                visible_param_count++;
            }
        }
        if (param_count > 0) {    // Список параметров для вывода в Excel
            data_parameters = CreateVariantArray(visible_param_count, 1);
            for (int i=0; i <= param_count-1; i++)
            {
                if ( !UserParams[i]->isVisible() )
                {
                    continue;
                }

                if (UserParams[i]->type != "separator")
                {
                    data_parameters.PutElement(UserParams[i]->getCaption() + ": " + UserParams[i]->getDisplay(), i+1, 1);
                }
                else
                {
                    data_parameters.PutElement("[" + UserParams[i]->getCaption() + "]", i+1, 1);
                }
            }
        }

        // Вывод данных на лист Excel
        Variant range_title = msexcel.WriteToCell(Worksheet1, param_excel.title_label , 1, 1);
        Variant range_createtime = msexcel.WriteToCell(Worksheet1, "По состоянию на: " + DateTime.DateTimeString(), 2, 1);
        Variant range_parameters;
        if (param_count > 0)
        {
            range_parameters = msexcel.WriteTable(Worksheet1, data_parameters, 3, 1);
        }

        Variant range_tablehead = msexcel.WriteTable(Worksheet1, data_head, 3 + visible_param_count, 1);
        Variant range_tablebody = msexcel.WriteTable(Worksheet1, data_body, 4 + visible_param_count, 1, &df_body);

        msexcel.SetRangeFormat(range_tablehead, cf_head);
        msexcel.SetRangeFormat(range_tablebody, cf_body);
        msexcel.SetRangeFormat(range_title, cf_title);
        msexcel.SetRangeFormat(range_createtime, cf_createtime);
        if (param_count > 0)
        {
            msexcel.SetRangeFormat(range_parameters, cf_createtime);
        }


        Variant range_all = msexcel.GetRangeFromRange(range_tablehead, 1, 1, msexcel.GetRangeRowsCount(range_tablebody)+1, msexcel.GetRangeColumnsCount(range_tablebody));


        if (this->param_excel.title_height > 0)
        {
            msexcel.SetRowHeight(range_tablehead, this->param_excel.title_height);    // Задаем высоту заголовка таблицы
        }

        msexcel.SetAutoFilter(range_all);   // Включаем автофильтр
        msexcel.SetColumnsAutofit(range_all);  // Ширина ячеек по содержимому


        // Настройка заголовка (размер и тп)
        for (int i=0; i < ExcelFieldCount; i++)   // Заполнение если есть поля в ExcelFields
        {    //CELLFORMAT cf_cell;
            //cf_cell.bSetFontColor = true;
            //cf_cell.FontColor = clGreen;

            if (param_excel.Fields[i].bwraptext >= 0)
            {
                CELLFORMAT cf_cell;
                cf_cell.bWrapText = param_excel.Fields[i].bwraptext;
                msexcel.SetRangeFormat(range_tablehead, cf_cell, 1, i+1);
            }

            if (param_excel.Fields[i].width >= 0)
            {
                msexcel.SetColumnWidth(range_tablehead, i+1, param_excel.Fields[i].width);    // Задаем высоту заголовка таблицы
                //msexcel.SetColumnWidth(range_tablehead, 1);    // Задаем высоту заголовка таблицы
            }
        }

        //msexcel.SetRowsAutofit(range_tablehead);


        // Выводим текст sql-запроса на второй лист
        Variant Worksheet2 = msexcel.GetSheet(Workbook, 2);


        Variant range_sqltext;
        int PartMaxLength = 4000;  // 8 192  - максимальная длина строки в ячейке EXCEL
        int n = ceil( (float) _mainQueryText.Length() / PartMaxLength);
        for (int i = 1; i <= n; i++)
        {
            AnsiString sQueryPart = _mainQueryText.SubString(((i-1) * PartMaxLength) + 1, PartMaxLength);
            range_sqltext = msexcel.WriteToCell(Worksheet2, sQueryPart, i, 1);
            msexcel.SetRangeFormat(range_sqltext, cf_sql);
        }

        df_body.clear();

        if (DstFileName == "")
        {
            msexcel.SetVisible(Workbook);
        }
        else
        {
            msexcel.SaveDocument(Workbook, DstFileName);
            VarClear(Workbook);
            VarClear(Worksheet1);
            VarClear(Worksheet2);
            msexcel.CloseApplication();
            _resultFiles.push_back(DstFileName);
        }


        //if (ExportMode == EM_EXCEL_FILE) {
        //    msexcel.SaveAsDocument(Workbook, DstFileName);
        //    msexcel.CloseExcel();
        //} else {
            //Workbook.OlePropertySet("Name", "blabla");
        //    msexcel.SetVisibleExcel(true, true);
        //}
    }

    // Освобождение памяти
    VarClear(data_head);
    data_head = NULL;

    VarClear(data_body);
    data_body = NULL;

    CoUninitialize();

    if (fDone)
    {
        throw Exception("Прерывание.");
    }

}
*/

//---------------------------------------------------------------------------
// Заполнение Excel файла с использованием шаблона xlt
void __fastcall TDocumentWriter::ExportToExcelTemplate(const TExcelExportParams* excelExportParams, TDataSet* QueryTable, TDataSet* QueryFields)
{
    CoInitialize(NULL);

    String TemplateFullName = excelExportParams->templateFilename; // Абсолютный путь к файлу-шаблону

    // Открываем шаблон MS Excel
    MSExcelWorks msexcel;
    Variant Workbook;
    Variant Worksheet;

    try
    {
        msexcel.OpenApplication();
        Workbook = msexcel.OpenDocument(TemplateFullName);
        Worksheet = msexcel.GetSheet(Workbook, 1);
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
        String msg = "Ошибка при открытии файла-шаблона " + TemplateFullName + ".\nОбратитесь к системному администратору.";
        throw Exception(msg);
    }

    // Сначала делаем замену полей
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
            result.addResultFile(excelExportParams->resultFilename);
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

    CoUninitialize();
}



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

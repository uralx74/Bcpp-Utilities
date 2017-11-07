#include "WordUtil.h"
#include <cassert>

using namespace strtools;
using namespace vartools;
using namespace fstools;
//---------------------------------------------------------------------------
//
MergeTable::MergeTable()
{
    FieldsCount = 0;
    CurrentRecordIndex = 1;
    RecCount = 0;
    PagePerDocument = 500;
}

//---------------------------------------------------------------------------
// Усечение (на самом деле лишь меняет значение RecCount)
void __fastcall MergeTable::ShrinkRecords(int RecCount)
{
    if (RecCount <0)
    {
        this->RecCount = CurrentRecordIndex - 1;
    }
    else
    {
        this->RecCount = RecCount;
    }
}

//---------------------------------------------------------------------------
// Подготовка (создание) массива данных
void __fastcall MergeTable::PrepareFields(int ColCount)
{
    VariantClear(head);
    FieldsCount = ColCount;
    head = CreateVariantArray(1, ColCount);
}

//---------------------------------------------------------------------------
//
void __fastcall MergeTable::PrepareRecords(int RowCount)
{
    VariantClear(data);
    RecCount = RowCount;
    data = CreateVariantArray(RowCount, FieldsCount);
}

//---------------------------------------------------------------------------
//
void __fastcall MergeTable::AddField(int FieldIndex, const AnsiString &FieldName)
{
     head.PutElement(FieldName, 1, FieldIndex);
}

//---------------------------------------------------------------------------
//
void __fastcall MergeTable::PutRecord(const AnsiString &Value, int RecordIndex, int FieldIndex)
{
    data.PutElement(Value, RecordIndex, FieldIndex);

    if (RecCount < CurrentRecordIndex)
    {
        RecCount = CurrentRecordIndex;  // Возможно стоит сделать иначе для использования готовых массивов данных
    }
}

//---------------------------------------------------------------------------
// Вставка значения в текущую строку
void __fastcall MergeTable::PutRecord(int FieldIndex, const AnsiString &Value)
{
    data.PutElement(Value, CurrentRecordIndex, FieldIndex);

}

//---------------------------------------------------------------------------
// Очистка массива
void __fastcall MergeTable::Free()
{
    VariantClear(data);
    VariantClear(head);

}

//---------------------------------------------------------------------------
//
void __fastcall MergeTable::Next()
{
    CurrentRecordIndex++;
    // Для предотвращения возможности выхода за пределы массива
    //if (CurrentRecordIndex > VarArrayHighBound(data,1) {
    //      RedimVariantArray(data, RecCount, fields.size()+100);
    //}
}

//---------------------------------------------------------------------------
//
void __fastcall MergeTable::First()
{
    CurrentRecordIndex = 1;
}

//---------------------------------------------------------------------------
//
void __fastcall MSWordWorks::SetDisplayAlerts(bool flg)
{
    WordApp.OlePropertySet("DisplayAlerts", flg);
}


void __fastcall MSWordWorks::OptimizePerformance(bool flg)
{
    /*wdDoc.ActiveWindow.View.Type = wd.WdViewType.wdNormalView;
    wdApp.Options.Pagination = false;*/
    WordApp.OlePropertySet("ScreenUpdating", flg);
}

//---------------------------------------------------------------------------
// Отрыть документ Word
Variant __fastcall MSWordWorks::OpenWord()
{
	try
    {
        // Создание объекта Word.Application
        WordApp = CreateOleObject("Word.Application.8");

        // Ищем handle окна
        randomize();
        String OldTitle = WordApp.OlePropertyGet("Caption");
        String TempTitle = "Temp - " + IntToStr(random(1000000));
        WordApp.OlePropertySet("Caption", TempTitle.c_str());
        //Handle = FindWindow(NULL, TempTitle.c_str());
        Handle = FindWindow("OpusApp", TempTitle.c_str());
        WordApp.OlePropertySet("Caption", OldTitle.c_str());

        // Отключить режим показа предупреждений.
        WordApp.OlePropertySet("DisplayAlerts", false);

        // Отключить проверку грамматики для ускорения работы
		WordApp.OlePropertyGet("Options").OlePropertySet("CheckSpellingAsYouType", false);
        WordApp.OlePropertyGet("Options").OlePropertySet("CheckGrammarAsYouType", false);
        WordApp.OlePropertyGet("Options").OlePropertySet("CheckGrammarWithSpelling", false);

        // Отображение Word.Application
        WordApp.OlePropertySet("Visible", false);

        // Инициализация ссылки на документы
	  	_documents = WordApp.OlePropertyGet("Documents");

        return WordApp;
    }
    catch (Exception &e)
    {
       try
       {
           CloseApplication();
        }
        catch(...)
        {
        }
        throw Exception(e);
    }
}

/*//---------------------------------------------------------------------------
// Отрыть документ Word - оставлена для совместимости со старыми версиями
Variant __fastcall MSWordWorks::OpenWord(const String &DocumentFileName, bool fAsTemplate)
{
    Variant Document;
	try
    {
        // Создание объекта Word.Application
        WordApp = CreateOleObject("Word.Application.8");

        randomize();

        // Ищем handle окна
        WideString OldTitle = WordApp.OlePropertyGet("Caption");
        WideString TempTitle = "Temp - " + IntToStr(random(1000000));
        WordApp.OlePropertySet("Caption", TempTitle.c_str());
        //Handle = FindWindow(NULL, TempTitle.c_str());
        Handle = FindWindow("OpusApp", TempTitle.c_str());
        WordApp.OlePropertySet("Caption", OldTitle);


        // Отключить режим показа предупреждений.
        WordApp.OlePropertySet("DisplayAlerts", false);

        // Отключить проверку грамматики для ускорения работы
		WordApp.OlePropertyGet("Options").OlePropertySet("CheckSpellingAsYouType", false);
        WordApp.OlePropertyGet("Options").OlePropertySet("CheckGrammarAsYouType", false);
        WordApp.OlePropertyGet("Options").OlePropertySet("CheckGrammarWithSpelling", false);

        // Отображение Word.Application
        WordApp.OlePropertySet("Visible", false);

        // Инициализация ссылки на документы
	  	Documents = WordApp.OlePropertyGet("Documents");

        // Выбор документа из файла
        if (fAsTemplate)
        {
        	// Каждый вновь создаваемый документ получает индекс Item = 1
        	Documents.OleProcedure("Add", DocumentFileName, false,0);
        	Document = Documents.OleFunction("Item",1); // Доступ к документу
        }
        else
        {
        	Document = Documents.OleFunction("Open", DocumentFileName);
        }
        return Document;


    }
    catch (Exception &e)
    {
        try
        {
            if (!VarIsEmpty(Document))     // Если шаблон открыт
            {
                CloseDocument(Document);
            }
        }
        catch(...)
        {}
        try
        {
           CloseApplication();
        }
        catch(...)
        {
        }
        throw Exception(e);
    }
} */

//---------------------------------------------------------------------------
// Отрыть документ Word из файла
Variant __fastcall MSWordWorks::OpenDocument(const String &DocumentFileName, bool fAsTemplate)
{
    // Выбор документа из файла
    // У Ole-процедуры Open существует множество дополнительных параметров
    Variant document;
    if ( fAsTemplate )
    {
        // Открывается существующий документ
    	document = _documents.OleFunction("Open", DocumentFileName);
    }
    else
    {
        // Создается новый документ, если открываемый документ - шаблон
    	// Каждый вновь создаваемый документ получает индекс Item = 1
    	document = _documents.OleFunction("Add", DocumentFileName.c_str(), false, 0);
    }
    return document;
}

//---------------------------------------------------------------------------
// Получение документа по индексу
Variant __fastcall MSWordWorks::GetDocument(int DocIndex)
{
	try
    {
        if (DocIndex >= 0)
        {
            //int k = WordApp.OlePropertyGet("Documents").OlePropertyGet("Count");
	        return WordApp.OlePropertyGet("Documents").OleFunction("Item", DocIndex);
        }
        else
        {
	        return WordApp.OlePropertyGet("ActiveDocument");
        }
    }
    catch (...)
    {
    	return NULL;
    }
}

//---------------------------------------------------------------------------
//
Variant __fastcall MSWordWorks::GetPage(Variant Document, int PageIndex)
{
   	return Document.OlePropertyGet("Item", wdPropertyPages).OlePropertyGet("Value");
   	//return Document.OlePropertyGet("BuiltInDocumentProperties").OlePropertyGet("Item", wdPropertyPages).OlePropertyGet("Value");
}

//---------------------------------------------------------------------------
//
int MSWordWorks::GetCurrentPageNumber(Variant Document)
{
    int page = Document.OlePropertyGet("Selection").OlePropertyGet("Information", wdActiveEndPageNumber);
    return page;
}

//---------------------------------------------------------------------------
// Скрыть или показать документ
void MSWordWorks::SetVisible(bool fVisible)
{
	// Отображение Word.Application
	WordApp.OlePropertySet("Visible", fVisible);
}

//---------------------------------------------------------------------------
// Вставка текста в закладку (закладка заменяется изображением)
void __fastcall MSWordWorks::SetTextToBookmark(Variant Document, String BookmarkName, WideString Text)
{
    Variant Bookmark = Document.OlePropertyGet("Bookmarks").OleFunction("Item", (OleVariant)BookmarkName);
	Bookmark.OlePropertyGet("Range").OlePropertySet("Text", Text);
}

//---------------------------------------------------------------------------
// Вставка изображения в закладку (текст заменяется изображением)
Variant MSWordWorks::SetPictureToBookmark(Variant Document, String BookmarkName, String PictureFileName, int Width, int Height)
{
    Variant Bookmarks=Document.OlePropertyGet("Bookmarks");
    Variant Bookmark = Bookmarks.OleFunction("Item", (OleVariant)BookmarkName);
    Variant picture = Bookmark.OlePropertyGet("Range").OlePropertyGet("InlineShapes").OleFunction("AddPicture", PictureFileName, false, true);

    // Не проверено!
    Bookmark.OleProcedure("Delete");


    return picture;
    //vBookmark.OlePropertyGet("Range").OlePropertySet("Text","12");
    /*// Выбор Bookmark по имени
    Variant Bookmark = Bookmarks.OleFunction("Item", (OleVariant)BookmarkName);
    // Вставить изображение из файла
    Bookmark.OlePropertyGet("Range").OlePropertyGet("InlineShapes").OleProcedure("AddPicture", PictureFileName, false, true);
	//Второй вариант
	// Определение выбранного участка в документе
    //Variant Selection = WordApp.OlePropertyGet("Selection");
    //Variant Selection = Document.OlePropertyGet("Selection");

	// возможно нужно удалить!!!!
    /*Variant InlineShape, PictureShape;
    InlineShape =  Document.OlePropertyGet("InlineShapes").OleFunction("Item", 1);
    //InlineShape = Field.OlePropertyGet("Range").OlePropertyGet("InlineShapes").OleFunction("Item", i);
    */

}

//---------------------------------------------------------------------------
// Возвращает вектор с именами полей FormFields
std::vector<String> __fastcall MSWordWorks::GetFormFields(Variant Document)
{
    // Рабочий вариант, но без проверки на существования поля с именем FieldName
    //Variant Field = Document.OlePropertyGet("FormFields").OleFunction("Item", (OleVariant)FieldName);
    //Field.OlePropertyGet("Range").OlePropertySet("Text", Text);

    Variant FormFields = Document.OlePropertyGet("FormFields");
    Variant Field = Unassigned;
    int n = FormFields.OlePropertyGet("Count");

    std::vector<String> vFormFields;
    vFormFields.reserve(n);


    for (int i = 1; i <= n; i++)   // Заполняем вектор с именами полей
    {
        String Name = UpperCase(FormFields.OleFunction("Item", i).OlePropertyGet("Result"));
        vFormFields.push_back(Name);
        /*if (FieldName == Name) {
            Field = Document.OlePropertyGet("FormFields").OleFunction("Item", (OleVariant)FieldName);
        }*/
    }

    return vFormFields;
}

//---------------------------------------------------------------------------
// Вставка текста в поле (поле заменяется текстом) - быстрый вариант, без проверки существования поля
void __fastcall MSWordWorks::SetTextToFieldF(Variant Document, String FieldName, WideString Text)
{
    // Рабочий вариант, но без проверки на существования поля с именем FieldName
    Variant Field = Document.OlePropertyGet("FormFields").OleFunction("Item", (OleVariant)FieldName);
    Field.OlePropertyGet("Range").OlePropertySet("Text", Text);
}

//---------------------------------------------------------------------------
// Вставка текста в поле (поле заменяется текстом) - быстрый вариант, без проверки существования поля
void __fastcall MSWordWorks::SetTextToFieldF(Variant Document, int fieldIndex, WideString Text)
{
    // Рабочий вариант, но без проверки на существования поля с именем FieldName
    Variant Field = Document.OlePropertyGet("FormFields").OleFunction("Item", fieldIndex);
    Field.OlePropertyGet("Range").OlePropertySet("Text", Text);
}


//---------------------------------------------------------------------------
// Вставка текста в поле (поле заменяется текстом)
void __fastcall MSWordWorks::SetTextToField(Variant Document, String FieldName, WideString Text)
{
    Variant FormFields = Document.OlePropertyGet("FormFields");
    Variant Field = Unassigned;
    FieldName = UpperCase(FieldName);
    int n = FormFields.OlePropertyGet("Count");
    for (int i = 1; i <= n; i++)   // Цикл по каждому полю. Проверяем наличие поля FieldName
    {
        String Name = UpperCase(FormFields.OleFunction("Item", i).OlePropertyGet("Result"));
        if (FieldName == Name)
        {
            Field = Document.OlePropertyGet("FormFields").OleFunction("Item", (OleVariant)FieldName);
        }
    }

    if ( !Field.IsEmpty() )
    {
        Field.OlePropertyGet("Range").OlePropertySet("Text", Text);
    }



/*  Variant Fields = Document.OlePropertyGet("MailMerge").OlePropertyGet("Fields");
    int FieldsCount = Fields.OlePropertyGet("Count");
    Variant Field = Fields.OleFunction("Item", 1);
    Variant Code = Field.OlePropertyGet("Code");
    String Text = code.OlePropertyGet("Text");
    String Type = Field.OlePropertyGet("Type");
*/

}

/* Возвращает количество полей MailMerge */
int __fastcall MSWordWorks::GetMailMergeFieldCount(Variant Document)
{
    Variant Fields = Document.OlePropertyGet("MailMerge").OlePropertyGet("Fields");
    return Fields.OlePropertyGet("Count");
}

/* Задает свойства документа */
void __fastcall MSWordWorks::SetBuiltInProperty(Variant Document, int property, const String& value)
{
    //Variant dp = Document.OlePropertyGet("BuiltInDocumentProperties");
    Variant p = Document.OlePropertyGet("BuiltInDocumentProperties").OlePropertyGet("Item", property);
    p.OlePropertySet("Value", (OleVariant)value);

/*typedef enum WdBuiltInProperty
{
wdPropertyTitle = 1,
wdPropertySubject = 2,
wdPropertyAuthor = 3,
wdPropertyKeywords = 4,
wdPropertyComments = 5,
wdPropertyTemplate = 6,
wdPropertyLastAuthor = 7,
wdPropertyRevision = 8,
wdPropertyAppName = 9,
wdPropertyTimeLastPrinted = 10,
wdPropertyTimeCreated = 11,
wdPropertyTimeLastSaved = 12,
wdPropertyVBATotalEdit = 13,
wdPropertyPages = 14,
wdPropertyWords = 15,
wdPropertyCharacters = 16,
wdPropertySecurity = 17,
wdPropertyCategory = 18,
wdPropertyFormat = 19,
wdPropertyManager = 20,
wdPropertyCompany = 21,
wdPropertyBytes = 22,
wdPropertyLines = 23,
wdPropertyParas = 24,
wdPropertySlides = 25,
wdPropertyNotes = 26,
wdPropertyHiddenSlides = 27,
wdPropertyMMClips = 28,
wdPropertyHyperlinkBase = 29,
wdPropertyCharsWSpaces = 30
} WdBuiltInProperty;*/

    //CustomDocumentProperties = Document.OlePropertyGet("CustomDocumentProperties");
}


void __fastcall MSWordWorks::SetVariableValue(Variant Document, const String& variableName, const String& value)
{
    Variant variables = Document.OlePropertyGet("Variables");
    Variant variable = variables.OleFunction("Item", (OleVariant)variableName);
    variable.OlePropertySet("Value", (OleVariant)value);
}


/* Преобразует InlineShape в плавающий Shape
   zOrder - положение относительно текста
   4 - перед текстом
   5 - за текстом
*/
Variant __fastcall MSWordWorks::ConverInlineShapeToShape(Variant inlineShape, int zOrder)
{
    // Конвертируем в Shape
    Variant shape = inlineShape.OleFunction("ConvertToShape");

    // Расположение изображения перед текстом
    shape.OleFunction("ZOrder", zOrder);  // 4 - msoBringInFrontOfText

    // Сохранять пропорции
    shape.OlePropertySet("LockAspectRatio", true);

    // Выравнивание по вертикали по центру строки
    shape.OlePropertySet("Top", -999995);    // wdShapeCenter = -999995

    // Выравнивание относительно
    shape.OlePropertySet("RelativeVerticalPosition", 3); //  wdRelativeVerticalPositionLine = 3

    //
    //WrapFormat.AllowOverlap = True

    return shape;
}

/* Задает размеры плавающего shape */
void __fastcall MSWordWorks::SetShapeSize(Variant shape, int width, int height)
{
    // Устанавливаем размер изображения
    //if (Width != 0 && Height != 0) {
    	shape.OlePropertySet("Width", width);
    	shape.OlePropertySet("Height", height);
    //}
}

/* Задает положение плавающего Shape
   Не проверена!
*/
void __fastcall MSWordWorks::SetShapePos(Variant shape, int x, int y)
{
    // Устанавливаем расположение на листе
    shape.OleProcedure("IncrementLeft", x);
    shape.OleProcedure("IncrementTop", y);
}

/* Устанавливает картинку на место Range
   Проверена*/
Variant __fastcall MSWordWorks::SetPictureToRange(Variant Document, Variant Range, String PictureFileName)
{
    try
    {
        Variant InlineShapes = Range.OlePropertyGet("InlineShapes");
        Variant InlineShape = InlineShapes.OleFunction("AddPicture", PictureFileName.c_str(), false, true);
        return InlineShape;
    }
    catch (Exception &e)
    {
        throw Exception(e);
    }
}

/* Переработать! Так как Field не имеет свойства Range */
Variant __fastcall MSWordWorks::SetPictureToFormField(Variant Document, Variant Field, String PictureFileName, int Width, int Height)
{
    try
    {
        Variant InlineShapes = Field.OlePropertyGet("Range").OlePropertyGet("InlineShapes");
        Variant InlineShape = InlineShapes.OleFunction("AddPicture", PictureFileName.c_str(), false, true);
        // OleFunction("AddPictureBullet", "test.bmp");

        // Не проверено!
        Field.OleProcedure("Delete");

        if (Width != 0 || Height != 0)
        {
            InlineShape.OlePropertySet("Width", Width);
            InlineShape.OlePropertySet("Height", Height);
        }

        return InlineShape;
    }
    catch (Exception &e)
    {
        throw(Exception("Exception has occurred.\nFile \"" + PictureFileName + "\" not found."));
    }
}



/*//---------------------------------------------------------------------------
// Вставка рисунка в поле (поле остается)
Variant __fastcall MSWordWorks::SetPictureToField(Variant Document, int fieldIndex, String PictureFileName, int Width, int Height)
{
    Variant Field = Document.OlePropertyGet("Fields").OleFunction("Item", fieldIndex);
    return SetPictureToField(Document, Field, PictureFileName, Width, Height);
}

//---------------------------------------------------------------------------
// Вставка рисунка в поле (поле остается)
Variant __fastcall MSWordWorks::SetPictureToField(Variant Document, String FieldName, String PictureFileName, int Width, int Height)
{
    Variant Field = Document.OlePropertyGet("Fields").OleFunction("Item", (OleVariant)FieldName);
    return SetPictureToField(Document, Field, PictureFileName, Width, Height);
} */

//---------------------------------------------------------------------------
// Вставка рисунка в поле (поле остается)
Variant __fastcall MSWordWorks::SetPictureToFormField(Variant Document, int fieldIndex, String PictureFileName, int Width, int Height)
{
    Variant Field = Document.OlePropertyGet("FormFields").OleFunction("Item", fieldIndex);
    return SetPictureToFormField(Document, Field, PictureFileName, Width, Height);
}

//---------------------------------------------------------------------------
// Вставка рисунка в поле (поле остается)
Variant __fastcall MSWordWorks::SetPictureToFormField(Variant Document, String FieldName, String PictureFileName, int Width, int Height)
{
    Variant Field = Document.OlePropertyGet("FormFields").OleFunction("Item", (OleVariant)FieldName);
    return SetPictureToFormField(Document, Field, PictureFileName, Width, Height);
}

//---------------------------------------------------------------------------
// Поиск текста с заменой
void __fastcall MSWordWorks::FindTextForReplace(Variant document, String Text, String ReplaceText, bool fReg)
{
        document.OleProcedure("Activate");
        // Определение выбранного участка в документе
        //Variant Selection = Document.OlePropertyGet("Selection");
        Variant Selection = WordApp.OlePropertyGet("Selection");

        // Поиск текста по всему документу
		Variant Find = Selection.OlePropertyGet("Find");
        Find.OleProcedure("Execute", Text.c_str()/*Текст, который будем менять*/, fReg/*учитывать регистрe*/,
        	false/*Только полное слово*/,false/*Учитывать универсальные символы*/,false/*Флажок Произносится как*/,
        	false/*Флажок Все словоформы*/,true/*Искать вперед*/,1/*Активация кнопки Найти далее*/,
        	false/* Задание формата */, ReplaceText.c_str()/*На что заменить*/,2/*Заменить все*/);   // Этот вариант работает

        /*
        Find.OleProcedure("ClearFormatting");                                         // Этот вариант НЕ работает, надо разбираться
        Find.OlePropertyGet("Replacement").OleProcedure("ClearFormatting");
        Find.OlePropertySet("Text",Text);
        Find.OlePropertyGet("Replacement").OlePropertySet(Text,ReplaceText);
        Find.OlePropertySet("Forward",True);
        Find.OlePropertySet("Wrap",1);
        Find.OlePropertySet("Format",False);
        Find.OlePropertySet("MatchCase",False);
        Find.OlePropertySet("MatchWholeWord",False);
        Find.OlePropertySet("MatchWildcards",False);
        Find.OlePropertySet("MatchSoundsLike",False);
        Find.OlePropertySet("MatchAllWordForms",False);
        Find.OleProcedure("Execute",2);   /**/
}

//---------------------------------------------------------------------------
// (недоделано!) копирование страницы по номеру   (недоделано!)(недоделано!)(недоделано!)(недоделано!)(недоделано!)
Variant MSWordWorks::CopyPage(int PageNumber)
{
	//Variant Selection = Document.OlePropertyGet("Selection");
	// Переместить курсор
    WordApp.OlePropertyGet("Selection").OleFunction("WholeStory");
    //CurrentSelection.OleProcedure("MoveUp", wdLine, 1, wdExtend);
    //Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend
    return WordApp.OlePropertyGet("Selection").OleFunction("Copy");
}

//---------------------------------------------------------------------------
//  Вставка страницы по номеру(недоделано!)(недоделано!)(недоделано!)(недоделано!)(недоделано!)(недоделано!)(недоделано!)
void __fastcall MSWordWorks::PastePage(Variant Document, int PageNumber)
{
    Variant Selection = WordApp.OlePropertyGet("Selection");
    //Variant Selection = Document.OlePropertyGet("Selection");
	// Переместить курсор в конец документа
    Selection.OleProcedure("EndKey", wdStory);
	//CurrentSelection.OleProcedure("InsertNewPage");
    Selection.OleFunction("Paste");
    //Selection.OleFunction("PasteAndFormat", 0);
    //Selection.OleFunction("PasteAndFormat", Page);
}

//---------------------------------------------------------------------------
//
void __fastcall MSWordWorks::InsertFile(Variant Document, AnsiString FileName)
{
    // Переместить курсор в конец документа
    Variant Selection = Document.OlePropertyGet("Selection");
    Selection.OleProcedure("EndKey", wdStory);
    Selection.OleProcedure("InsertFile", FileName.c_str(), "", false, false);
}

//---------------------------------------------------------------------------
// Переместить курсор в начало документа (недоделано!)
void __fastcall  MSWordWorks::MoveUpCursor(Variant Document)
{
	//Variant Selection = Document.OlePropertyGet("Selection");
	// Переместить курсор
    Variant Selection = Document.OlePropertyGet("Selection");
	Selection.OleProcedure("MoveUp", 7, 1);
}

//---------------------------------------------------------------------------
// Сохранить в файл
void __fastcall MSWordWorks::SaveAsDocument(Variant Document, String FileName/*, bool fAddToRecentFiles*/)
{
    // Сохранение документа в файл
    // У данной Ole-процедуры множество дополнительных параметров
    // Также существует аналог этой процедуры - SaveAs2
    Document.OleProcedure("SaveAs", FileName.c_str());
}

//---------------------------------------------------------------------------
// Возвращает имя файла документа
String __fastcall MSWordWorks::GetDocumentFilename(Variant Document)
{
    return Document.OlePropertyGet("Fullname");
}

//---------------------------------------------------------------------------
//
void __fastcall MSWordWorks::SetActiveDocument(Variant Document)
{
	// Активация документа
    Document.OleProcedure("Activate");
}

//---------------------------------------------------------------------------
// Создать таблицу
Variant MSWordWorks::CreateTable(Variant Document, int nCols, int nRows)
{
    Variant range = Document.OleFunction("Range");
	// Создане таблицы в области Range
    return Document.OlePropertyGet("Tables").OleFunction("Add", range, (OleVariant) nCols, (OleVariant) nRows);

	// Выбор существующей таблицы
	//Table = Tables.OleFunction("Item", 1);
	//RowCount = Table.OlePropertyGet("Rows").OlePropertyGet("Count");
	//ColCount = Table.OlePropertyGet("Columns").OlePropertyGet("Count");
}

// Получить таблицу по индексу
Variant MSWordWorks::GetTableByIndex(Variant Document, int index)
{
    return Document.OlePropertyGet("Tables").OleFunction("Item", index);
}

//---------------------------------------------------------------------------
// Перейти к закладке
void __fastcall MSWordWorks::GoToBookmark(Variant Document, String BookmarkName)
{
    Document.OleFunction("Range").OleProcedure("GoTo", wdGoToBookmark, 0, 0, WideString(BookmarkName));
//    Document.OlePropertyGet("Selection").OleProcedure("GoTo",(int)-1, 0, 0, WideString(BookmarkName));
}

//---------------------------------------------------------------------------
// Перейти к тексту
void __fastcall MSWordWorks::GoToText(Variant Document, String Text, bool fReg, bool fWord)
{
    // Определение выбранного участка в документе
    //Variant Selection = WordApp.OlePropertyGet("Selection");
    //CurrentSelection = WordApp.OlePropertyGet("Selection");

    Variant Selection = Document.OleFunction("Range");

    // Поиск текста по всему документу
	Variant Find = Selection.OlePropertyGet("Find");
    Find.OleProcedure("Execute", Text/*Текст, который будем менять*/, fReg/*учитывать регистрe*/,
        fWord/*Только полное слово*/,false/*Учитывать универсальные символы*/,false/*Флажок Произносится как*/,
        false/*Флажок Все словоформы*/,true/*Искать вперед*/,1/*Активация кнопки Найти далее*/,
        false/* Задание формата */, 0/*На что заменить*/,0/*Заменить все*/);   // Этот вариант работает

    //СurrentSelection.OleProcedure("GoTo",(int)-1, 0, 0, WideString(BookmarkName));
}

//---------------------------------------------------------------------------
// Вставить изображение
void __fastcall MSWordWorks::InsertPicture(Variant Document, String PictureFileName, int Width, int Height)
{
     // Вставить изображение из файла в позицию CurrentSelection
     Variant Selection = Document.OleFunction("Range");
     Selection.OlePropertyGet("InlineShapes").OleProcedure("AddPicture", "C:\\_project\\InsertPicToWord\\tmp\\podpis.bmp", false, true);
}

//---------------------------------------------------------------------------
// Вставить текст
void __fastcall MSWordWorks::InsertText(Variant Document, WideString Text)
{
     // Вставить текст в позицию CurrentSelection
     Variant Selection = Document.OleFunction("Range");
     Selection.OleProcedure("TypeText", Text);
}

//---------------------------------------------------------------------------
// Вставить из Clipboard (необходимо следить за содержимым Clipboard)
bool MSWordWorks::PasteFromClipboard()
{
     // Вставить из буффера
	WordApp.OlePropertyGet("Selection").OleFunction("Paste");
    return true;
}

//---------------------------------------------------------------------------
//  Закрытие приложения Word с закрытием всех открытых документов
void __fastcall MSWordWorks::CloseApplication()
{
	// Закрытие приложения Word (с запосом на сохранение документа)
    if (!WordApp.IsEmpty())
    {
        Variant document;
        while (_documents.OlePropertyGet("Count") > 0)
        {
            document = WordApp.OlePropertyGet("ActiveDocument");
            document.OleFunction("Close", false);
        }
	    WordApp.OleProcedure("Quit");
    }
}

//---------------------------------------------------------------------------
// Закрытие документа
void __fastcall MSWordWorks::CloseDocument(Variant Document, bool fCloseAppIfNoDoc)
{
	Document.OleFunction("Close", false);

    if (fCloseAppIfNoDoc && _documents.OlePropertyGet("Count") == 0)
    {
        CloseApplication();
    }
}

//---------------------------------------------------------------------------
// Определение количества страниц
int MSWordWorks::GetPagesCount(Variant Document)
{
	return Document.OlePropertyGet("BuiltInDocumentProperties").OlePropertyGet("Item", wdPropertyPages).OlePropertyGet("Value");
}

//---------------------------------------------------------------------------
// Слияние с объектом MERGETABLE в файл
std::vector<String> __fastcall MSWordWorks::MergeDocumentToFiles(Variant TemplateDocument, MERGETABLE &md)
{
    //md.TemplateDocument = OpenWord(md.TemplateFileName, true);

    int nFiles;
    if (md.PagePerDocument <= 0)          // Расчитываем кол-во результирующих файлов
    {
        md.PagePerDocument = md.RecCount;
        nFiles = 1;
    }
    else
    {
        nFiles = ceil((double)md.RecCount/(double)md.PagePerDocument);
    }

    int nPad = IntToStr(nFiles).Length();  // Кол-во знаков в индексе в имени файлов

    std::vector<String> vFiles;
    vFiles.reserve(nFiles);

    int FileIndex = 0;
    for (int i = 1; i <= md.RecCount; i = i + md.PagePerDocument)
    {
        FileIndex++;
        //AnsiString filename = md.ResultFileNamePrefix + str_pad(IntToStr(FileIndex), nPad, "0", STR_PAD_LEFT) + ".doc";
        String counterStr = StrPadL(IntToStr(FileIndex), nPad, "0");


        // Добавляем к имени файла порядковый номер
        //String filename = ReplaceField(md.resultFilename, "[:counter]", counterStr);

        /* 2017-11-07
        TReplaceFlags replaceflags = TReplaceFlags() << rfReplaceAll << rfIgnoreCase;
        String filename = StringReplace(md.resultFilename, "[:counter]", counterStr, replaceflags);
        */
        String filename = ExtractFilePath(md.resultFilename) + ExtractFileName(md.resultFilename) + counterStr + ExtractFileExt(md.resultFilename);   // 2017-11-07


        Variant ResultDocument = MergeDocument(TemplateDocument, md, i);

        // Сохраняем документ, без помещения его в список последних файлов (AddToRecentFiles = false - 6й параметр)
        try
        {
            // Преобразуем все поля в значения
            UnlinkFields(ResultDocument.OlePropertyGet("Fields"));  // В процессе тестирования с 2017-03-31
            ResultDocument.OleProcedure("SaveAs", filename.c_str(), 0, false, "", false);
        }
        catch (Exception &e)
        {
            throw Exception("Ошибка при сохранении в файл\n" + filename);
        }
        vFiles.push_back(GetDocumentFilename(ResultDocument));   // 2017-11-07 Проверить
        CloseDocument(ResultDocument);

    }
    //CloseDocument(TemplateDocument);

    return vFiles;
    //return FileIndex;
}

//---------------------------------------------------------------------------
// Слияние с объектом MERGETABLE в объект Word - Document
Variant __fastcall MSWordWorks::MergeDocument(Variant TemplateDocument, MERGETABLE &md, int StartIndex)
{
    randomize();
    AnsiString TmpFileName = "ds" + IntToStr(random(100000000)) + ".html";

    TmpFileName = GetTempPathEx() + TmpFileName;

    //int ArrayRowsCount = md.RecCount;

    //int ArrayRowsCount = VarArrayHighBound(ArrayData.data, 1) - VarArrayLowBound(ArrayData.data, 1)+1;
    //int ArrayColsCount = VarArrayHighBound(md.data, 2) - VarArrayLowBound(md.data, 2)+1;

    //int ArrayRowsCount = ArrayData.data2[0].size();
    //int ArrayColsCount1 = md.FieldsCount;

    //int startCol = VarArrayLowBound(ArrayData.data, 2);
    int startCol = 1;
    int LastRecordIndex;

    if (StartIndex <= 0 )
    {
        StartIndex = 1;
    }
    int PagesCount = StartIndex + md.PagePerDocument-1;
    if (PagesCount > md.RecCount)
    {
        PagesCount = md.RecCount;
    }

     ofstream out(TmpFileName.c_str());

    // Заголовок HTML
    out<<"<html>\n";
    out<<"<head><META http-equiv=""content-type"" content=""text/html; charset=windows-1251""></head>\n";
    out<<"<body>\n<table>";

    // Заголовки
    out<<"<tr>";
    for (int j = startCol; j <= md.FieldsCount; j++)
    {
        AnsiString s ="<td>"  + md.head.GetElement(1, j) + "</td>";
        out<< s.c_str();
    }
    out<<"</tr>";
    out<<"\n";

    // Тело
    for (int i = StartIndex; i <= PagesCount; i++)
    {
        out<<"<tr>";
        for (int j = startCol; j <= md.FieldsCount; j++)
        {
            //AnsiString s ="<td>#@#sep_"  + md.data.GetElement(i, j) + "</td>";  // Закомментировано 2016-07-21
            AnsiString s ="<td>"  + md.data.GetElement(i, j) + "</td>";
            out<< s.c_str();
        }
        out<<"</tr>";
        out<<"\n";
    }
    out<<"</table></body></html>";

    out.close();

    Variant Document = MergeDocumentFromFile(TemplateDocument, TmpFileName, 1, PagesCount);

    //FindTextForReplace("#@#sep_", ""); // Закомментировано 2016-07-21
    // Заменяет разрыв раздела на разрыв страницы
    // это необходимо, чтобы пользователи могли печатать диапазон страниц.
    // Иначе в документа получается множество листов и одна страница 
    FindTextForReplace(Document, "^b", "^m");

    for (int i = 0; i < 5; i ++)       // Delete temporary file
    {
        if (remove(TmpFileName.c_str()) == 0)
        {
            break;
        }
        Sleep(500);
    }

    return Document;
}

/*
//---------------------------------------------------------------------------
// Слияние с объектом MERGETABLE в объект Word - Document
Variant __fastcall MSWordWorks::MergeDocument(Variant TemplateDocument, MERGETABLE &md, int FirstRecordIndex)
{
    randomize();
    AnsiString TmpFileName = "ds" + IntToStr(random(100000000)) + ".html";
    TmpFileName = GetTempPath() + TmpFileName;

    int ArrayRowsCount = md.RecCount;

    //int ArrayRowsCount = VarArrayHighBound(ArrayData.data, 1) - VarArrayLowBound(ArrayData.data, 1)+1;
    int ArrayColsCount = VarArrayHighBound(md.data, 2) - VarArrayLowBound(md.data, 2)+1;

    //int ArrayRowsCount = ArrayData.data2[0].size();
    int ArrayColsCount1 = md.fields.size();

    //int ArrayRowsCount = VarArrayHighBound(ArrayData.data, 1) - VarArrayLowBound(ArrayData.data, 1)+1;
    //int ArrayColsCount = VarArrayHighBound(ArrayData.data, 2) - VarArrayLowBound(ArrayData.data, 2)+1;

    //int startCol = VarArrayLowBound(ArrayData.data, 2);
    int startCol = 1;
    int LastRecordIndex;

    if (FirstRecordIndex <= 0 )
        FirstRecordIndex = 1;

    //int delta = ArrayRowsCount - FirstRecordIndex + 1;
    int PagesCount;

    if (md.PagePerDocument <= 0)
    {
        //PagePerDocument = ArrayRowsCount - FirstRecordIndex + 1;
        PagesCount = md.RecCount;
    } else if (md.PagePerDocument >= delta) {
        //PagePerDocument = delta;
        PagesCount = md.PagePerDocument;
    }
    LastRecordIndex = FirstRecordIndex + PagePerDocument - 1;

    int FieldsCount_Document = TemplateDocument.OlePropertyGet("Fields").OlePropertyGet("Count");

    ofstream out(TmpFileName.c_str());

    out<<"<html>\n";
    out<<"<head><META http-equiv=""content-type"" content=""text/html; charset=windows-1251""></head>\n";
    out<<"<body>\n<table>";


    // Заголовки
    out<<"<tr>";
    for (std::map<AnsiString, int>::iterator field = ArrayData.fields.begin(); field != ArrayData.fields.end(); ++field)
    {
        AnsiString s ="<td>"  + field->first + "</td>";
        out<< s.c_str();
    }
    out<<"</tr>";
    out<<"\n";

    // Тело
    for (int i = FirstRecordIndex; i <= LastRecordIndex; i++)
    {
        out<<"<tr>";
        for (int j = startCol; j <= ArrayColsCount; j++)
        {

            //AnsiString s ="<td>"  + ArrayData.data2[i][j] + "</td>";
            AnsiString s ="<td>"  + ArrayData.data.GetElement(i, j) + "</td>";
            out<< s.c_str();
        }
        out<<"</tr>";
        out<<"\n";
    }
    out<<"</table></body></html>";

    out.close();


    Variant Document = MergeDocumentFromFile(TemplateDocument, TmpFileName, 1, PagePerDocument);

    for (int i = 0; i < 5; i ++) {      // Delete temporary file
        if (remove(TmpFileName.c_str()) == 0)
            break;
        Sleep(500);
    }

    return Document;
}  */

//---------------------------------------------------------------------------
//
Variant __fastcall MSWordWorks::MergeDocument(Variant TemplateDocument, const Variant &ArrayData, int FirstRecordIndex, int PagePerDocument, int titleRowIndex)
{
    randomize();
    AnsiString TmpFileName = "ds" + IntToStr(random(100000000)) + ".html";
    TmpFileName = GetTempPathEx() + TmpFileName;


    int ArrayRowsCount = VarArrayHighBound(ArrayData, 1) - VarArrayLowBound(ArrayData, 1)+1;
    int ArrayColsCount = VarArrayHighBound(ArrayData, 2) - VarArrayLowBound(ArrayData, 2)+1;

    //int startRow = VarArrayLowBound(*ArrayData, 1);
    int startCol = VarArrayLowBound(ArrayData, 2);
    int LastRecordIndex;


    if (titleRowIndex <= 0)
    {
        titleRowIndex = VarArrayLowBound(ArrayData, 1);
    }

    if (FirstRecordIndex <= 0 )
    {
        FirstRecordIndex = titleRowIndex + 1;
    }

    int delta = ArrayRowsCount - FirstRecordIndex + 1;
    if (PagePerDocument <= 0)
    {
        PagePerDocument = ArrayRowsCount - FirstRecordIndex + 1;
    } 
    else if (PagePerDocument >= delta) 
    {
        PagePerDocument = delta;
    }

    LastRecordIndex = FirstRecordIndex + PagePerDocument - 1;

/*    if (RecordCount <= 0)
    {
        LastRecordIndex = VarArrayHighBound(ArrayData, 1);
    }
    else
    {
        LastRecordIndex = FirstRecordIndex + RecordCount;
        if (LastRecordIndex > ArrayRowsCount) {
            LastRecordIndex = ArrayRowsCount;
            RecordCount =
        }
    }     */

    //int FieldsCount_Document = TemplateDocument.OlePropertyGet("Fields").OlePropertyGet("Count");

    //int FieldsCount = FieldsCount_Document < ArrayColsCount? FieldsCount_Document : ArrayColsCount;

    ofstream out(TmpFileName.c_str());

    out<<"<html>\n";
    out<<"<head><META http-equiv=""content-type"" content=""text/html; charset=windows-1251""></head>\n";
    out<<"<body>\n<table>";

    // Заголовки
    out<<"<tr>";
    for (int j = startCol; j <= ArrayColsCount; j++)
    {
        AnsiString s ="<td>"  + ArrayData.GetElement(titleRowIndex, j) + "</td>";
        out<< s.c_str();
    }
    out<<"</tr>";
    out<<"\n";


    for (int i = FirstRecordIndex; i <= LastRecordIndex; i++)
    {
        out<<"<tr>";
        for (int j = startCol; j <= ArrayColsCount; j++)
        {
            // AnsiString s ="<td>#@#sep_"  + ArrayData.GetElement(i, j) + "</td>";   // Закомментированно 2016-07-21
            AnsiString s ="<td>#@#sep_"  + ArrayData.GetElement(i, j) + "</td>";
            out<< s.c_str();
        }
        out<<"</tr>";
        out<<"\n";
    }
    out<<"</table></body></html>";

    out.close();

    Variant Document = MergeDocumentFromFile(TemplateDocument, TmpFileName, 1, PagePerDocument);
    //FindTextForReplace("#@#sep_", "");  // Закомментировано 2016-07-21

    for (int i = 0; i < 5; i ++)       // Delete temporary file
    {
        if (remove(TmpFileName.c_str()) == 0)
        {
            break;
        }
        Sleep(500);
    }

    return Document;
}

//---------------------------------------------------------------------------
// Слияние из готового файла с данными (html)
Variant __fastcall MSWordWorks::MergeDocumentFromFile(Variant TemplateDocument, AnsiString DatasetFileName, int FirstRecordIndex, int PagePerDocument)
{
    Variant ResultDocument;
    Variant MailMerge;
    Variant PasswordDocument, PasswordTemplate, WritePasswordDocument, WritePasswordTemplate, SQLStatement, SQLStatement1;
    PasswordDocument = "";
    PasswordTemplate = "";
    WritePasswordDocument = "";
    WritePasswordTemplate = "";
    SQLStatement = "";
    //SQLStatement = "SELECT * FROM [Лист1$]";
    SQLStatement = "SELECT * FROM `Table`";
    SQLStatement1 = "SELECT * FROM `Table`";
    //String chemin, texte;
    //texte = "test";

    MailMerge = TemplateDocument.OlePropertyGet("MailMerge");
    MailMerge.OlePropertySet("MainDocumentType", 0);    // wdFormLetters = 0
    //MailMerge.OleProcedure("OpenDataSource", chemin.c_str(), 1, true, true, false, false, PasswordDocument, PasswordTemplate, false, WritePasswordDocument, WritePasswordTemplate, texte.c_str(), SQLStatement, SQLStatement1, false);
    // 2017-11-07 MailMerge.OleProcedure("OpenDataSource", DatasetFileName.c_str(), 0, false, false, true, false, PasswordDocument, PasswordTemplate, false, WritePasswordDocument, WritePasswordTemplate, texte.c_str(), SQLStatement, SQLStatement1, false);
    MailMerge.OleProcedure("OpenDataSource", DatasetFileName.c_str(), 0, false, false, true, false, PasswordDocument, PasswordTemplate, false, WritePasswordDocument, WritePasswordTemplate, "", SQLStatement, SQLStatement1, false);
    MailMerge.OlePropertySet("Destination", 0);
    MailMerge.OlePropertySet("SuppressBlankLines", 0);

    int LastRecordIndex;
    if (FirstRecordIndex <= 0)
    {
        FirstRecordIndex = 1;           //wdDefaultFirstRecord = 1
    }

    if (PagePerDocument  <= 0)
    {
        LastRecordIndex = 0xFFFFFFF0;   //wdDefaultLastRecord = 0xFFFFFFF0
    }
    else
    {
        LastRecordIndex = FirstRecordIndex + PagePerDocument - 1;
    }

    MailMerge.OlePropertyGet("DataSource").OlePropertySet("FirstRecord", FirstRecordIndex);
    MailMerge.OlePropertyGet("DataSource").OlePropertySet("LastRecord", LastRecordIndex);

    // Выполняем слияние
    MailMerge.OleProcedure("Execute", false);

    // Освобождаем набор данных
    MailMerge.OlePropertySet("MainDocumentType", 0xFFFFFFFF);    // wdNotAMergeDocument = 0xFFFFFFFF


    // Возвращаем новый документ
    WordApp = MailMerge.OlePropertyGet("Application");
    return WordApp.OlePropertyGet("Documents").OleFunction("Item", 1);


    /*MailMerge.ExecFunction("OpenDataSource") <<XLSFileName               // Name
                                    <<0                         // Format
                                    <<false                     // ConfirmConversions
                                    <<false                     // ReadOnly
                                    <<true                      // LinkToSource
                                    <<false                     // AddToRecentFiles
                                    <<EmptStr                   // PasswordDocument
                                    <<EmptStr                   // PasswordTemplate
                                    <<false                     // Revert
                                    <<EmptStr                   // WritePasswordDocument
                                    <<EmptStr                   // WritePasswordTemplate
                                    <<"Entire Spreadsheet"      // Connection
                                    <<EmptStr                   // SQLStatement
                                    <<EmptStr                   // SQLStatement1
                                    <<false                     // OpenExclusive
                                    <<8                         // SubType
         );*/

}

//---------------------------------------------------------------------------
// Слияние из готового файла с данными (html)
std::vector<String> MSWordWorks::ExportToWordFields(TDataSet* dataSet, Variant Document, const String& resultPath, int PagePerDocument)
{
    try
    {
        int recordCount = dataSet->RecordCount - dataSet->RecNo + 1;

        Variant fields = Document.OleFunction("Range").OlePropertyGet("Fields");
        std::vector<TFieldLink> links = assignDataSetToRangeFields(fields, DFT_MERGEFIELD, dataSet, "");

        // Вместо не найденных полей вставляем замещающий текст
        int linkedFieldCount = 0;
        for (std::vector<TFieldLink>::iterator it = links.begin(); it != links.end(); it++)
        {
            if ( it->dsField == NULL )
            {
                it->docField.OleFunction("Select");
                Variant vSelection = WordApp.OlePropertyGet("Selection");
                Variant vrange = vSelection.OlePropertyGet("Range");
                //vSelection.OleProcedure("TypeText","hello");
                vrange.OlePropertySet("Text", ("NA_" + it->docFieldName).c_str());
                vrange.OlePropertySet("HighlightColorIndex", 7); // wdYellow = 7 ;
            }
            /*else
            {
                linkedFieldCount++;
            }*/
            //it->documentField.OlePropertyGet("Result").OlePropertySet("Text", it->datasetField->AsString.c_str());
        }

        // Ассоциируем заново, так как могли быть вложеные поля
        links = assignDataSetToRangeFields(fields, DFT_MERGEFIELD, dataSet, "");
        linkedFieldCount = links.size();

        // Подготавливаем хранилище для html-таблицы
        MERGETABLE mergetable;
        mergetable.resultFilename = resultPath;
        mergetable.PrepareFields(linkedFieldCount);
        mergetable.PrepareRecords(recordCount);
        mergetable.PagePerDocument = PagePerDocument;

         // Создаем структуру для html
        int i = 1;
        for (std::vector<TFieldLink>::iterator it = links.begin(); it != links.end(); it++)
        {
            if ( it->dsField == NULL )
            {
                continue;
            }

            mergetable.AddField(i++, it->dsFieldName);
        }

        // Заполняем html таблицу
        for (int i = 1; i <= recordCount; i++)
        {
            int j = 1;
            for (std::vector<TFieldLink>::iterator it = links.begin(); it != links.end(); it++)
            {
                 if ( it->dsField == NULL )
                 {
                        continue;
                 }

                 //String FieldName = mergetable.head.GetElement(1, j);

                 mergetable.PutRecord(j++, it->dsField->AsString);
            }
            dataSet->Next();
            mergetable.Next();
        }

        // Закоментированно 2017-02-21
        //MSWordWorks msword;
        //return msword.MergeDocumentToFiles(Document, mergetable);

        return MergeDocumentToFiles(Document, mergetable);
    }
    catch (Exception &e)
    {
        throw Exception(e); // Добавлено 2016-03-25. Проверить!
    }

    //return std::vector<String> ();
}

/* Заменяет поля FormFields, используя значения из текущей строки dataSet
   В качестве имени поля используется значение "Текст по умолчанию".
*/
void MSWordWorks::ReplaceFormFields(Variant Document, TDataSet* dataSet)
{
    if (!dataSet->Active || dataSet->Eof)
    {
        return;
    }

    Variant FormFields = Document.OlePropertyGet("FormFields");
    Variant Field = Unassigned;
    int n = FormFields.OlePropertyGet("Count");

    for (int i = n; i > 0; i--)
    {
        Variant formFieldsItem = FormFields.OleFunction("Item", i);
        String fieldNameCode = UpperCase(formFieldsItem.OlePropertyGet("Result"));
        String fieldName;

        bool isImg = false;
        int imgParam_zOrder;

        if (fieldNameCode.Pos("[IMG]") == 1)
        {
            isImg = true;
            imgParam_zOrder = 4;
        }
        else if (fieldNameCode.Pos("[SHAPE]") == 1)
        {
            isImg = true;
            imgParam_zOrder = 5;
        }

        if ( isImg )
        {
            int posClosingBracket = fieldNameCode.Pos("]");
            fieldName = fieldNameCode.SubString(posClosingBracket + 1, fieldNameCode.Length() - posClosingBracket);
        }
        else
        {
            fieldName = fieldNameCode;
        }

        TField* Field = dataSet->Fields->FindField(fieldName);
        if (Field != NULL) // Если нашли поле
        {
            if (isImg)
            {
                String imgPath = Field->AsString;
                if ( FileExists(imgPath) )
                {
                    Variant inlineShape = SetPictureToFormField(Document, i, imgPath);

                    if (imgParam_zOrder == 5)
                    {
                        ConverInlineShapeToShape(inlineShape, imgParam_zOrder);
                    }
                }
                else
                {
                    SetTextToFieldF(Document, i, "Файл изображения не найден! (" + imgPath + ")");
                }
            }
            else
            {
                SetTextToFieldF(Document, i, Field->AsString.c_str());
            }
        }
    }
}

void MSWordWorks::ReplaceVariablesAll(Variant Document, TDataSet* dataSet, TDocFieldType fieldType, const String& fieldNamePrefix)
{
    ReplaceVariablesDocumentBody(Document, dataSet, fieldType, fieldNamePrefix);   // Тело документа
    ReplaceVariablesHeadersAndFooters(Document, dataSet, fieldType, fieldNamePrefix);   // Колонтитулы
}

/* Заменяет поля DOCVARIABLE
   Сопоставляет поля из DataSet и DOCVARIABLE
   Создает Variables в соответствии с именами полей DOCVARIABLE
   Задает значения этих Variables
   и Обновляет поля DOCVARIABLE
*/
void MSWordWorks::ReplaceVariablesDocumentBody(Variant Document, TDataSet* dataSet, TDocFieldType fieldType, const String& fieldNamePrefix)
{
    Variant fields = Document.OleFunction("Range").OlePropertyGet("Fields");
    ReplaceVariables_(Document, dataSet, fields, fieldType, fieldNamePrefix);
}

/* Обновляет значение полей в колонтитулах во всех разделах */
void MSWordWorks::ReplaceVariablesHeadersAndFooters(Variant Document, TDataSet* dataSet, TDocFieldType fieldType, const String& fieldNamePrefix)
{
    // Обновляем поля в колонтитулах во всех разделах
    Variant sections = Document.OlePropertyGet("Sections"); // Разделы
    int sectionsCount = sections.OlePropertyGet("Count");
    for (int i = 1; i <= sectionsCount; i++)
    {
        Variant section = sections.OleFunction("Item", i);

        // Верхний колонтитул
        Variant headers = section.OlePropertyGet("Headers");
        int headersCount = headers.OlePropertyGet("Count");
        for (int j = 1; j <= headersCount; j++)
        {
            Variant header = headers.OleFunction("Item", j);

            //
            ReplaceVariables_(Document, dataSet, header.OlePropertyGet("Range").OlePropertyGet("Fields"), fieldType, fieldNamePrefix);
        }

        // Нижний колонтитул
        Variant footers = section.OlePropertyGet("Footers");
        int footersCount = footers.OlePropertyGet("Count");
        for (int j = 1; j <= footersCount; j++)
        {
            Variant footer = footers.OleFunction("Item", j);

            //
            ReplaceVariables_(Document, dataSet, footer.OlePropertyGet("Range").OlePropertyGet("Fields"), fieldType, fieldNamePrefix);
        }

    }

    FixAfterProcessingHeadersAndFooters(Document);  // Исправление глюка
}


/* Test */
void MSWordWorks::ReplaceVariables_(Variant Document, TDataSet* dataSet, Variant Fields, TDocFieldType fieldType, const String& fieldNamePrefix)
{
    std::vector<TFieldLink> links = assignDataSetToRangeFields(Fields, fieldType, dataSet, fieldNamePrefix);

    if ( links.size() == 0 || !dataSet->Active || dataSet->Eof)     // Если кол-во сопоставленных полей = 0, то выходим
    {
        return;
    }

    Variant variables = Document.OlePropertyGet("Variables");
    for (std::vector<TFieldLink>::iterator it = links.begin(); it != links.end(); it++)
    {
        if ( it->dsField != NULL )
        {
            if (it->dsField->AsString != "")    // Чтобы не появлялось сообщение "Ошибка! Переменная документа не указана."
            {
                Variant variable = variables.OleFunction("Item", (OleVariant)it->docFieldName);
                variable.OlePropertySet("Value", (OleVariant)it->dsField->AsString);
                it->docField.OleFunction("Update");     // returns True if success
            }
            it->docField.OleFunction("Unlink");     // returns True if success   // Преобразует поле в значение
        }
    }
}


/* Заменяет поля DOCVARIABLE на изображения */
void MSWordWorks::ReplaceImageVariables(Variant Document, TDataSet* dataSet, const String& fieldNamePrefix)
{
    Variant fields = Document.OleFunction("Range").OlePropertyGet("Fields");
    std::vector<TFieldLink> links = assignDataSetToRangeFields(fields, DFT_DOCVARIABLE, dataSet, fieldNamePrefix);

    if ( links.size() == 0 || !dataSet->Active || dataSet->Eof)
    {
        return;
    }

    for (std::vector<TFieldLink>::iterator it = links.begin(); it != links.end(); it++)
    {
        if ( it->dsField == NULL )
        {
            continue;
        }

        String imagePath = it->dsField->AsString.c_str();
        //String imagePath = dataSet->Fields->FieldByNumber(it->datasetFieldIndex)->AsString.c_str();
        if ( FileExists(imagePath) )
        {
            Variant rangeResult = it->docField.OlePropertyGet("Result");
            Variant inlineShape = SetPictureToRange(Document, rangeResult, imagePath);
            //if (imgParam_zOrder == 5)
            ConverInlineShapeToShape(inlineShape, 5);
            it->docField.OleProcedure("Delete");
        }
    }
}

/*
    // Получаем образец строки
    Variant rows = table.OlePropertyGet("Rows");
    Variant rowTemplate = rows.OleFunction("Item", 2);
    Variant formatedText = rowTemplate.OlePropertyGet("Range").OlePropertyGet("FormattedText");

    // Добавляем строку таблицы для позиционирования при вставке
    // (возможно здесь можно доработать, но сойдет и так. Позже эту строку удалим)
    Variant rowTmp = rows.OleFunction("Add");

    int currentRow = 3;

    while ( !dataSet->Eof ) // Цикл по dataSet
    {
        // Добавляем новую строку по образцу и получаем список полей
        rowTmp.OlePropertyGet("Range").OlePropertySet("FormattedText" , formatedText );
        Variant fields = rows.OleFunction("Item", currentRow++).OlePropertyGet("Range").OlePropertyGet("Fields");

        for (std::vector<TLinkFields>::iterator it = links.begin(); it < links.end(); it++)
        {
            Variant field = fields.OleFunction("Item", it->second);
            field.OlePropertyGet("Result").OlePropertySet("Text", dataSet->Fields->FieldByNumber(it->first)->AsString.c_str());
            //field.OleProcedure("Delete");   // если нужно удалять поля, то цикл сделать в обратном порядке
        }
        dataSet->Next();

    }

    // Удаляем вспомогательную строку и строку-образец
    rowTmp.OleProcedure("Delete");
    rowTemplate.OleProcedure("Delete");

*/


/* Вычисляет имя поля указанного типа */
String MSWordWorks::getFieldName(Variant field, int fieldType)
{
    int type = field.OlePropertyGet("Type");

    if (fieldType != 0 && type != fieldType)
    {
        return String("");
    }

    // Определяем кодовое имя поля по числовому коду типа поля
    String fieldDescr = "";
    switch(type)
    {
    case 64:   // wdFieldDocVariable
    {
        fieldDescr = "DOCVARIABLE";
        break;
    }
    case 59:    // wdFieldMergeField
    {
        fieldDescr = "MERGEFIELD";
        break;
    }
    case 34:
    {
        fieldDescr = "FUNCTION";
        break;
    }
    default:
    {
        return String("");
    }
    }

    Variant code = field.OlePropertyGet("Code");   // code is a range of field text
    String text = code.OlePropertyGet("Text");
    int textLength = text.Length();

    int p1 = 0;
    int p2 = 0;

    // Определяем позицию конца типа поля
    if ( fieldDescr != "")
    {
        p1 = text.Pos(fieldDescr) + fieldDescr.Length() + 1;

        for (int i = p1; i < text.Length(); i++)
        {
            if ( text[i] == ' ' )
            {
                p1++;
            }
            else
            {
                break;
            }
        }
    }
    else
    {
        return String("");
    }
    // Определяем позицию начала имени поля
    for (int i = p1; i < textLength; i++)
    {
        if ( text[i] != ' ' && text[i] != '"' )
        {
            p1 = i;
            break;
        }
    }

    // Определяем позицию конца имени поля
    for (int i = p1; i < textLength; i++)
    {
        if ( text[i] == ' ' || text[i] == '"')
        {
            p2 = i;
            break;
        }
    }
    if (p2 == 0)
    {
        p2 = textLength;
    }
    return text.SubString(p1, p2 - p1);
}

/* Удаляет префикс из строки */
String MSWordWorks::deletePrefix(String value, String prefix)
{
    int prefixLength = prefix.Length();
    if (prefixLength == 0)
    {
        return value;
    }
    int valueLength = value.Length();
    if ( value.Pos(prefix) == 1 && prefixLength < valueLength)
    {
        return value.SubString(prefixLength + 1, valueLength - prefixLength);
    }
    else
    {
        return "";
    }
}

/* Возвращает список для привязки полей из dataSet к table
*/
std::vector<TFieldLink> MSWordWorks::assignDataSetToRangeFields(Variant fields, TDocFieldType fieldType, TDataSet* dataSet, const String& fieldNamePrefix)
{
    std::vector<TFieldLink> result;
    result.reserve(dataSet->FieldCount);

    int n = fields.OlePropertyGet("Count");
    int prefixLength = fieldNamePrefix.Length();

    for (int i = 1; i <= n; i++ )         // цикл по полям в строке
    {
        Variant field = fields.OleFunction("Item", i);
        String docFieldName = getFieldName(field, fieldType); // 2017-03-31
        if (docFieldName == "")     // Если поле указанного типа не найдено (не определено имя)
        {
            continue;
        }
        String dsFieldName = "";

        // Определяем имя поля для поиска в dataSet
        // с учетом возможного использования префикса и без
        if ( prefixLength > 0 )
        {
            if ( docFieldName.Length() > prefixLength && docFieldName.SubString(1, prefixLength) == fieldNamePrefix)
            {
                dsFieldName = docFieldName.SubString(prefixLength+1, docFieldName.Length() - prefixLength);     // 2017-03-31
            }
            else
            {
                continue;
            }
        }
        else
        {
            dsFieldName = docFieldName;
        }

        if (dsFieldName != "")
        {
            // Сопоставляем поле с полем в источнике
            //int fieldIndex = field.OlePropertyGet("Index");   //  Это значение является сквозным значением на весь документ

            TField* fieldOfDataSet = dataSet->Fields->FindField(dsFieldName);
            if ( fieldOfDataSet )       // Если удалось уйти в dataset
            {                          
                result.push_back(TFieldLink(i, field, docFieldName, fieldOfDataSet->FieldNo, fieldOfDataSet, dsFieldName));
                //result.push_back(std::make_pair(fieldOfDataSet->FieldNo, i));
            }
            else
            {
                result.push_back(TFieldLink(i, field, docFieldName, -1, NULL, dsFieldName));
            }
        }
    }

    return result;
}

/* Выводить данные из dataSet в таблицу table
   в таблице должны быть указаны имена полей из dataSet
   с использованием полей DOCVARIABLE */
void MSWordWorks::writeDataSetToTable(Variant table, TDataSet* dataSet, const String& fieldNamePrefix)
{
    Variant tableRange = table.OlePropertyGet("Range");
    Variant fields = tableRange.OlePropertyGet("Fields");
    std::vector<TFieldLink> links = assignDataSetToRangeFields(fields, 64, dataSet, fieldNamePrefix);

    if ( links.size() == 0 || !dataSet->Active || dataSet->Eof)
    {
        return;
    }

    // Копируем строку таблицы с переменными в буфер обмена (плохой вариант)
    //Variant rows = table.OlePropertyGet("Rows");
    //Variant row = rows.OleFunction("Item", 2);
    //Variant selection = row.OlePropertyGet("Range").OleFunction("Select");//.OlePropertyGet("Selection");//.OleFunction("Copy");
    //selection = WordApp.OlePropertyGet("Selection");
    //selection.OleProcedure("Copy");

    // Получаем образец строки
    // и сохраняем форматированный текст, который будем вставлять в добавляемые строки
    Variant rows = table.OlePropertyGet("Rows");
    Variant rowTemplate = rows.OleFunction("Item", 2);
    Variant formatedText = rowTemplate.OlePropertyGet("Range").OlePropertyGet("FormattedText");

    int rowCount = rows.OlePropertyGet("Count");
    int footerRowCount = rowCount - 2;

    // Получаем доступ к переменны документа
    Variant variables = tableRange.OlePropertyGet("Document").OlePropertyGet("Variables");


    // Добавляем пустую строку в конец таблицы для позиционирования при вставке
    // (возможно здесь можно доработать, но сойдет и так. Позже эту строку удалим)
    // или назначаем строкой для позиционирования строку итогов (подвал таблицы)
    Variant rowTmp;     // Строка для позиционирования
    int currentRow = 3;

    if (footerRowCount > 0)
    {
        rowTmp = rows.OleFunction("Item", 3);
    }
    else
    {
        rowTmp = rows.OleFunction("Add");
    }


    //currentRow--;

    while ( !dataSet->Eof ) // Цикл по dataSet
    {
        // Добавляем новую строку по образцу
        // в позицию rowTmp. Фактически перед строкой rowTmp (перед последней строкой таблицы).
        // И получаем список полей в этой строке
        //rowTmp.OlePropertyGet("Range").OlePropertyGet("FormattedText").OlePropertySet("FormattedText" , formatedText );

        rowTmp.OlePropertyGet("Range").OlePropertySet("FormattedText" , formatedText );

        Variant fields = rows.OleFunction("Item", currentRow).OlePropertyGet("Range").OlePropertyGet("Fields");

        fields.OleProcedure("ToggleShowCodes");     // Добавлено для того, чтобы после Update применился формат, указанный в поле DOCVARIABLE
                                                    // иначе по неизвестной причине формат не применяется

        for (std::vector<TFieldLink>::iterator it = links.begin(); it != links.end(); it++)
        {
            if ( it->dsField != NULL ) // Если есть присоединенное поле в  dataset, то подставляем значение
            {
                Variant field = fields.OleFunction("Item", it->docFieldIndex );
                Variant variable = variables.OleFunction("Item", (OleVariant)it->docFieldName);
                if ( it->dsField->AsString != "" )  // Необходимо, т.к при присвоении "" переменной возникает ошибка "Переменная документа не указана"
                {
                    variable.OlePropertySet("Value", it->dsField->AsString.c_str());
                    field.OleFunction("Update");
                }
            }

        }
        //fields.OleFunction("Unlink");     // Не! делаем преобразование полей в значения в целях оптимизации
                                            // так как для этого необходим дополнительный цикл с проверкой "была ли замена"
        dataSet->Next();    // Следующая строка таблицы
        currentRow++;
    }

    // Если у таблицы есть подвал
    if (footerRowCount > 0)
    {
        // Обновляем значения полей в строке итогов
        Variant fields = rowTmp.OlePropertyGet("Range").OlePropertyGet("Fields");
        UpdateFields(fields, 34);
    }
    else
    {
        // Удаляем вспомогательную строку для позиционирования
        rowTmp.OleProcedure("Delete");
    }

    // Удаляем строку-шаблон
    rowTemplate.OleProcedure("Delete");

}


/* Обновляет значения в полях определенного типа
*/
void MSWordWorks::UpdateFields(Variant fields, int fieldType, const String& fieldNamePrefix)
{
    // 2017-08-23 проверить функцию целиком!!!!

    int n = fields.OlePropertyGet("Count");

    for (int i = 1; i <= n; i++ )         // цикл по полям в строке
    {
        Variant field = fields.OleFunction("Item", i);

        String fieldName = getFieldName(field);
        //.OlePropertyGet("Name");

        if ((fieldType = -1 || field.OlePropertyGet("Type") == fieldType) && (fieldNamePrefix=="" || fieldName == deletePrefix(fieldName,fieldNamePrefix) ) )
        {
            field.OleFunction("Update");
        }
    }
}




/* Обновляет значение полей в колонтитулах во всех разделах */
void MSWordWorks::UpdateFieldsHeadersAndFooters(Variant Document, int fieldType, const String& fieldNamePrefix)
{
    // Обновляем поля в колонтитулах во всех разделах
    Variant sections = Document.OlePropertyGet("Sections"); // Разделы
    int sectionsCount = sections.OlePropertyGet("Count");
    for (int i = 1; i <= sectionsCount; i++)
    {
        Variant section = sections.OleFunction("Item", i);

        // Верхний колонтитул
        Variant headers = section.OlePropertyGet("Headers");
        int headersCount = headers.OlePropertyGet("Count");
        for (int j = 1; j <= headersCount; j++)
        {
            Variant header = headers.OleFunction("Item", j);
            UpdateFields(header.OlePropertyGet("Range").OlePropertyGet("Fields"), fieldType, fieldNamePrefix);
        }

        // Нижний колонтитул
        Variant footers = section.OlePropertyGet("Footers");
        int footersCount = footers.OlePropertyGet("Count");
        for (int j = 1; j <= footersCount; j++)
        {
            Variant footer = footers.OleFunction("Item", j);
            UpdateFields(footer.OlePropertyGet("Range").OlePropertyGet("Fields"), fieldType, fieldNamePrefix);
        }

    }

    FixAfterProcessingHeadersAndFooters(Document);  // Исправление глюка
}



/* Обновляет все поля в документе */
void MSWordWorks::UpdateAllFields(Variant Document, int fieldType, const String& fieldNamePrefix)
{
    // Обновляем поля в теле документа
    UpdateFields(Document.OlePropertyGet("Fields"), fieldType, fieldNamePrefix);

    UpdateFieldsHeadersAndFooters(Document, fieldType, fieldNamePrefix);


    // Обновление разделов содержания  2017-06-23 Uncompleted
    //ActiveDocument.TablesOfContents(1).Update
}


/* Обновляет все поля в документе */
void MSWordWorks::UpdateAllFieldsFast(Variant Document)
{
    Document.OleProcedure("PrintPreview");
    Document.OleProcedure("ClosePrintPreview");
}


/* Исправление после работы с верхним и нижним колонтитулом
   Исправляет глюк с отступом тела документа от верхнего края границы листа,
   возникающий про обращении к полям в колонтитулах
*/
void MSWordWorks::FixAfterProcessingHeadersAndFooters(Variant Document)
{
    Variant window = Document.OlePropertyGet("ActiveWindow");
    Variant view = window.OlePropertyGet("View");
    view.OlePropertySet("SeekView", 9);     // wdSeekCurrentPageHeader = 9
    view.OlePropertySet("SeekView", 0);     // wdSeekMainDocument = 0
}



/* Преобразует поля в простые значения
*/
void MSWordWorks::UnlinkFields(Variant fields)
{
    fields.OleFunction("Unlink");
}

/* Преобразует все поля в документе в простые значения
*/
void MSWordWorks::UnlinkAllFields(Variant Document)
{
    // Сначала преобразуем поля в теле документа
    UnlinkFields(Document.OlePropertyGet("Fields"));

    // Затем в колонтитулах
    Variant sections = Document.OlePropertyGet("Sections");
    int sectionsCount = sections.OlePropertyGet("Count");

    for (int i = 1; i <= sectionsCount; i++)
    {
        Variant section = sections.OleFunction("Item", i);

        // Верхний колонтитул
        Variant headers = section.OlePropertyGet("Headers");
        int headersCount = headers.OlePropertyGet("Count");
        for (int j = 1; j <= headersCount; j++)
        {
            Variant header = headers.OleFunction("Item", j);
            UnlinkFields(header.OlePropertyGet("Range").OlePropertyGet("Fields"));
        }

        // Нижний колонтитул
        Variant footers = section.OlePropertyGet("Footers");
        int footersCount = footers.OlePropertyGet("Count");
        for (int j = 1; j <= footersCount; j++)
        {
            Variant footer = footers.OleFunction("Item", j);
            UnlinkFields(footer.OlePropertyGet("Range").OlePropertyGet("Fields"));
        }
    }

    FixAfterProcessingHeadersAndFooters(Document);

}



/*
    //Variant t = variables.OleFunction("Add", "имя переменной");
*/




/*

SaveAs2 аналог SaveAs. Между ними сущетсвует небольшое отличие (см. параметр CompatibilityMode)

ActiveDocument.SaveAs2 FileName
        FileFormat:=wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, CompatibilityMode:=14



wdFormatDocument                    =  0
wdFormatDocument97                  =  0
wdFormatDocumentDefault             = 16
wdFormatDOSText                     =  4
wdFormatDOSTextLineBreaks           =  5
wdFormatEncodedText                 =  7
wdFormatFilteredHTML                = 10
wdFormatFlatXML                     = 19
wdFormatFlatXMLMacroEnabled         = 20
wdFormatFlatXMLTemplate             = 21
wdFormatFlatXMLTemplateMacroEnabled = 22
wdFormatHTML                        =  8
wdFormatPDF                         = 17
wdFormatRTF                         =  6
wdFormatTemplate                    =  1
wdFormatTemplate97                  =  1
wdFormatText                        =  2
wdFormatTextLineBreaks              =  3
wdFormatUnicodeText                 =  7
wdFormatWebArchive                  =  9
wdFormatXML                         = 11
wdFormatXMLDocument                 = 12
wdFormatXMLDocumentMacroEnabled     = 13
wdFormatXMLTemplate                 = 14
wdFormatXMLTemplateMacroEnabled     = 15
wdFormatXPS                         = 18
wdFormatOfficeDocumentTemplate      = 23
wdFormatMediaWiki                   = 24



//Documents.OleFunction("Open", FileName, ConfirmConversions, ReadOnly, AddToRecentFiles,
//    PasswordDocument, PasswordTemplate, Revert,
//    WritePasswordDocument, WritePasswordTemplate, Format   )

*/


/*
WdFieldType Enumeration (Word)

Office 2013 and later Other Versions 
GitHub-Mark-64px	
Contribute to this content
Use GitHub to suggest and submit changes. See our guidelines for contributing to VBA documentation.
Specifies a Microsoft Word field. Unless otherwise specified, the field types described in this enumeration can be added interactively to a Word document by using the Field dialog box. See the Word Help for more information about specific field codes.
Name
Value
Description
wdFieldAddin
81
Add-in field. Not available through the Field dialog box. Used to store data that is hidden from the user interface.
wdFieldAddressBlock
93
AddressBlock field.
wdFieldAdvance
84
Advance field.
wdFieldAsk
38
Ask field.
wdFieldAuthor
17
Author field.
wdFieldAutoNum
54
AutoNum field.
wdFieldAutoNumLegal
53
AutoNumLgl field.
wdFieldAutoNumOutline
52
AutoNumOut field.
wdFieldAutoText
79
AutoText field.
wdFieldAutoTextList
89
AutoTextList field.
wdFieldBarCode
63
BarCode field.
wdFieldBidiOutline
92
BidiOutline field.
wdFieldComments
19
Comments field.
wdFieldCompare
80
Compare field.
wdFieldCreateDate
21
CreateDate field.
wdFieldData
40
Data field.
wdFieldDatabase
78
Database field.
wdFieldDate
31
Date field.
wdFieldDDE
45
DDE field. No longer available through the Field dialog box, but supported for documents created in earlier versions of Word.
wdFieldDDEAuto
46
DDEAuto field. No longer available through the Field dialog box, but supported for documents created in earlier versions of Word.
wdFieldDisplayBarcode
99
DisplayBarcode field.
wdFieldDocProperty
85
DocProperty field.
wdFieldDocVariable
64
DocVariable field.
wdFieldEditTime
25
EditTime field.
wdFieldEmbed
58
Embedded field.
wdFieldEmpty
-1
Empty field. Acts as a placeholder for field content that has not yet been added. A field added by pressing Ctrl+F9 in the user interface is an Empty field.
wdFieldExpression
34
= (Formula) field.
wdFieldFileName
29
FileName field.
wdFieldFileSize
69
FileSize field.
wdFieldFillIn
39
Fill-In field.
wdFieldFootnoteRef
5
FootnoteRef field. Not available through the Field dialog box. Inserted programmatically or interactively.
wdFieldFormCheckBox
71
FormCheckBox field.
wdFieldFormDropDown
83
FormDropDown field.
wdFieldFormTextInput
70
FormText field.
wdFieldFormula
49
EQ (Equation) field.
wdFieldGlossary
47
Glossary field. No longer supported in Word.
wdFieldGoToButton
50
GoToButton field.
wdFieldGreetingLine
94
GreetingLine field.
wdFieldHTMLActiveX
91
HTMLActiveX field. Not currently supported.
wdFieldHyperlink
88
Hyperlink field.
wdFieldIf
7
If field.
wdFieldImport
55
Import field. Cannot be added through the Field dialog box, but can be added interactively or through code.
wdFieldInclude
36
Include field. Cannot be added through the Field dialog box, but can be added interactively or through code.
wdFieldIncludePicture
67
IncludePicture field.
wdFieldIncludeText
68
IncludeText field.
wdFieldIndex
8
Index field.
wdFieldIndexEntry
4
XE (Index Entry) field.
wdFieldInfo
14
Info field.
wdFieldKeyWord
18
Keywords field.
wdFieldLastSavedBy
20
LastSavedBy field.
wdFieldLink
56
Link field.
wdFieldListNum
90
ListNum field.
wdFieldMacroButton
51
MacroButton field.
wdFieldMergeBarcode
98
MergeBarcode field.
wdFieldMergeField
59
MergeField field.
wdFieldMergeRec
44
MergeRec field.
wdFieldMergeSeq
75
MergeSeq field.
wdFieldNext
41
Next field.
wdFieldNextIf
42
NextIf field.
wdFieldNoteRef
72
NoteRef field.
wdFieldNumChars
28
NumChars field.
wdFieldNumPages
26
NumPages field.
wdFieldNumWords
27
NumWords field.
wdFieldOCX
87
OCX field. Cannot be added through the Field dialog box, but can be added through code by using the AddOLEControl method of the Shapes collection or of the InlineShapes collection.
wdFieldPage
33
Page field.
wdFieldPageRef
37
PageRef field.
wdFieldPrint
48
Print field.
wdFieldPrintDate
23
PrintDate field.
wdFieldPrivate
77
Private field.
wdFieldQuote
35
Quote field.
wdFieldRef
3
Ref field.
wdFieldRefDoc
11
RD (Reference Document) field.
wdFieldRevisionNum
24
RevNum field.
wdFieldSaveDate
22
SaveDate field.
wdFieldSection
65
Section field.
wdFieldSectionPages
66
SectionPages field.
wdFieldSequence
12
Seq (Sequence) field.
wdFieldSet
6
Set field.
wdFieldShape
95
Shape field. Automatically created for any drawn picture.
wdFieldSkipIf
43
SkipIf field.
wdFieldStyleRef
10
StyleRef field.
wdFieldSubject
16
Subject field.
wdFieldSubscriber
82
Macintosh only. For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.
wdFieldSymbol
57
Symbol field.
wdFieldTemplate
30
Template field.
wdFieldTime
32
Time field.
wdFieldTitle
15
Title field.
wdFieldTOA
73
TOA (Table of Authorities) field.
wdFieldTOAEntry
74
TOA (Table of Authorities Entry) field.
wdFieldTOC
13
TOC (Table of Contents) field.
wdFieldTOCEntry
9
TOC (Table of Contents Entry) field.
wdFieldUserAddress
62
UserAddress field.
wdFieldUserInitials
61
UserInitials field.
wdFieldUserName
60
UserName field.
wdFieldBibliography
97
Bibliography field.
wdFieldCitation
96
Citation field.*/

//rowTmp.OlePropertyGet("Range").OleProcedure("Collapse", 0); /*wdCollapseEnd   */

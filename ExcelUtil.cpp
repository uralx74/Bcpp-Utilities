#include "ExcelUtil.h"

using namespace strtools;
//----------------------------------------------------------------------------
// ��������� ����� � ����
void __fastcall MSExcelWorks::SaveDocument(Variant& workbook, const AnsiString& FileName)
{
    if (IsReadOnly(workbook))
    {
        throw Exception("Can't to save the document. The document is readonly.");
    }

    if (FileName != "") {
        try
        {    // ���������� ��������� � ����� ����
		    workbook.OleProcedure("SaveAs", FileName.c_str());
        } 
        catch (Exception &e)
        {
            throw Exception("�� ������� ��������� �������� � ������ \"" + FileName + "\"");
        }
    }
    else 
    {
        try 
        {  // ���������� ����� ��������� ���������
		    workbook.OleProcedure("Save");
        } 
        catch (Exception &e)
        {
            throw Exception("�� ������� ��������� ��������. " + e.Message);
        }

    }
}

//----------------------------------------------------------------------------
// ���������� true ���� �������� � ������ ReadOnly
bool __fastcall MSExcelWorks::IsReadOnly(Variant& workbook)
{
    return workbook.OlePropertyGet("ReadOnly");
}


//----------------------------------------------------------------------------
// ��������/��������� ����������
void __fastcall MSExcelWorks::SetAutoFilter(Variant& range)
{
    // expression .AutoFilter(Field, Criteria1, Operator, Criteria2, VisibleDropDown)
    range.OleFunction("AutoFilter");
}

//----------------------------------------------------------------------------
// ��������� ������� ������� �� �����������
void __fastcall MSExcelWorks::SetColumnsAutofit(Variant& range)
{
    //range.OlePropertyGet("EntireColumn").OleProcedure("AutoFit");     // ��� ������������� ������� ��������� �������!
    range.OlePropertyGet("Columns").OleProcedure("AutoFit");
}

//----------------------------------------------------------------------------
//
void __fastcall MSExcelWorks::SetColumnWidth(Variant& worksheet, int ColumnIndex, int width)
{
    if (width >= 0)
    {
        worksheet.OlePropertyGet("Columns").OlePropertyGet("Item", ColumnIndex).OlePropertySet("ColumnWidth", width);
    }
    else
    {
        worksheet.OlePropertyGet("Columns").OlePropertyGet("Item", ColumnIndex).OleProcedure("AutoFit");
    }
}

//----------------------------------------------------------------------------
// ��������� ������ ������� �� �����������
void __fastcall MSExcelWorks::SetRowsAutofit(Variant& range)
{
    //range.OlePropertyGet("EntireRow").OleProcedure("AutoFit"); // ��� ������������� ������� ��������� �������!
    range.OlePropertyGet("Rows").OleProcedure("AutoFit");
}

//----------------------------------------------------------------------------
//
int __fastcall MSExcelWorks::GetRowHeight(Variant& range)
{
    return range.OlePropertyGet("RowHeight");
}

//----------------------------------------------------------------------------
//
void __fastcall MSExcelWorks::SetRowHeight(Variant& range, int Height)
{
    //range.OlePropertyGet("Rows").OlePropertySet("RowWidth", Width);
    range.OlePropertySet("RowHeight", Height);
}

//----------------------------------------------------------------------------
//
void __fastcall MSExcelWorks::SetColumnWidth(Variant& range, int Width)
{
    range.OlePropertySet("ColumnWidth", Width);
}

//----------------------------------------------------------------------------
// ��������� ���������� ����� � Range
int __fastcall MSExcelWorks::GetRangeRowsCount(Variant& range)
{
     return range.OlePropertyGet("Rows").OlePropertyGet("Count");
}

//----------------------------------------------------------------------------
// ��������� ���������� �������� � Range
int __fastcall MSExcelWorks::GetRangeColumnsCount(Variant& range)
{
    return range.OlePropertyGet("Columns").OlePropertyGet("Count");
}

//----------------------------------------------------------------------------
// ���������� ������ Range
AnsiString __fastcall MSExcelWorks::GetRangeFormat(Variant& range)
{
    AnsiString result;
    result = range.OlePropertyGet("NumberFormat");
    return result;
}

//----------------------------------------------------------------------------
// ������� ������ Range
void __fastcall MSExcelWorks::ClearFormats(Variant& range)
{
    range.OleProcedure("ClearFormats");
}

//----------------------------------------------------------------------------
// ������� ���� ����
void __fastcall MSExcelWorks::ClearWorksheet(Variant& Worksheet)
{
    Worksheet.OlePropertyGet("Cells").OleProcedure("Clear");
}

//----------------------------------------------------------------------------
// ���������� ������ � ������� ����� �� �����
// Object ����� ���� Workbook ��� Worksheet.
std::vector<TNamedRange> __fastcall MSExcelWorks::GetNamesFromObject(Variant& Object, const String& prefix)
{
    Variant names = Object.OlePropertyGet("Names");
    int namesCount = names.OlePropertyGet("Count");
    std::vector<TNamedRange> result;
    result.reserve(namesCount);

    if (prefix == "")
    {
        for(int i = 1; i <= namesCount; i++)
        {
            Variant namesItem = names.OleFunction("Item", i);
            String name = namesItem.OlePropertyGet("Name");
            Variant range = namesItem.OlePropertyGet("RefersToRange");
            result.push_back(std::make_pair(name, range));
        }
    }
    else
    {
        for(int i = 1; i <= namesCount; i++)
        {
            Variant namesItem = names.OleFunction("Item", i);
            String name = namesItem.OlePropertyGet("Name");
            name = deletePrefix(name, prefix);
            if (name != "")
            {
                Variant range = namesItem.OlePropertyGet("RefersToRange");
                result.push_back(std::make_pair(name, range));
            }
        }
    }
    return result;
}

/* ������� ������� �� ������ */
String MSExcelWorks::deletePrefix(String value, String prefix)
{

     // ���������� ��� ���� ��� ������ � dataSet
    // � ������ ���������� ������������� �������� � ���
    /*if ( prefixLength > 0 )
    {
        String dsFieldName = "";
        int cellNameLength = cellName.Length(); // �������� ����� ����� ������

        if (cellNameLength > prefixLength)
        {
            String cellNamePrefix = cellName.SubString(1, prefixLength);    // �������� ������� �� ��������� ����� ������
            if (cellNamePrefix == prefix)
            {
                dsFieldName = cellName.SubString(prefixLength+1, cellNameLength - prefixLength);
            }
            else
            {
                throw Exception("Prefix of cell name does't match with dataset prefix.");
            }

            return dsFieldName;
        }
        else
        {
            throw Exception("Length of cell name less whan lenght of prefix.");
        }
    }

    return cellName;*/

    // ��������� ��� �������!
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

//----------------------------------------------------------------------------
// ���������� ������ � ������� ����� � �����
// ������ Workbook ����� ���� Worksheet
// ����������! � ���������� �������!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
std::vector<AnsiString> __fastcall MSExcelWorks::GetNamesFromWorkbook(Variant& Workbook)
{
    Variant vNames = Workbook.OlePropertyGet("Names");
    int nNamesCount = vNames.OlePropertyGet("Count");
    std::vector<String> vFields;
    vFields.reserve(nNamesCount);

    for(int i=1; i < nNamesCount + 1; i++)
    {
        String sName = vNames.OleFunction("Item", i).OlePropertyGet("Name");

        // ��� ����� �������� ������
        int p1 = sName.Pos("!");
        if (p1 > 0)              // ����� ����� ������ ����� ! (������: ����1!���)
        {
            sName = sName.SubString(p1+1, sName.Length()-p1);
        }
        vFields.push_back(sName);
    }
    return vFields;   
}

//----------------------------------------------------------------------------
// ��������� Range
Variant __fastcall MSExcelWorks::GetRange(Variant& Worksheet, int firstRow, int firstCol, int countRow, int countCol)
{
    Variant sell_left_top = Worksheet.OlePropertyGet("Cells", firstRow, firstCol);
	Variant sell_right_bottom = Worksheet.OlePropertyGet("Cells", firstRow+countRow-1, firstCol+countCol-1);
	Variant range = Worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);
    return range;
}

Variant __fastcall MSExcelWorks::GetRangeByNameGlobal(Variant& Workbook, const String& RangeName)
{
    Variant vNames = Workbook.OlePropertyGet("Names");

    Variant vName = vNames.OleFunction("Item", RangeName.c_str());

    return vName.OlePropertyGet("RefersToRange");
}

//----------------------------------------------------------------------------
// ��������� Range �� �����
Variant __fastcall MSExcelWorks::GetRangeByName(Variant& Worksheet, const String& RangeName)
{
    try
    {
        //Variant range = Worksheet.OlePropertyGet("Cells", RangeName);       // ������� ����� ��������� �� �����
        Variant range = Worksheet.OlePropertyGet("Range", RangeName.c_str());       // ������� ����� ��������� �� �����
        return range;
    }
    catch (EOleSysError &e)
    {                       // ���� ���� � ������ �� �������
        return Variant::Empty();
        //throw Exception("�� ������� ���������� Range �� ����� " + RangeName);
    }

/*    Variant Workbook = Worksheet.OlePropertyGet("Parent");

    Variant vNames = Workbook.OlePropertyGet("Names");
    AnsiString strName;
    //AnsiString strRefersTo, strSheetName;
    //String srCellName("");
    int nNameCount = vNames.OlePropertyGet("Count");

    for(int i=1; i < nNameCount + 1; i++) {
        strName = vNames.OleFunction("Item", i).OlePropertyGet("Name");
        int n = strName.Pos("!");
        strName = strName.SubString(n+1, strName.Length() - n);

         if(strName == RangeName) {
            return vNames.OleFunction("Item", i).OlePropertyGet("RefersToRange");

            //int nRow = vNames.OleFunction("Item", i).OlePropertyGet("RefersToRange").OlePropertyGet("Row");
            //int nCol = vNames.OleFunction("Item", i).OlePropertyGet("RefersToRange").OlePropertyGet("Column");
            //strRefersTo = vNames.OleFunction("Item", i).OlePropertyGet("RefersTo");
            //strSheetName = strRefersTo.SubString(2, strRefersTo.Pos("!") - 2);
            //return WriteTable(worksheet, ArrayData,  nRow, nCol, DataFormat);
        }
    }
    //VarClear(
    return Unassigned; */
}

//----------------------------------------------------------------------------
// ��������� Range ������ ��������� Range
Variant __fastcall MSExcelWorks::GetRangeFromRange(Variant& range, int firstRow, int firstCol, int countRow, int countCol)
{
    Variant Cells = range.OlePropertyGet("Cells");
    Variant sell_left_top = Cells.OlePropertyGet("Item", firstRow, firstCol);
	Variant sell_right_bottom = Cells.OlePropertyGet("Item", firstRow+countRow-1, firstCol+countCol-1);
    Variant Worksheet = range.OlePropertyGet("Worksheet");
    //Variant Worksheet = range.OlePropertyGet("Parent");
    return Worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);
}

//------------------------------------------------------------------------------
// ��������� ������ ����� ���������� �� Range                   !!! �� ���������
Variant __fastcall MSExcelWorks::ReadRange(Variant& worksheet, int firstRow, int firstCol, int lastRow, int lastCol)
{

    Variant ArrayData;
    Variant sell_left_top = worksheet.OlePropertyGet("Cells", firstRow, firstCol);
	Variant sell_right_bottom = worksheet.OlePropertyGet("Cells", lastRow, lastCol);
	Variant range = worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);
    ArrayData = range.OlePropertyGet("Value");

    return ArrayData;
}

//------------------------------------------------------------------------------
// �������� ����� ������� � ������ ������                       !!! �� ��������
void __fastcall MSExcelWorks::CopyArray(const Variant &SrcArrayData, Variant* ArrayData,  int srcFirstRow, int srcFirstCol, int srcLastRow, int srcLastCol, int dstFirstRow, int dstFirstCol)
{
    for (int i = srcFirstRow; i <= srcLastRow; i ++)
    {
        for (int j = srcFirstCol; j <= srcLastCol; j++)
        {
            ArrayData->PutElement(SrcArrayData.GetElement(i,j), i, j);   // ����������!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        }
    }
}

//----------------------------------------------------------------------------
// ������ ������ ����� � Range
void __fastcall MSExcelWorks::SetRangeDataFormat(Variant& range, AnsiString& format)
{
    if (format != "")    // "m/d/yyyy" "@" "0.00" "General"
    {
        try
        {
            range.OlePropertySet("NumberFormat", format.c_str());
        }
        catch (...)
        {
        }
    }
}

//----------------------------------------------------------------------------
// ������ ������ ����� � Range �������� ������������
void __fastcall MSExcelWorks::SetRangeFormat(Variant& range,  const TCellFormat& cf, int firstRow, int firstCol, int countRow, int countCol)
{
	// ������������ ��������� range
    //Variant cell_left_top = range.OlePropertyGet("Cells", firstRow, firstCol);
	//Variant cell_right_bottom = range.OlePropertyGet("Cells", firstRow + countRow - 1, firstCol + countCol - 1);
	//Variant range_tmp = range.OlePropertyGet("Range", cell_left_top, cell_right_bottom);

    Variant range_tmp = GetRangeFromRange(range, firstRow, firstCol, countRow, countCol);
 	SetRangeFormat(range_tmp, cf);
}

//----------------------------------------------------------------------------
// ������ ������ ����� � Range
void __fastcall MSExcelWorks::SetRangeFormat(Variant& range, const TCellFormat& cf)
{
    if (cf.DataFormat != "")    // "m/d/yyyy" "@" "0.00" "General"
    {
        try {
            range.OlePropertySet("NumberFormat", cf.DataFormat.c_str());
        }
        catch (...)
        {
        }
    }

    // ��� 2 - �� ������ ����, 3 - �� ������, 4 - �� �������)
    if (cf.HorizontalAlignment)
    {
        range.OlePropertySet ("HorizontalAlignment", cf.HorizontalAlignment);
    }
    if (cf.VerticalAlignment)
    {
        range.OlePropertySet ("VerticalAlignment", cf.VerticalAlignment - 1);
    }
    if (cf.ShrinkToFit > -1)   // ���������� ������� ������ �� ������ ������
    {
        range.OlePropertySet("ShrinkToFit", cf.ShrinkToFit);
    }

    if (cf.bWrapText > -1) // ���������� �� ������
    {
        range.OlePropertySet("WrapText", cf.bWrapText);
    }

    Variant font = range.OlePropertyGet("Font");
    if (cf.FontStyle.Contains(TCellFormat::fsNormal))
    {
        font.OlePropertySet("Bold", false);
        font.OlePropertySet("Italic", false);
        font.OlePropertySet("Underline", false);
    }
    else
    {
        if (cf.FontStyle.Contains(TCellFormat::fsBold))
        {
            font.OlePropertySet("Bold", true);
        }
        if (cf.FontStyle.Contains(TCellFormat::fsItalic))
        {
            font.OlePropertySet("Italic", true);
        }
        if (cf.FontStyle.Contains(TCellFormat::fsUnderline))
        {
            font.OlePropertySet("Underline", true);
        }
    }

    if (cf.FontSize > 0)
    {
        font.OlePropertySet("Size", cf.FontSize);
    }

    if (cf.bSetFontColor)     // ���� ������ � ������
    {
        font.OlePropertySet("Color", cf.FontColor);
    }

    if (cf.bSetFillColor)
    {
        range.OlePropertyGet("Interior").OlePropertySet("Color", cf.FillColor);
    }

    if (cf.Width != -1)
    {
        SetColumnWidth(range, cf.Width);
    }

    if (cf.Height != -1)
    {
        SetColumnWidth(range, cf.Height);
    }


    //if (cf.bWrapText)

    // ��������!!!!!
    if (cf.BorderStyle >= 0)
    {
        Variant borders = range.OlePropertyGet("Borders");
        //borders.OlePropertySet("LineStyle", cf.BorderStyle);
        for (int i = TCellFormat::xlEdgeLeft; i <= TCellFormat::xlInsideVertical; i++)
        {
            if (cf.BordeLine.Contains(i))
            {
      	        //try {range.OlePropertyGet("Borders", i).OlePropertySet("LineStyle", cf.BorderStyle);} catch(...) {};
      	        range.OlePropertyGet("Borders", i).OlePropertySet("LineStyle", cf.BorderStyle);
            }
        }
    }

}


/* ������������� ������ ����� ��� ���������, �������� ��� �� �������� */
void __fastcall MSExcelWorks::SetRangeColumnsFormat(Variant& range, std::vector<TCellFormat> cf)
{
    int lastCol = GetRangeColumnsCount(range);
	Variant Columns = range.OlePropertyGet("Columns");
    int cfSize = cf.size();

    for (int i = 0; i < lastCol && i < cfSize; i++)          // ��� ������� ������� �� ������ ������
    {
        SetRangeFormat(range, cf[i], 1, i+1, 1, 1);
    }
}


//----------------------------------------------------------------------------
// ������ ������ �������� � Range
void __fastcall MSExcelWorks::SetRangeColumnsDataFormat(Variant& range, const TDataFormat &df)
{
    int lastCol = GetRangeColumnsCount(range);
	Variant Columns = range.OlePropertyGet("Columns");
    int SizeDF = df.size();

    for (int i = 0; i < lastCol && i < SizeDF; i++)          // ��� ������� ������� �� ������ ������
    {
        SetRangeDataFormat(Columns.OlePropertyGet("Item", i+1), df[i]);
    }
}

//----------------------------------------------------------------------------
// �������� ������ ����� �� ������ ��������� � ������
// Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
void __fastcall MSExcelWorks::CopyRangeFormat(Variant& range_src, Variant& range_dst)
{
    Variant Selection = ExcelApp.OlePropertyGet("Selection");   // ���������� ���������
    range_src.OleProcedure("Copy");
    range_dst.OleProcedure("PasteSpecial",-4122, -4142, false, false);
    ExcelApp.OlePropertySet("CutCopyMode", false);      // �������� ����� �����������
    Selection.OleProcedure("Select");                   // ��������������� ���������
}

//----------------------------------------------------------------------------
// ������� ���������� ������� � ��������� �� ����� �������
Variant __fastcall MSExcelWorks::WriteTableToRange(Variant& range, const Variant &ArrayData,  int firstRow, int firstCol, bool extendAllow, TDataFormat *DataFormat)
{
    Variant ArrayRowsCount = VarArrayHighBound(ArrayData, 1) - VarArrayLowBound(ArrayData, 1)+1;
    Variant ArrayColsCount = VarArrayHighBound(ArrayData, 2) - VarArrayLowBound(ArrayData, 2)+1;
    Variant rangeColumnsCount = GetRangeColumnsCount(range);

    //Variant range_tmp = GetRangeFromRange(range, 1, 1, ArrayRowsCount, rangeColumnsCount);
    Variant range_result = range;

    if (ArrayRowsCount > 1 && extendAllow)     // ���� ���-�� ����� ������ 1, �� ���������� ��������� ��������
    {
        // 2017-08-01
        Variant extend_range = GetRangeFromRange(range, 2, 1, ArrayRowsCount-1, rangeColumnsCount);
        extend_range = InsertRows(extend_range);          // ��������� ������ �� ������� ���� (���������� � ������ ����� - ������ ������)

        if (DataFormat == NULL)     // ���� �� ����� ������ ������, �� �������� �� ��������� ���������
        {
            CopyRangeFormat(range, extend_range);
        }

        range_result = GetRangeFromRange(range, 1, 1, ArrayRowsCount, rangeColumnsCount);   // ��������, ������������ � �������� ����������
    }

    Variant range_tmp = GetRangeFromRange(range, 1, 1, ArrayRowsCount, ArrayColsCount);         // ��������� �������� ��� ������ �������
    if (DataFormat != NULL)
    {
        SetRangeColumnsDataFormat(range_tmp, *DataFormat);
    }

    range_tmp.OlePropertySet("Value", ArrayData);		// ����� ������ � ��������. ��� ����� ���� ����� ��� ������� ���������� ������

    return range_result;
}

//----------------------------------------------------------------------------
// ������� ���������� ������� � ��������� �� ����� �������
Variant __fastcall MSExcelWorks::WriteTable(Variant& worksheet, const Variant &ArrayData,  int firstRow, int firstCol, TDataFormat *DataFormat)
{
    Variant ArrayRowsCount = VarArrayHighBound(ArrayData, 1) - VarArrayLowBound(ArrayData, 1)+1;
    Variant ArrayColsCount = VarArrayHighBound(ArrayData, 2) - VarArrayLowBound(ArrayData, 2)+1;

	int lastRow = firstRow + ArrayRowsCount - 1; // firstRow .. lastRow
	int lastCol = firstCol + ArrayColsCount - 1; // firstCol  .. lastCol

    Variant sell_left_top = worksheet.OlePropertyGet("Cells", firstRow, firstCol);
	Variant sell_right_bottom = worksheet.OlePropertyGet("Cells", lastRow, lastCol);
	Variant range = worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);

    return WriteTableToRange(range, ArrayData, firstRow, firstCol, false, DataFormat);

/*    if (DataFormat != NULL) {
        SetRangeColumnsFormat(range, *DataFormat);
    }

    range.OlePropertySet("Value", ArrayData);		// ����� ������ � ��������. ��� ����� ���� ����� ��� ������� ���������� ������
    return range; */
}

//----------------------------------------------------------------------------
// ������� ���������� ������� � ��������� (�����������) �� ����� �������
Variant __fastcall MSExcelWorks::WriteTable(Variant& worksheet, const Variant &ArrayData, AnsiString CellName, TDataFormat *DataFormat)
{
    Variant vNames = worksheet.OlePropertyGet("Names");

    //int nNameCount = vNames.OlePropertyGet("Count");

    AnsiString strName;

    Variant range = GetRangeByName(worksheet, CellName);

    int nRow = range.OlePropertyGet("Row");
    int nCol = range.OlePropertyGet("Column");

    return WriteTable(worksheet, ArrayData,  nRow, nCol, DataFormat);
}

//------------------------------------------------------------------------------
// ���������� � �������� ������ ������� ���� � �����
void __fastcall MSExcelWorks::DateTimeCreateDoc(Variant& wst, int Row, int Col)
{
	AnsiString txt = "���� �������� ���������: "+DateTimeToStr(Now());
	WriteToCell(wst, txt, Row, Col);
}

//------------------------------------------------------------------------------
// ������� � ������ ������ � �������� �������
Variant __fastcall MSExcelWorks::WriteToRange(Variant range, const String &txt, String format)
{
    //if (range.IsEmpty())
    if ( VarIsClear(range) )
    {
        return range;
    }
    if (format != "")
    {
        range.OlePropertySet("NumberFormat", format.c_str());   // ������������� ������ ������ ��� ������
    }
	//range.OlePropertySet("Value", txt.c_str());
	range.OlePropertySet("Value", (OleVariant)txt);
	return range;
}

/*//------------------------------------------------------------------------------
// ������� � ������ ������ � �������� �������   2
Variant __fastcall MSExcelWorks::WriteToRange(const AnsiString& txt, const AnsiString& sRangeName, AnsiString format)
{
    Variant range = GetRangeByName(sRangeName);
    return WriteToRange(txt, range, format);
}   */

//------------------------------------------------------------------------------
// ������� � ������ ������ � �������� �������
Variant __fastcall MSExcelWorks::WriteToCell(Variant& worksheet, const AnsiString &txt, int Row, int Col, AnsiString format)
{
	Variant range = worksheet.OlePropertyGet("Cells").OlePropertyGet("Item", Row, Col);
    return WriteToRange(range, txt, format);
}

//------------------------------------------------------------------------------
// ������� � ������ � �������� ������ ������ � �������� �������
Variant __fastcall MSExcelWorks::WriteToCell(Variant& worksheet, const AnsiString &txt, AnsiString CellName, AnsiString format)
{
    Variant range = GetRangeByName(worksheet, CellName);

    return WriteToRange(range, txt, format);
}

//------------------------------------------------------------------------------
// ������� ������� � ������
Variant __fastcall MSExcelWorks::WriteFormulaToCell(Variant& wst, const AnsiString &txt, int Row, int Col, bool fBold)
{
	Variant range = wst.OlePropertyGet("Cells").OlePropertyGet("Item", Row, Col);
	range.OlePropertySet("FormulaR1C1", txt.c_str());
    range.OlePropertyGet("Font").OlePropertySet("Bold", fBold);

	return range;
}

//------------------------------------------------------------------------------
// ������� ������� � ������
Variant __fastcall MSExcelWorks::WriteFormula(Variant& worksheet, const AnsiString &txt, int Row, int Col, int countRow, int countCol,  bool fBold)
{
    Variant sell_left_top = worksheet.OlePropertyGet("Cells", Row, Col);
	Variant sell_right_bottom = worksheet.OlePropertyGet("Cells", Row+countRow-1, Col+countCol-1);
	Variant range = worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);
	//Variant range = GetRange(wst, Row, Col, Row+countRow-1, Col+countCol-1);
    // wst.OlePropertyGet("Cells").OlePropertyGet("Item", Row, Col);
	range.OlePropertySet("FormulaR1C1", txt.c_str());
    range.OlePropertyGet("Font").OlePropertySet("Bold",fBold);

	return range;
}

/*//------------------------------------------------------------------------------
// ������� ������� � ������
Variant __fastcall MSExcelWorks::WriteFormula(Variant wst, int firstRow, int firstCol, int lastRow = 0, int lastCol = 0);
*/

//---------------------------------------------------------------------------
// ��������� ������ �� ������ ������� Excel
AnsiString __fastcall MSExcelWorks::ReadCell(Variant& worksheet, int Row, int Col)
{
    return worksheet.OlePropertyGet("Cells").OlePropertyGet("Item", Row, Col);
}

//---------------------------------------------------------------------------
// ��������� ������ �� ������ ������� Excel
Variant __fastcall MSExcelWorks::ReadCellFormula(Variant& worksheet, int Row, int Col)
{
    return worksheet.OlePropertyGet("Cells").OlePropertyGet("Item", Row, Col).OlePropertyGet("Formula");
}

//------------------------------------------------------------------------------
// ������� ������ ���� varVariant
Variant MSExcelWorks::CreateVariantArray(int RowCount, int ColCount)
{
    int Bounds[4] = {1, RowCount, 1, ColCount};
//    return VarArrayCreate(Bounds, varString);
    return VarArrayCreate(Bounds, 3, varVariant);
}

//------------------------------------------------------------------------------
// Increase the length of the variant array
void MSExcelWorks::RedimVariantArray(Variant &ArrayData, int RowCount)
{
    VarArrayRedim(ArrayData, RowCount);
}

//------------------------------------------------------------------------------
// ���������� ������ ������ � �������
std::vector<AnsiString> __fastcall MSExcelWorks::GetDataFormat(const Variant &ArrayData, int RowIndex)
{
    std::vector<AnsiString> formats;

    int firstCol = VarArrayLowBound(ArrayData, 2);
    int lastCol = VarArrayHighBound(ArrayData, 2);

    formats.reserve(lastCol - firstCol + 1);
    for (int i = firstCol; i <= lastCol; i++) {         // ��� ������� ������� �� ������ ������
        AnsiString s = ArrayData.GetElement(RowIndex, i);
        AnsiString format;

        // ������������� ������ ����� ��� ������
        if ( IsDate(s.c_str()) )
        {
            format = "��.��.����"; //"m/d/yyyy";    //"��.��.����";
        }
        else if ( IsFloat(s.c_str()) )
        {
            format = "0.00";
        }
        else if ( IsInt(s.c_str()) )   // �������, ��� ��� ������� ����� ����� �������� �� 0
        {
             format = "0";
        }
        else
        {
            format = "@";   // "General"
        }
        formats.push_back(format);
    }
    return formats;
}

//------------------------------------------------------------------------------
// ������������� �������� ����� Excel
void __fastcall MSExcelWorks::SetActiveWorkbook(Variant& Workbook)
{
    Workbook.OleProcedure("Activate");
}

//------------------------------------------------------------------------------
// ������������� �������� ���� � ����� Excel
void __fastcall MSExcelWorks::SetActiveWorksheet(Variant& worksheet)
{
    worksheet.OleProcedure("Activate");
}

//------------------------------------------------------------------------------
// ������������� ������ �� ������ ������ Excel
void __fastcall MSExcelWorks::SetActiveRange(Variant& Worksheet, int firstRow, int firstCol, int lastRow, int lastCol)
{
    GetRange(Worksheet, firstRow, firstCol, lastRow, lastCol).OleProcedure("Select");
}

//------------------------------------------------------------------------------
// ����������� ��������� Excel
// � ���������� ������ �� ������
void __fastcall MSExcelWorks::SetVisible(bool fVisible, bool fForeground)
{
    String ExcelCaption = ExcelApp.OlePropertyGet("Caption");
    HWND hExcelWindow = FindWindow("XLMAIN", ExcelCaption.c_str());

	if(!ExcelApp.IsEmpty()) {
    	ExcelApp.OlePropertySet("Visible", fVisible);
        //SetWindowPos(hExcelWindow, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE || SWP_NOSIZE);
    	// ExcelApp.OlePropertySet("Activate", true);
        ExcelApp.OlePropertySet("DisplayAlerts", true);         // ���������� ��������������

        if (fVisible && fForeground) {
            // ���� �� �������� ����
            ExcelApp.OlePropertySet("UserControl", true);
            //SetForegroundWindow(hExcelWindow);
            //SetWindowPos(hExcelWindow, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE || SWP_NOSIZE);
            //SetWindowPos(hExcelWindow, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE || SWP_NOSIZE);
            SendMessage(hExcelWindow,WM_SYSCOMMAND,SC_MAXIMIZE,0);
        }
    }
}

//------------------------------------------------------------------------------
// ����������� ��������� Excel
// � ���������� ������ �� ������
void __fastcall MSExcelWorks::SetVisible(Variant Workbook, bool fVisible, bool fForeground)
{
    // ����������!
    // �������� ����� ��������� Workbook �� �������� ����
    SetVisible(fVisible, fForeground);
}


//---------------------------------------------------------------------------
//������ Excel
Variant __fastcall MSExcelWorks::OpenApplication()
{
    if(ExcelApp.IsEmpty())
    {
        try {
    	    ExcelApp = CreateOleObject("Excel.Application");
            ExcelApp.OlePropertySet("DisplayAlerts", false);
            WorkBooks = ExcelApp.OlePropertyGet("Workbooks");
            ExcelApp.OlePropertySet("Visible", false);
        } catch (Exception &exception) {
            VariantClear(ExcelApp);
            throw Exception("�� ������� ������� �������� Excel");
        }
  	}
}

/* ������������ �������� �������� � ������� MSExcelWorks */
void __fastcall MSExcelWorks::AssignDocument(Variant Workbook)
{
    //Variant worksheet = range_src.OlePropertyGet("Parent");
    //Variant workbook = worksheet.OlePropertyGet("Parent");
    ExcelApp = Workbook.OlePropertyGet("Parent");
    WorkBooks = ExcelApp.OlePropertyGet("Workbooks");
}


//---------------------------------------------------------------------------
// ��������� �������� Workbook (���� TemplateName="", �� ������� ����� ��������)
Variant __fastcall MSExcelWorks::OpenDocument(AnsiString TemplateName)
{
    Variant Book;
    if (TemplateName != "")    // ������� ������������ ��������
	{
        try
        {
            // ����. 2017-01-24
            Book = WorkBooks.OleFunction("Open",  TemplateName.c_str());

            // ����. 2016-12-06
            //WorkBooks.OleProcedure("Open",  TemplateName.c_str());
            //Book = WorkBooks.OlePropertyGet("Item", 1);
  	    }
        catch (Exception &e)
        {
            throw Exception("������ ��� �������� �����: " + TemplateName + "." + e.Message);
        }
    }
    else                       // ������� ����� ��������
    {
        try
        {
            Book = WorkBooks.OleFunction("Add");
        }
        catch (Exception &e)
        {
            VariantClear(ExcelApp);
            throw Exception("������ �������� ��������� Excel.");
        }
    }
  	return Book;
}

/*
//------------------------------------------------------------------------------
//
Variant __fastcall MSExcelWorks::OpenWorksheetFromFile(AnsiString& FileName)
{
        Variant workbook;
        Variant worksheet;
        workbook = OpenDocument(FileName);
        worksheet = GetSheet(workbook, 1);
        return worksheet;
} */

//---------------------------------------------------------------------------
// ���������� �����
Variant __fastcall MSExcelWorks::AddSheet(Variant& Workbook, int SheetIndex)
{
    Variant Sheets = Workbook.OlePropertyGet("Worksheets");
    Variant Sheet;

    Variant position;
    if (SheetIndex <= 0)   // ���������� � ����� ����� (�� ���������)
    {
        position = Sheets.OlePropertyGet("Count");
        Variant After = Sheets.OlePropertyGet("Item", position);
        Sheet = Sheets.OleFunction("Add", Variant().NoParam(), After);
    }
    else if (SheetIndex == 1)    // ���������� � ������ �����
    {
        Sheet = Sheets.OleFunction("Add");
    }
    else
    {                        // ���������� � ������� SheetIndex
        position = SheetIndex - 1;
        Variant After = Sheets.OlePropertyGet("Item", position);
        Sheet = Sheets.OleFunction("Add", Variant().NoParam(), After);
    }

    return Sheet;
}

void __fastcall MSExcelWorks::SetSheetName(Variant Sheet, const String &SheetName)
{
    Sheet.OlePropertySet("Name", SheetName);
}

//---------------------------------------------------------------------------
// ���������� �������� �������� Excel
Variant __fastcall MSExcelWorks::GetActiveSheet()
{
    if (!VarIsNull(ExcelApp))   // VarIsEmpty
    {
        return ExcelApp.OlePropertyGet("ActiveSheet");
    }
    else
    {
        return NULL;
    }
}

/* ���������� ������ � ������� �����*/
int _fastcall MSExcelWorks::GetSheetsCount(Variant& Workbook)
{
    return Workbook.OlePropertyGet("Worksheets").OlePropertyGet("Count");
}


//---------------------------------------------------------------------------
// �������� �� ������                               // ����������, ����� ��������� ����� �����
Variant __fastcall MSExcelWorks::GetSheet(Variant& Workbook, int SheetIndex)
{
    if (!VarIsNull(ExcelApp))   // VarIsEmpty
    {
        try
        {
            Variant Worksheet = Workbook.OlePropertyGet("Worksheets", SheetIndex);
            return  Worksheet;
        }
        catch(...)
        {
            throw Exception("����� �� ������� �������.");
        }
        //return  WorkSheets.OlePropertyGet("Item", SheetIndex);
    }
    else
    {
        return NULL;
    }
}

//---------------------------------------------------------------------------
// �������� ���������� Excel � ��������� ���� �������� ������� ����
void __fastcall MSExcelWorks::CloseApplication()
{
	// �������� ���������� Excel (��� ������� �� ���������� ���������)
    if (!ExcelApp.IsEmpty())
    {
        Variant Workbook;
        while (WorkBooks.OlePropertyGet("Count") > 0)
        {
            Workbook = ExcelApp.OlePropertyGet("ActiveWorkbook");
            Workbook.OleFunction("Close", false);
        }
        //ExcelApp.OlePropertySet("DisplayAlerts", false);         // ���������� ��������������
        WorkBooks = Unassigned;   // ������������ VarClear(WorkBooks)
	    ExcelApp.OleProcedure("Quit");
        ExcelApp = Unassigned;
    }
}

//---------------------------------------------------------------------------
// ������� �������� ������� ����� Excel
void __fastcall MSExcelWorks::CloseWorkbook(Variant Workbook, bool fCloseAppIfNoDoc)
{
	Workbook.OleFunction("Close", false);
    Workbook = Unassigned;	// 2017-05-17 ���������
    if (fCloseAppIfNoDoc && WorkBooks.OlePropertyGet("Count") == 0)
    {
        CloseApplication();
    }
}

/* ��������� ��� ��� ������� range */
Variant __fastcall MSExcelWorks::AddName(Variant Object, const String& name, Variant range, const String& comment)
{
    Variant names = Object.OlePropertyGet("Names");

    Variant vName = names.OleFunction("Add", (OleVariant)name, range);

    if ( comment != "")
    {
        vName.OlePropertySet("Comment", (OleVariant)comment);
    }

    return vName;
}

/* �������� �������� ��� ����� */
void __fastcall MSExcelWorks::ChangeNamedRange(Variant old_range, Variant new_range)
{
    Variant vName = old_range.OlePropertyGet("Name");
    vName.OlePropertySet("RefersToR1C1", new_range);
}


//------------------------------------------------------------------------------
// ������������� ������� ������� ������� ������ ����� ��������� ���������
void __fastcall MSExcelWorks::DrawBorders(Variant& range, bool r7, bool r8, bool r9, bool r10, bool r11, bool r12)
{
	// r7  - ������� �����
  	// r8  - ������� ������
  	// r9  - ������� �����
  	// r10 - ������� ������
  	// r11 - ����� ������
  	// r12 - ����� �������

  	if (r7) range.OlePropertyGet("Borders",7).OlePropertySet("LineStyle", 1);
  	if (r8) range.OlePropertyGet("Borders",8).OlePropertySet("LineStyle", 1);
  	if (r9) range.OlePropertyGet("Borders",9).OlePropertySet("LineStyle", 1);
  	if (r10) range.OlePropertyGet("Borders",10).OlePropertySet("LineStyle", 1);
  	if (r11) try { range.OlePropertyGet("Borders",11).OlePropertySet("LineStyle", 1); } catch(...) {}
  	if (r12) try { range.OlePropertyGet("Borders",12).OlePropertySet("LineStyle", 1);} catch(...) {}
}

//------------------------------------------------------------------------------
// ���������
void __fastcall MSExcelWorks::RangeShtrich(Variant& wst, int firstRow, int firstCol, int CountRow, int lastCol, int Shtrich)
{
  int lastRow = firstRow + CountRow - 1; // ����� �������� ������
  Variant sell_left_top = wst.OlePropertyGet("Cells", firstRow, firstCol);
  Variant sell_right_bottom = wst.OlePropertyGet("Cells", lastRow, lastCol);
  Variant diap = wst.OlePropertyGet("Range", sell_left_top, sell_right_bottom);

  diap.OlePropertyGet("Interior").OlePropertySet("Pattern", Shtrich);
}

//------------------------------------------------------------------------------
// ����������� �����
Variant __fastcall MSExcelWorks::MergeCells(Variant& worksheet, int firstRow, int firstCol, int lastRow, int lastCol)
{
    //int lastRow = firstRow + CountRow - 1; // ����� �������� ������
    Variant sell_left_top = worksheet.OlePropertyGet("Cells", firstRow, firstCol);
    Variant sell_right_bottom = worksheet.OlePropertyGet("Cells", lastRow, lastCol);
    Variant range = worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);
    range.OlePropertySet("MergeCells", true);
    // Across - True, ����� ���������� ������ � ������ ������ ���������� ��������� ��� ��������� �������. �������� �� ��������� False.

    return range;
}

//------------------------------------------------------------------------------
//
inline int MSExcelWorks::GetRangeFirstRow(Variant range)
{
    return range.OlePropertyGet("Row");
}

//------------------------------------------------------------------------------
//
inline int MSExcelWorks::GetRangeFirstColumn(Variant range)
{
    return range.OlePropertyGet("Column");
}

//----------------------------------------------------------------------------
// ������������ ������� ������ ��������� ��� ������� (��������� Ctrl+D � MS Excel)
// �� ���������!!!!
void __fastcall MSExcelWorks::FillDown(Variant& range)
{
    range.OleProcedure("FillDown");
}

//----------------------------------------------------------------------------
// ������� ������ ����� �� ������� ����
void __fastcall MSExcelWorks::InsertRows(Variant& worksheet, int RowIndex, int RowsCount)
{
    if (RowsCount < 1)
    {
        return;
    }
    
    Variant Rows = worksheet.OlePropertyGet("Rows");
    //Variant Row = Rows.OlePropertyGet("Item", RowIndex, 5);

    try
    {
        Variant Row = Rows.OlePropertyGet("Range", (IntToStr(RowIndex) + ":" + IntToStr(RowIndex+RowsCount-1)).c_str());
        Row.OleProcedure("Insert", 0xFFFFEFE7, 0);
    }
    catch (Exception &e)
    {
        throw Exception("�� ������� �������� ������ � ��������.\n���������� ����������� ����� " + IntToStr(RowsCount) + ".");
    }

    // OleProcedure("Insert", xlDown, xlFormatFromLeftOrAbove);
    // xlDown = 0xFFFFEFE7,
    // xlToLeft = 0xFFFFEFC1,
    // xlToRight = 0xFFFFEFBF,
    // xlUp = 0xFFFFEFBE
    // xlFormatFromLeftOrAbove = 0
    // xlFormatFromRightOrBelow = 1
}


//----------------------------------------------------------------------------
// ������� ������ ����� � ���������� ������������ � range �� ������� ����
// ���������� ��������� �� ����� Excel ��������� - ����� ������������ ���� ��������... - �� ������� ����
Variant __fastcall MSExcelWorks::InsertRows(Variant& range)
{
    Variant Worksheet = range.OlePropertyGet("Worksheet");
    int RangeFirstRow = GetRangeFirstRow(range);
    int RangeFirstCol = GetRangeFirstColumn(range);
    int RangeRowsCount = GetRangeRowsCount(range);
    int RangeColsCount = GetRangeColumnsCount(range);

    range.OleProcedure("Insert", -4121, 0);  // xlDown

    return GetRange(Worksheet, RangeFirstRow, RangeFirstCol, RangeRowsCount, RangeColsCount);
}

//---------------------------------------------------------------------------
// ����������� � ������� ��������� range � ������� Row, Count ������������ ��������� ���������
Variant __fastcall MSExcelWorks::CopyRangeEx(Variant& worksheet, const Variant& range, int RowIndent, int ColIndent, bool fCopyData)
{
    int rowsCount = GetRangeRowsCount(range);
    int colsCount = GetRangeColumnsCount(range);

    int rangeFirstRow = GetRangeFirstRow(range);
    int rangeFirstCol = GetRangeFirstColumn(range);

    int firstRow = rangeFirstRow + rowsCount + RowIndent;
    int firstCol = rangeFirstCol + ColIndent;

    Variant range_new = CopyRange(worksheet, range, firstRow, firstCol, fCopyData);

    return range_new;
}

//---------------------------------------------------------------------------
// ����������� � ������� ��������� range � ������� Row, Count ������������ ��������� ���������
Variant __fastcall MSExcelWorks::CopyRangeEx(Variant& worksheet, AnsiString sRangeName, int RowIndent, int ColIndent, bool fCopyData)
{
    Variant range = GetRangeByName(worksheet, sRangeName);
    return CopyRangeEx(worksheet, range, RowIndent, ColIndent, fCopyData);
}

//---------------------------------------------------------------------------
// ����������� � ������� ��������� range � ������� Row, Count
Variant __fastcall MSExcelWorks::CopyRange(Variant& worksheet, const Variant& range, int Row, int Col, bool fCopyData)
{
    //Variant worksheet = range.OlePropertyGet("Worksheet");

    int rowsCount = GetRangeRowsCount(range);
    int colsCount = GetRangeColumnsCount(range);

    Variant range_new = GetRange(worksheet, Row, Col, rowsCount, colsCount);
    range_new.OleProcedure("Insert", -4121);  // ������� ������ ����� � �������� �� ������� ���� (xlDown)
    range_new = GetRange(worksheet, Row, Col, rowsCount, colsCount);

    range.OleProcedure("Copy", range_new);  // ����������� �� ��������� � ����� �������� (��� �������)

    if (!fCopyData)
    {
        range_new.OleProcedure("ClearContents");
    }

    return range_new;
}

/* ��������� ��� ������ */
String MSExcelWorks::getCellName(Variant cell)
{
    try
    {
        Variant temp = cell.OlePropertyGet("Name");
        return temp.OlePropertyGet("Name");
    }
    catch(...)  // ���� ������ �� ����� �����
    {
        return "";
    }
}

/* ��������� ��� ������ ��� �������� */
String MSExcelWorks::getCellName(Variant cell, const String& prefix)
{
    String cellName = getCellName(cell);
    return deletePrefix(cellName, prefix);
}

    /*int prefixLength = prefix.Length();     // ����� ��������

    // ���������� ��� ���� ��� ������ � dataSet
    // � ������ ���������� ������������� �������� � ���
    if ( prefixLength > 0 )                                    // ������ 2017-08-10
    {
        String dsFieldName = "";
        int cellNameLength = cellName.Length(); // �������� ����� ����� ������

        if (cellNameLength > prefixLength)
        {
            String cellNamePrefix = cellName.SubString(1, prefixLength);    // �������� ������� �� ��������� ����� ������
            if (cellNamePrefix == prefix)
            {
                dsFieldName = cellName.SubString(prefixLength+1, cellNameLength - prefixLength);
            }
            else
            {
                throw Exception("Prefix of cell name does't match with dataset prefix.");
            }

            return dsFieldName;
        }
        else
        {
            throw Exception("Length of cell name less whan lenght of prefix.");
        }
    }

    return cellName; */


/* ���������� ������ ��� �������� ����� �� dataSet � table
*/
std::vector<TLinkFields> MSExcelWorks::assignDataSetToRangeFields(Variant range, TDataSet* dataSet, const String& fieldNamePrefix)
{
    std::vector<TLinkFields> result;
    result.reserve(dataSet->FieldCount);

    int prefixLength = fieldNamePrefix.Length();
    Variant cells = range.OlePropertyGet("Cells");
    int n = cells.OlePropertyGet("Count");
    //int n = GetRangeColumnsCount(range);

    // ����  <= n 2017-07-28
    for (int i = 1; i <= n; i++ )         // ���� �� ����� � ������
    {

        Variant cell = cells.OlePropertyGet("Item", 1, i);
        String cellName = getCellName(cell, fieldNamePrefix);


        if (cellName != "")
        {
            // ������������ ���� � ����� � ���������
            //int fieldIndex = field.OlePropertyGet("Index");   //  ��� �������� �������� �������� ��������� �� ���� ��������

            TField* fieldOfDataSet = dataSet->Fields->FindField(cellName);
            if ( fieldOfDataSet != NULL)
            {
                result.push_back(std::make_pair(fieldOfDataSet->FieldNo, i));
            }
        }
    }

    return result;
}

/* ����� 2017-02-15 */
void MSExcelWorks::writeDataSetToSingleRange(Variant worksheet, TDataSet* dataSet, const String& fieldNamePrefix)
{
   //GetRangeByName(range, "fa_pack_id");
   //GetRangeByName(worksheet, "t");
   //Variant Parent = worksheet.OlePropertyGet("Parent");

    //std::vector<String> fieldNames = GetNamesFromWorkbook(worksheet);
    std::vector<TNamedRange> fieldNames = GetNamesFromObject(worksheet, fieldNamePrefix);

    for (std::vector<TNamedRange>::iterator it = fieldNames.begin(); it != fieldNames.end(); it++)
    {
        TField* dataSetField = dataSet->Fields->FindField( (*it).first );
        if ( dataSetField )
        {
            //worksheet.OlePropertyGet("Parent");
            //Variant range = GetRangeByName(worksheet, *it);
            WriteToRange( (*it).second, dataSetField->AsString);
        }
    }

    /*
        std::vector<String> vExcelFields = GetNamesFromWorkbook(Workbook);

        for (std::vector<String>::iterator itExcelField = vExcelFields.begin(); itExcelField < vExcelFields.end(); itExcelField++)
        {
            TField* pField = QTable->Fields->FindField(*itExcelField);
            if (pField)
            {
                try {
                    Variant range = GetRangeByName(Worksheet, *itExcelField);
                    // ���������������� 2016-11-17
                    //if ( !VarIsEmpty(range) ) {
                    if ( !VarIsClear(range) )
                    {
                        WriteToRange(QTable->FieldByName(*itExcelField)->AsString, range);
                    }
                }
                catch (...)
                {
                }
            }
        } */


}

/* ����� 2017-02-15 */
Variant MSExcelWorks::writeDataSetToTableRange(Variant tableRange, TDataSet* dataSet, const String& fieldNamePrefix)
{
    std::vector<TLinkFields> links = assignDataSetToRangeFields(tableRange, dataSet, fieldNamePrefix);


    if ( links.size() == 0 || !dataSet->Active || dataSet->Eof)
    {
        return Variant::Empty();
    }

    Variant cells = tableRange.OlePropertyGet("Cells");
    int RecordCount = dataSet->RecordCount;
    int RangeColumnsCount = cells.OlePropertyGet("Count");

    Variant data_body = CreateVariantArray(RecordCount, RangeColumnsCount);
    VarArrayLock(data_body);

    int j = 1;
    for (; !dataSet->Eof; dataSet->Next() )
    {

        //rowTmp.OlePropertyGet("Range").OlePropertySet("FormattedText" , formatedText );
        //Variant fields = rows.OleFunction("Item", currentRow++).OlePropertyGet("Range").OlePropertyGet("Fields");

        for (std::vector<TLinkFields>::iterator it = links.begin(); it != links.end(); it++)
        {
            String value = dataSet->Fields->FieldByNumber(it->first)->AsString;
            data_body.PutElement(value, j, it->second);

            //Variant field = fields.OleFunction("Item", it->second);
            //field.OlePropertyGet("Result").OlePropertySet("Text", dataSet->Fields->FieldByNumber(it->first)->AsString.c_str());
        }
        j++;
    }

    VarArrayUnlock(data_body);


    // ������� �� ����, � ������� rPos, cPos
    //Variant range = WriteTable(Worksheet, data_body, RangeFirstRow, RangeFirstColumn);
    Variant range = WriteTableToRange(tableRange, data_body, 1, 1, true);

    // ������������ ������
    VarClear(data_body);

    return range;



    /*int j = 1;
    for (QTable->First(); !QTable->Eof; QTable->Next() )
    {
        for (int i = 1; i <= RangeColumnsCount; i++)
        {
            String value;

            if ( bindingList[i-1].second )
            {
                value = QTable->FieldByName( bindingList[i-1].first )->AsString;
            }
            else
            {
                value = "NA:" + bindingList[i-1].first;
            }

            data_body.PutElement(value, j, i);
        }
        j++;
    }                   */
}






//---------------------------------------------------------------------------
// ��������� ����������� ������� � ������������ ���������
// ���������������� ���������� �� TOraQuery
// ���������� �������, � ���������� ������� !!!!!!!!!!!!!!!!!!!!!!!!!!
Variant MSExcelWorks::ExportToExcelTable(TDataSet* QTable, Variant Worksheet, String RangeName, bool fUnbounded)
{
    // �������� range, ���� ���������� �������� �������,
    // �������� ������ � ������� �������� � ���� range
    Variant range_body = GetRangeByName(Worksheet, RangeName);
    int RangeColumnsCount = GetRangeColumnsCount(range_body);
    int RangeRowsCount = GetRangeRowsCount(range_body);
    int RangeFirstRow = GetRangeFirstRow(range_body);
    int RangeFirstColumn = GetRangeFirstColumn(range_body);

    int RecordCount = QTable->RecordCount;


    // �������� �� ��, ��� ���������� ����� � ��������� �� ��������� ���-�� ����� � ��������� Range
    if (RecordCount > RangeRowsCount)
    {
        if (fUnbounded)
        {
            Variant range_new = GetRange(Worksheet, RangeFirstRow+1, RangeFirstColumn, RecordCount - 1, RangeColumnsCount);
            range_new = InsertRows(range_new);          // ��������� ������ �� ������� ���� (���������� � ������ ����� - ������ ������)
            CopyRangeFormat(range_body, range_new);     // �������� ����� ����� (���������� ����� ������)

            int rowHeight = GetRowHeight(range_body);   // �������� ������ ����� (������ �� ���������� CopyRangeFormat)
            SetRowHeight(range_new, rowHeight);
        }
        else
        {
            throw Exception("Error. The source dataset contains too much records.");
        }
    }


    // ��������� ������� vector<String>
    // ��� ������ �������� - ������ ������ � range_body?
    // �������� �������� - ��� ���� � TQuery (� ��� ������ � range_body)

    // ���� �� ��������� range_body � ��������� ������ �� Names
    // ��������� ������ ������� ����� (= ������ ����� � range_body)
    //std::vector<String> vs;
    //vs.reserve(RangeColumnsCount);


    std::vector< std::pair<String, bool> > bindingList; // ������ �������� -  ��� ����, ������ - ������� ��������� ������
    bindingList.reserve(RangeColumnsCount);


    Variant Cells = range_body.OlePropertyGet("Cells");
    for (int i = 1; i <= RangeColumnsCount; i++)
    {
        Variant Cell = Cells.OlePropertyGet("Item", 1, i);

        String CellName;
        try
        {
            CellName = Cell.OlePropertyGet("Name").OlePropertyGet("Name");
        }
        catch (Exception &e)
        {
            throw Exception("Error. The range is not solid. Cell in column " + IntToStr(i) + " is not named."); // �� ��������� ��� ������ � Range
        }

        TField* field = QTable->Fields->FindField(CellName);
        bindingList.push_back(std::pair<String, bool>(CellName, field != NULL ));

        /*
        if (field)     // ���������, ���������� �� ���� � ���� ������ � ����� ���������
        {
            vs.push_back(CellName);
        }
        else
        {
            // commented 2016-01-24
            //throw Exception("Error. Not enough field " + CellName + ".");
        } */
    }

    // ��������� ������ data_body ���������� �� QTable
    Variant data_body = CreateVariantArray(RecordCount, RangeColumnsCount);
    VarArrayLock(data_body);
    int j = 1;
    for (QTable->First(); !QTable->Eof; QTable->Next() )
    {
        for (int i = 1; i <= RangeColumnsCount; i++)
        {
            String value;

            if ( bindingList[i-1].second )
            {
                value = QTable->FieldByName( bindingList[i-1].first )->AsString;
            }
            else
            {
                value = "NA:" + bindingList[i-1].first;
            }

            data_body.PutElement(value, j, i);
        }
        j++;
    }
    VarArrayUnlock(data_body);


    // ������� �� ����, � ������� rPos, cPos
    Variant range = WriteTable(Worksheet, data_body, RangeFirstRow, RangeFirstColumn);

    // ������������ ������
    VarClear(data_body);

    return range;

}

//---------------------------------------------------------------------------
// �������� �������� ����������� ����������
// �� ��������������� �������� �� TOraQuery
void MSExcelWorks::ExportToExcelFields(TDataSet* QTable, Variant Worksheet)
{
    try
    {
        // �������� range, ���� ���������� �������� �������,
        // �������� ������ � ������� �������� � ���� range
        //Variant range_body = msexcel.GetRangeByName(Worksheet, RangeName);
        //int FieldCount = msexcel.GetRangeColumnsCount(range_body);

        if ( !(QTable != NULL && QTable->RecordCount > 0) )
        {
            return;
        }

        Variant Workbook = Worksheet.OlePropertyGet("Parent");
        std::vector<String> vExcelFields = GetNamesFromWorkbook(Workbook);

        for (std::vector<String>::iterator itExcelField = vExcelFields.begin(); itExcelField < vExcelFields.end(); itExcelField++)
        {
            TField* pField = QTable->Fields->FindField(*itExcelField);
            if (pField)
            {
                try {
                    Variant range = GetRangeByName(Worksheet, *itExcelField);
                    // ���������������� 2016-11-17
                    //if ( !VarIsEmpty(range) ) {
                    if ( !VarIsClear(range) )
                    {
                        WriteToRange(QTable->FieldByName(*itExcelField)->AsString, range);
                    }
                }
                catch (...)
                {
                }
            }
        }
    }
    catch(...)
    {
    }
}

//---------------------------------------------------------------------------
// ��������� ���������� ����
// ������������ ��� ��������� ��� ������ ������
void MSExcelWorks::BeforeUpdate()
{
    ExcelApp.OlePropertySet("ScreenUpdating", true);
}

//---------------------------------------------------------------------------
// ��������� ���������� ����
// ������������ ��� ��������� ��� ������ ������
void MSExcelWorks::AfterUpdate()
{
    ExcelApp.OlePropertySet("ScreenUpdating", true);
}


/*
//---------------------------------------------------------------------------
// ������ ������ �� ������ ������� Excell
AnsiString __fastcall ReadCell(Variant wst, int Row, int Col)
{
  return(Trim(wst.OlePropertyGet("Cells").OlePropertyGet("Item", Row, Col)));
}

//---------------------------------------------------------------------------
// �������� ��������
Variant __fastcall actPage(Variant appl)
{
  return(appl.OlePropertyGet("ActiveSheet"));
}



    //Variant Application = WorksheetOrWorkbook.OlePropertyGet("Application");
    //Variant Parent = WorksheetOrWorkbook.OlePropertyGet("Parent");
    //if (VarType(Parent) == VarType(Application)) {      // ���� �������� Workbook


*/


//---------------------------------------------------------------------------


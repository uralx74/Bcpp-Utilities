//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "MainFormUnit.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TMainForm *MainForm;
//---------------------------------------------------------------------------
__fastcall TMainForm::TMainForm(TComponent* Owner)
    : TForm(Owner)
{
}
//---------------------------------------------------------------------------

void __fastcall TMainForm::Button1Click(TObject *Sender)
{
    MainDataModule->EsaleSession->Open();
    MainDataModule->Raion->Open();



  /*  TWordExportParams wordExportParams;
    //wordExportParams.resultFilename = ExtractFilePath(resultFilename) + ExtractFileName(resultFilename) + "_[:counter].doc";
    wordExportParams.templateFilename = "c:\\PROGRS\\current\\GroupDoc\\report\\template_document_notice.dotx";
    //wordExportParams.resultFilename = "c:\\PROGRS\\current\\GroupDoc\\report\\document_notice_[:counter].doc";
    wordExportParams.pagePerDocument = 500;
 */


    MSWordWorks msword;// = new MSWordWorks();
    Variant Document;   // Шаблон

    Variant word = msword.OpenWord();

    #ifndef NDEBUG
    msword.SetVisible(true);
    msword.SetDisplayAlerts(true);
    #endif

    Document =  msword.OpenDocument("c:\\_PROGRS\\current\\util\\Test\\WordTest\\report\\template_document_stop.dotx", false);

    /*Variant tables = Document.OlePropertyGet("Tables");
    Variant table = tables.OleFunction("Item", 1);

    MainDataModule->Raion->Open();
    msword.writeDataSetToTable(table, MainDataModule->Raion);*/



    String imagePath = "c:\\_PROGRS\\current\\Sweety\\report\\visa\\Р.М. Юсупов.png";

    Variant fields = Document.OleFunction("Range").OlePropertyGet("Fields");
    Variant field = fields.OleFunction("Item", 1 );


    Variant rangeResult = field.OlePropertyGet("Result");
    Variant inlineShape = msword.SetPictureToRange(Document, rangeResult, imagePath);

    Variant shape = msword.ConverInlineShapeToShape(inlineShape, 5);


    //shape.OlePropertyGet("WrapFormat").OlePropertySet("Type", 1); //WdWrapType::wdWrapThrough);





    //msword.SetShapePos(shape, 10, 10);

}


//---------------------------------------------------------------------------


void __fastcall TMainForm::Button2Click(TObject *Sender)
{
    _shape.OlePropertySet("RelativeVerticalPosition", wf++);
    //_shape.OlePropertyGet("WrapFormat").OlePropertySet("Type", wf++);
}
//---------------------------------------------------------------------------


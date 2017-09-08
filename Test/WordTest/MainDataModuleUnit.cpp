//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "MainDataModuleUnit.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "DBAccess"
#pragma link "MemDS"
#pragma link "Ora"
#pragma link "DBAccess"
#pragma link "DBAccess"
#pragma resource "*.dfm"
TMainDataModule *MainDataModule;
//---------------------------------------------------------------------------
__fastcall TMainDataModule::TMainDataModule(TComponent* Owner)
    : TDataModule(Owner)
{
}
//---------------------------------------------------------------------------

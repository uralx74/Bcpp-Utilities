//---------------------------------------------------------------------------

#ifndef MainDataModuleUnitH
#define MainDataModuleUnitH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include "MemDS.hpp"
#include "Ora.hpp"
#include <DB.hpp>
#include "DBAccess.hpp"

//---------------------------------------------------------------------------
class TMainDataModule : public TDataModule
{
__published:	// IDE-managed Components
    TOraQuery *PISMA;
    TOraSession *CcbSession;
    TOraQuery *Raion;
    TOraSession *EsaleSession;
private:	// User declarations
public:		// User declarations
    __fastcall TMainDataModule(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TMainDataModule *MainDataModule;
//---------------------------------------------------------------------------
#endif

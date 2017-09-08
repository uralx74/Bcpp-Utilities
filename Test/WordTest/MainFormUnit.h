//---------------------------------------------------------------------------

#ifndef MainFormUnitH
#define MainFormUnitH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <vector>
#include "MainDataModuleUnit.h"
#include "DocumentWriter.h"
#include "Ora.hpp"
//#include "MSWordWorks.h"

//---------------------------------------------------------------------------
class TMainForm : public TForm
{
__published:	// IDE-managed Components
    TButton *Button1;
    TButton *Button2;
    void __fastcall Button1Click(TObject *Sender);
    void __fastcall Button2Click(TObject *Sender);
private:	// User declarations
public:		// User declarations
    __fastcall TMainForm(TComponent* Owner);
    Variant _shape;
    int wf;
    //void assignDataSetToTableFields(Variant table, TDataSet* dataSet);
    //std::vector<std::pair<int, int> > assignDataSetToTableFields(Variant table, TDataSet* dataSet);

};
//---------------------------------------------------------------------------
extern PACKAGE TMainForm *MainForm;
//---------------------------------------------------------------------------
#endif

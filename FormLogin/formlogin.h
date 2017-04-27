//---------------------------------------------------------------------------

#ifndef FormLoginH
#define FormLoginH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <Mask.hpp>
#include <ExtCtrls.hpp>
#include <Graphics.hpp>
#include <Buttons.hpp>
#include "DBAccess.hpp"
#include "Ora.hpp"
#include <Db.hpp>
#include "Registry.hpp"
#include "..\util\taskutils.h"
#include "..\util\keyboardutil.h"
#include <ActnList.hpp>


/*

class TLoginForm;  // опережающее объ€вление


class TLoginFormDestroyer
{
private:
    TLoginForm* p_instance;
public:    
    ~TLoginFormDestroyer();
    void initialize( TLoginForm* p );
};  */





//---------------------------------------------------------------------------
class TLoginForm : public TForm
{
__published:	// IDE-managed Components
    TBitBtn *CancelBtn;
    TGroupBox *GroupBox1;
    TLabel *Label1;
    TLabel *Label2;
    TMaskEdit *PasswordMaskEdit;
    TEdit *UsernameEdit;
    TImage *Image1;
    TActionList *ActionList1;
    TAction *Login;
    TAction *Cancel;
    TTimer *Timer1;
    TPanel *KBLayoutPanel;
    TBitBtn *LoginBtn;
    void __fastcall FormShow(TObject *Sender);
    void __fastcall FormCreate(TObject *Sender);
    void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
    void __fastcall Timer1Timer(TObject *Sender);
    void __fastcall KBLayoutPanelClick(TObject *Sender);
    void __fastcall cancelAction(TObject *Sender);
    void __fastcall loginAction(TObject *Sender);
    
private:	// User declarations
    //static TLoginForm* p_instance;
    //static TLoginFormDestroyer destroyer;
    TLoginForm* _loginForm;


    AnsiString AppName;
    TKeyboardUtil KeyboardUtil;
    TOraSession *_session;
    TOraQuery* _rolesQuery;

    bool isSessionAssigned;
    int _retryCount;

protected:
    TLoginForm(TLoginForm const&);
    TLoginForm& operator= (TLoginForm const&);



    friend class TOraLoggerDestroyer;      // for access to p_instance

public:		// User declarations
    __fastcall TLoginForm(TComponent* Owner, TOraSession* const session, bool assignConnect = false);
    __fastcall ~TLoginForm();
    String getUsername();
    String getPassword();
    bool __fastcall execute();
    bool __fastcall checkRole(const String& role);

    std::vector<AnsiString>* __fastcall GetUserPriveleges();

};
//---------------------------------------------------------------------------
extern PACKAGE TLoginForm *LoginForm;
//---------------------------------------------------------------------------
#endif

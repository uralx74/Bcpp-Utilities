/*---------------------------------------------------------------------------
    TLoginForm LoginForm = new TLoginForm(Application);
    bool loggedon = LoginForm->Execute(OraSessionAuth);
    LoginForm->Free();

    if (loggedon) {
        Username =  UpperCase(Trim(OraSessionAuth->Username));
        OraSessionAuth->Disconnect();
        delete OraSessionAuth;
        return true;
    } else {
        return false;
    }


    TLoginForm& instance = TLoginForm::Instance();
//---------------------------------------------------------------------------*/

#include <vcl.h>
#pragma hdrstop

#include "FormLogin.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "DBAccess"
#pragma link "Ora"
#pragma resource "*.dfm"

const defaul_retry_count = 3;

/**/
__fastcall TLoginForm::TLoginForm(TComponent* Owner, TOraSession* const session, bool assignConnect)
    : TForm(Owner),
    _retryCount(defaul_retry_count),
    isSessionAssigned(assignConnect)
{
    if ( assignConnect )
    {
        _session = new TOraSession(this);
        _session->AssignConnect(session);
        _session->Username = "";
        _session->Password = "";
    }
    else
    {
        _session = session;
    }
    _session->Connected = false;
    _session->LoginPrompt = false;
}

/**/
__fastcall TLoginForm::~TLoginForm()
{
    /*if (isSessionAssigned)
    {
        delete _session;
    }*/
}

/**/
bool __fastcall TLoginForm::execute()
{
    //bool loggedon = LoginForm->Execute(Session);
    //_username = UpperCase(_username);
    //AddSystemVariable("username", _username);

    if (_session->Username == "")
    {
        this->ShowModal();
    }
    else
    {
        //_session->Username = Username;
        //_session->Password = Password;
        try
        {
            _session->Connect();
        }
        catch (...)
        {
        }
    }

    // Если удалось подключиться, то получаем список ролей
    if (_session->Connected)
    {
        _rolesQuery = new TOraQuery(this);
        _rolesQuery->Session = this->_session;
        _rolesQuery->SQL->Text = "select * from USER_ROLE_PRIVS";
        _rolesQuery->Open();
        _rolesQuery->Filtered = true;
    }

    return _session->Connected;
}

/**/
String TLoginForm::getUsername()
{
    return _session->Username;
}

/**/
String TLoginForm::getPassword()
{
    return _session->Password;
}

/**/
void __fastcall TLoginForm::FormShow(TObject *Sender)
{
    if (UsernameEdit->Text != "")
    {
        PasswordMaskEdit->SetFocus();
    }
    else
    {
        UsernameEdit->SetFocus();
    }
}

/**/
void __fastcall TLoginForm::FormCreate(TObject *Sender)
{
    AppName = "Software\\CES\\" + Application->Title;

    try
    {
        TRegistry* pReg = new TRegistry();
        pReg->RootKey = HKEY_CURRENT_USER;

        if (pReg->OpenKeyReadOnly(AppName))
        {
            UsernameEdit->Text = pReg->ReadString("Username");
            pReg->CloseKey();
        }
        else if ( pReg->OpenKeyReadOnly("Software\\Devart\\ODAC\\Connect\\" + Application->Title) )
        {
            UsernameEdit->Text = pReg->ReadString("Username");
            pReg->CloseKey();
        }
        delete pReg;

        KBLayoutPanel->Color = RGB(0,128,255);
    }
    catch (...)
    {
    }
}

/**/
void __fastcall TLoginForm::cancelAction(TObject *Sender)
{
    this->Close();
}

/* Выполнение попытки авторизоваться
*/
void __fastcall TLoginForm::loginAction(TObject *Sender)
{
    static int TryNumber = 0;

    _session->Username = UsernameEdit->Text;
    _session->Password = PasswordMaskEdit->Text;
    PasswordMaskEdit->Text = "";

    try
    {
        _session->Connected = true;
        this->Close();
    }
    catch (EOraError &DatabaseError)
    {
        _session->Password = "";
        MessageBoxStop(DatabaseError.Message);
        TryNumber++;
        if (TryNumber < _retryCount)
        {
            PasswordMaskEdit->SetFocus();
            //PasswordMaskEdit->SelStart = 0;
            //PasswordMaskEdit->SelLength = -1;
        }
        else
        {
            //this->_session = NULL;
            this->Close();
        }
    }
}

/* Проверка наличия роли */
bool __fastcall TLoginForm::checkRole(const String& role)
{
    if (_rolesQuery != NULL && _rolesQuery->Active )
    {
        _rolesQuery->Filter = "GRANTED_ROLE = '" + UpperCase(role) + "'";
        return _rolesQuery->RecordCount;
    }
    else
    {
        return false;
    }

    /*TOraQuery* Query = new TOraQuery(this);
    Query->Session = this->_session;
    Query->SQL->Text = "select * from ALL_TAB_PRIVS where GRANTEE = upper(:P_ROLE)";
    Query->Params->ParamValues["P_ROLE"] = role;
    Query->Open();

    bool result = Query->FieldByName("grantee")->AsString != "";

    Query->Close();
    delete Query;

    /*if (!result)
    {
        MessageBoxStop("Вы не имеете прав доступа к этой программе. \n\nОтказано в доступе.");
    } */
     /*
    return result;  */
}

/* Список доступных пользователю ролей */
std::vector<AnsiString>* __fastcall TLoginForm::GetUserPriveleges()
{

    TOraQuery* Query = new TOraQuery(NULL);
    Query->Session = this->_session;
    Query->SQL->Text = "select * from SESSION_ROLES";
    //Query->SQL->Text = "SELECT DISTINCT GRANTEE FROM ALL_TAB_PRIVS";
    Query->FetchAll = true;
    Query->Open();

    std::vector<AnsiString>* result = new std::vector<AnsiString>;
    result->reserve(Query->RecordCount);

    for (;!Query->Eof ; Query->Next())
    {
        result->push_back(Query->FieldByName("ROLE")->AsString);
    }

    Query->Close();
    delete Query;

    return result;
}

/**/
void __fastcall TLoginForm::FormClose(TObject *Sender, TCloseAction &Action)
{
    try
    {
        if (this->_session != NULL && this->_session->Connected)
        {
            TRegistry* pReg = new TRegistry();
            pReg->RootKey = HKEY_CURRENT_USER;
            if (pReg->OpenKey(AppName, true))
            {
                pReg->WriteString("Username", this->_session->Username);
                pReg->CloseKey();
            }
            delete pReg;
        }
    }
    catch (...)
    {
    }
    Timer1->Enabled = false;
}

/**/
void __fastcall TLoginForm::Timer1Timer(TObject *Sender)
{
    KBLayoutPanel->Caption = KeyboardUtil.GetLayout();

    //TKeyboardState::GetKeyboardState(VK_CAPITAL);

    /*if (ku.GetKeyboardState())
    {
    }
    else
    {
    }*/

}

/**/
void __fastcall TLoginForm::KBLayoutPanelClick(TObject *Sender)
{
    KeyboardUtil.SetNextLayout();
}




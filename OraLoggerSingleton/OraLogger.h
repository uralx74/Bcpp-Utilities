//---------------------------------------------------------------------------
#ifndef OraLoggerH
#define OraLoggerH

class TOraLogger;  // опережающее объ€вление


class TOraLoggerDestroyer
{
private:
    TOraLogger* p_instance;
public:    
    ~TOraLoggerDestroyer();
    void initialize( TOraLogger* p );
};

//------------------------------------------------------------------------------
//
class TOraLogger
{
private:
    static TOraLogger* p_instance;
    static TOraLoggerDestroyer destroyer;

    AnsiString _osMacAddress;
    AnsiString _osUsername;
    AnsiString _taskUsername;
    AnsiString _taskName;
    AnsiString _appVer;
    TOraSession* _oraSession;
    TOraQuery* _oraQueryLog;
    AnsiString _appId;

protected:
    TOraLogger();
    ~TOraLogger();
    friend class TOraLoggerDestroyer;      // for access to p_instance

public:
    //TOraLogger(TOraSession* OraSession, AnsiString s_os_mac_address, AnsiString s_task_user_name, AnsiString s_task_name, AnsiString s_app_ver);
    //~TOraLogger();
    void setParameters(const AnsiString& funcName, const AnsiString& threadId = "", AnsiString descr = "");
    bool Write(const AnsiString& funcName, const AnsiString& threadId = "", AnsiString descr = "");
};
//---------------------------------------------------------------------------
#endif
 
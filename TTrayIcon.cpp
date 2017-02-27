#include "TTrayIcon.h"


TTrayIcon::TTrayIcon(HWND MsgWnd, WORD iconIndex, HICON icon, String tip, UINT message)
{
    BOOL MyTaskBar;
    UINT uID = WM_USER + 26;
    UINT MyNotifyID = WM_USER + 27;

    //Данные для иконки на панели задач
    _nid.cbSize = sizeof(NOTIFYICONDATA);
    _nid.hWnd = MsgWnd;
    _nid.uID = uID;
    _nid.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP;
    _nid.uCallbackMessage = MyNotifyID;
    _nid.hIcon = icon;

    SetTip(tip);
}

TTrayIcon::~TTrayIcon()
{
}


/* Создает иконку*/
bool TTrayIcon::TaskBarAddIcon()
{
    bool res = Shell_NotifyIcon(NIM_ADD, &_nid);

    /*if(hIcon)
    {
        DestroyIcon(hIcon);
    }*/
    return res;
}

/* Удаляет иконку */
bool TTrayIcon::TaskBarDeleteIcon()
{
    return Shell_NotifyIcon(NIM_DELETE, &_nid);
}


/* Отображает иконку */
void __fastcall TTrayIcon::Show()
{
    TaskBarAddIcon();

	/*TIcon* icon = new TIcon();
    //icon->Assign(Application->Icon);
    icon->Assign(Image2->Picture->Icon);
    TaskBarAddIcon(this->Handle, 0, icon->Handle, "AdF", 1);*/
}

void __fastcall TTrayIcon::Hide()
{
	TaskBarDeleteIcon();
}

void __fastcall TTrayIcon::SetTip(const String& tip)
{
    StrCopy(_nid.szTip, tip.c_str());
}

void __fastcall TTrayIcon::SetIconIndex(int iconIndex)
{
}

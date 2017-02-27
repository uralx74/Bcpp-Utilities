/*
  Класс для отображения иконки в трее

  Автор: vsovchinnikov


Использование:
    TIcon* icon = new TIcon();
    //icon->Assign(Application->Icon);
    icon->Assign(Image1->Picture->Icon);
    trayIcon = new TTrayIcon(this->Handle, 0, icon->Handle, "My application", 1);
    trayIcon->Show();
*/

#ifndef TTrayIconH
#define TTrayIconH

#include <Graphics.hpp>

class TTrayIcon
{
private:
	NOTIFYICONDATA _nid;
	bool TaskBarAddIcon();
	bool TaskBarDeleteIcon();

public:
	TTrayIcon(HWND MsgWnd, WORD iconIndex, HICON icon, String tip, UINT message);
	~TTrayIcon();
	void __fastcall Show();
	void __fastcall Hide();
	void __fastcall SetTip(const String& tip);
	void __fastcall SetIconIndex(int iconIndex);
};

#endif // TTrayIconH

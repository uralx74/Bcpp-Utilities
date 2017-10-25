#ifndef HACK_CTRL
#define HACK_CTRL

/*
hack_ctrl.h
author: vsovchinnikov
e-mail: utnpsys@gmail.com
*/

#include <Classes.hpp>
#include <Controls.hpp>

namespace HackCtrl
{

/* ������ ���� Enabled �������� ��������� ����������
*/
void switchEnabledGroupBox(TGroupBox* groupBox)
{
    bool isEnabled = groupBox->Enabled;

    for (int i = 0; i < groupBox->ControlCount; i++)
    {
        groupBox->Controls[i]->Enabled = isEnabled;
    }
}

/* ����������� ������� ���������� � ���������� ������� (� ������� �� ��������)*/
void __fastcall RealignControls(TWinControl *parent)
{
    TControl *c;
    TAlign align;
    for(int i=0; i < parent->ControlCount; i++)
    {
        c = parent->Controls[i];
        align = c->Align;
        switch(align)
        {
        case alLeft:
        {
            c->Left = parent->Width;
            break;
        }
        case alRight:
        {
            c->Left = 0;
            break;
        }
        case alTop:
        {
            c->Top = parent->Height;
            break;
        }
        case alBottom:
        {
            c->Top = 0;
            break;
        }
        }
        c->Align = align;
    }
}

/* ���������� ���������� ������������ ��������� � ��������-���������� */
int __fastcall GetVisibleControlCount(TCustomControl *ContainerControl)
{
    int ControlCount = 0;
    for(int i=0; i < ContainerControl->ControlCount; i++)
    {
        if( ContainerControl->Controls[i]->Visible )
        {
            ControlCount++;
        };
    }
    return ControlCount;
}
//---------------------------------------------------------------------------


/* ��������� ������ ������� ������� */
int PageIndexFromTabIndex(TPageControl* pageControl, int tabIndex)
{
    int visiblePageCount = 0;
    for (int i = 0; i <= pageControl->PageCount; i++)
    {
        if ( pageControl->Pages[i]->TabVisible )
        {
            visiblePageCount++;
        }
        if (visiblePageCount > tabIndex)
        {
            return i;
        }
    }
    return -1;
}

}
#endif

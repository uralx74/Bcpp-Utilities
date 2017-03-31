#ifndef HACK_CTRL
#define HACK_CTRL

#include <Classes.hpp>
#include <Controls.hpp>

/* Меняет флаг Enabled дочерних элементов управления
*/
void switchEnabledGroupBox(TGroupBox* groupBox)
{
    bool isEnabled = groupBox->Enabled;

    for (int i = 0; i < groupBox->ControlCount; i++)
    {
        groupBox->Controls[i]->Enabled = isEnabled;
    }
}

/* Выравнивает элеметы управления в правильном порядке */
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


#endif

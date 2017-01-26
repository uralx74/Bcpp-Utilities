#ifndef COLOR_LIST_H
#define COLOR_LIST_H

#include <vector>

class ColorList
{
private:
    std::vector<TColor> _colorList; // Вектор цветов вкладок 

public:
    ColorList()
    {
        // Список цветов вкладок
        _colorList.push_back(static_cast<TColor>(RGB(180,255,20)));     // green
        _colorList.push_back(static_cast<TColor>(RGB(120,230,90)));     // green
        _colorList.push_back(static_cast<TColor>(RGB(0,190,90)));       // green
        _colorList.push_back(static_cast<TColor>(RGB(0,190,210)));      // blue
        _colorList.push_back(static_cast<TColor>(RGB(90,225,255)));     // blue
        _colorList.push_back(static_cast<TColor>(RGB(100,176,255)));    // blue
        _colorList.push_back(static_cast<TColor>(RGB(200,145,255)));    // violet
        _colorList.push_back(static_cast<TColor>(RGB(255,100,220)));    // violet
        _colorList.push_back(static_cast<TColor>(RGB(255,130,170)));    // red light
        _colorList.push_back(static_cast<TColor>(RGB(255,100,0)));      // red
        _colorList.push_back(static_cast<TColor>(RGB(255,180,50)));     // orange
        _colorList.push_back(static_cast<TColor>(RGB(255,255,0)));      // yellow

    }
    void addColor(TColor color)
    {
        _colorList.push_back(color);
    }

    void clear()
    {
        _colorList.clear();
    }

    TColor getColorByIndex(int index)
    {
        return _colorList[index % _colorList.size()];
    }
};

#endif
//---------------------------------------------------------------------------
#pragma hdrstop

#include "JsonObject.h"
#pragma package(smart_init)


__fastcall TJsonObject::TJsonObject()
{
}

__fastcall TJsonObject::TJsonObject(const String& text):
    _text(text),
    _resultCode(RT_UNDEFINED)
{
    _textLength = text.Length();
}

/* Reading node name
*/
String TJsonObject::readNode(const String& str, int& pos)
{
    if (pos < 0)
    {
        return NULL;
    }

    int p1 = -1;
    int p2 = 0;

    // Search the opening brace
    for (int i = pos; i < _textLength; i++)
    {
        if ( str[i] != ' ' )   // Условие выхода
        {
            if ( str[i] == '{' )   // Условие выхода
            {
                p1 = i;
                break;
            }
            else
            {
                _resultCode = RT_END_NODE;
                return NULL;
            }
        }
    }

    if (p1 < 0)
    {
        _resultCode = RT_END_TEXT;
        return NULL;
    }

    // Search the opening double quote
    for (int i = p1; ; i++)
    {
        if (i > _textLength)   // Условие выхода
        {
            _resultCode = RT_END_TEXT;
            pos = -1;
            return NULL;
        }

        if (str[i] == '\"')
        {
            p1 = i+1;
            break;
        }
    }

    for (int i = p1; i <= _textLength ;i++)
    {
        if (i > _textLength)   // Условие выхода
        {
            _resultCode = RT_END_TEXT;
            pos = -1;
            return NULL;
        }

        if (str[i] == '\"')
        {
            p2 = i;
            break;
        }
    }

    pos = p2 + 1;
    return str.SubString(p1, p2-p1);
}

/* Reading param name
*/
String TJsonObject::readParamName(const String& str, int& pos)
{
    if (pos < 0)
    {
        return NULL;
    }

    int p1 = pos;
    int p2 = 0;

    // Начало имени
    for (int i = p1; i < _textLength ;i++)
    {
        // Условие выхода
        if (str[i] == '}')
        {
            //pos = -1;
            pos = i + 1;

            _resultCode = RT_END_PARAMETERS;
            return NULL;
        }

        if (str[i] == '\"')
        {
            p1 = i+1;
            break;
        }
    }
    // Конец имени
    for (int i = p1; ;i++)
    {
        // Условие выхода
        if (str[i] == '}')
        {
            //pos = -1;
            return NULL;
        }
        if (str[i] == '\"')
        {
            p2 = i;
            break;
        }
    }
    pos = p2 + 1;
    return str.SubString(p1, p2-p1);
}

/* Reading param value
*/
Variant TJsonObject::readParamValue(const String& str, int& p1, TJsonDataType::Type& type)
{
    if (p1 < 0)
    {
        return NULL;
    }
    int p2 = 0;

    String result;


    for (int i = p1; ;i++)
    {
        if (str[i] == '}')
        {
            _resultCode = RT_END_PARAMETERS;
            return NULL;
        }
        if (str[i] != ' ' && str[i] != '=')
        {
            if (str[i] == '\"')
            {
                type = TJsonDataType::PT_STRING;
                p1 = i + 1;

                Variant result;
                VariantChangeType(result, result, 0, VT_LPSTR);
                result = readParamValueString(str, p1);

                return result;
            }
            else
            {
                type = TJsonDataType::PT_INTEGER;
                p1 = i;

                Variant result;
                VariantChangeType(result, result, 0, VT_I4);
                result = StrToInt(readParamValueInteger(str, p1));

                return result;
            }
            break;
        }
    }

    return str.SubString(p1, p2-p1);
}

/* Reading String value (quoted)
*/
Variant TJsonObject::readParamValueString(const String& str, int& pos)
{
    int p1 = pos;
    int p2 = 0;

    for (int i = p1; ;i++)
    {
        if (str[i] == '\"')
        {
            p2 = i;
            break;
        }
    }
    pos = p2 + 1;
    return str.SubString(p1, p2-p1);
}

/* Reading Integer value (unquoted)
*/
Variant TJsonObject::readParamValueInteger(const String& str, int& pos)
{
    int p1 = pos;
    int p2 = 0;

    for (int i = p1; ;i++)
    {
        if (str[i] == ' ' || str[i] == ',' || str[i] == '}')
        {
            p2 = i;
            break;
        }
    }
    pos = p2 + 1;

    return str.SubString(p1, p2-p1);
}

/* Задает текст JSON */
String TJsonObject::setText(const String& text)
{
    _text = text;
    _resultCode = RT_UNDEFINED;
    _textLength = text.Length();
}

/* Разбор JSON*/
void TJsonObject::parse()
{
    _pos = 1;
    TJsonDataType::Type type = TJsonDataType::PT_UNDEFINED;

    _resultCode = RT_CONTINUOUS;

    while ( _resultCode != RT_END_TEXT )
    {
        String sNode = readNode(_text, _pos);
        if (_resultCode != RT_CONTINUOUS )
        {
            break;
        }

        while ( (_resultCode != RT_END_PARAMETERS )&& _pos >= 0 )
        {
            String sParamName = readParamName(_text, _pos);
            if ( _resultCode != RT_END_PARAMETERS )
            {
                Variant sParamValue = readParamValue(_text, _pos, type);

                //Variant VT_INT
                _result.SubNode[sNode].Param[sParamName] = sParamValue;
                //_result[sNode].Value = sParamValue;
                //_result[sNode][sParamName] = sParamValue;
            }
            else
            {
                break;
            }
        }
    }
}

TJsonNode* TJsonObject::getRootNode()
{
    return &_result;
}




TJsonNode* TJsonNode::getSubNode(const String& nodeName)
{
    // Здесь сделать проверку на наличие
    if ( SubNode.find(nodeName) != SubNode.end() )
    {
        return &SubNode[nodeName];
    }
    else
    {
        return NULL;
    }
}

JsonParamType TJsonNode::getParam(const String& paramName, JsonParamType defaultValue)
{
    //
    if ( Param.find(paramName) != Param.end() )
    {
        return Param[paramName];
    }
    else
    {
        return defaultValue;
    }
}



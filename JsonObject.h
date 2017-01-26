/*
   @purpose: Содержит классы для работы с текстом в формате JSON
   @author: vsovchinnikov
   @created: 2017-01-25

   @note:
   1. Не реализован разбор subNod-ов.


   Пример использования:

    String str = "\"img\":{\"name\"=\"visa\",\"zorder\"=5, \"width\"=\"100\", \"height\" = \"50\"}";


    TJsonObject json(str);
    json.parse();

    TJsonNode* rootNode = json.getRootNode();
    TJsonNode* subNode = rootNode->getSubNode("img");

    Variant s = subNode->getParam("zorderh", 57);


    Принцип работы:
    При разборе текста учитывайте что после попытки считывания очередного, следующего за
    последним узла, параметра или значения
    позиция _pos будет указывать на символ - признак завершения списка узлов или параметров.
    Например при вызове считывании последного значения функцией readParam
    и очередного вызова этой же функции _pos на символ, следующей за символом
    фигурной скобки ('}'), означающего конец узла. При последующем считывании
    узла, поиск начнется с этого символа.

*/

#ifndef JsonObjectH
#define JsonObjectH

#include <Classes.hpp>
#include <map.h>

class TJsonDataType
{
public:
    typedef enum Type
    {
        PT_UNDEFINED,
        PT_STRING,
        PT_INTEGER,
        PT_DOUBLE,
        PT_NODE
    } _Type;

};

/**/
typedef Variant JsonParamType;

class TJsonNode
{
public:
    std::map<String, TJsonNode> SubNode;
    std::map<String, JsonParamType> Param;
    //std::map<String, String> Param;

public:
    TJsonNode* getSubNode(const String& nodeName);
    JsonParamType getParam(const String& nodeName, JsonParamType defaultValue);

};

/* TJsonObject
   Класс для парсинга JSON и хранения данных
*/
class TJsonObject
{

public:

    typedef enum _TResultType
    {
        RT_UNDEFINED,
        RT_CONTINUOUS,  // Успешное продолжение без замечаний
        RT_END_NODE,        // Закончился список узлов
        RT_END_PARAMETERS,  // Закончился список параметров
        RT_END_TEXT,
        RT_PARSE_ERROR
    } TResultType;

    __fastcall TJsonObject();
    __fastcall TJsonObject(const String& text);

    /* Parsing */
    String readNode(const String& str, int& pos);
    String readParamName(const String& str, int& pos);
    Variant readParamValue(const String& str, int& p1, TJsonDataType::Type& type);
    Variant readParamValueString(const String& str, int& pos);
    Variant readParamValueInteger(const String& str, int& pos);

    String setText(const String& text);
    void parse();

    TJsonNode* getRootNode();

private:
    int _pos;
    TResultType _resultCode;
    String _text;
    int _textLength;


    //typedef std::map<String, String> JsonParam;
    //typedef std::map<String, JsonParam> JsonNode;

    TJsonNode _result;
    //std::map<String, TJsonParam> _result;

public:
    NodeByName(const String& nodeName);
    ParamByName(const String& nodeName);
    isNodeExist(const String& nodeName);
    isParamExist(const String& nodeName);

};



//---------------------------------------------------------------------------
#endif

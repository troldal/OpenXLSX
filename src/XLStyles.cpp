//
// Created by Troldal on 02/09/16.
//

#include "XLStyles.h"
#include "XLDocument.h"

using namespace std;
using namespace OpenXLSX;

std::map<std::string, XLStyles> XLStyles::s_styles = {};

XMLNode *XLStyles::m_numberFormatsNode = nullptr;

XMLNode *XLStyles::m_fontsNode = nullptr;

XMLNode *XLStyles::m_fillsNode = nullptr;

XMLNode *XLStyles::m_bordersNode = nullptr;

XMLNode *XLStyles::m_cellFormatNode = nullptr;

XMLNode *XLStyles::m_cellStyleNode = nullptr;

XMLNode *XLStyles::m_colors = nullptr;


/**
 * @details
 */
XLStyles::XLStyles(XLDocument &parent,
                   const std::string &filePath)
    : XLAbstractXMLFile(parent, filePath),
      XLSpreadsheetElement(parent)
{
    ParseXMLData();
}

/**
 * @details
 */
XLStyles::~XLStyles()
{
    //CommitXMLData();
}

/**
 * @details
 */
bool XLStyles::ParseXMLData()
{
    m_numberFormatsNode = XmlDocument()->RootNode()->ChildNode("numFmts");
    m_fontsNode = XmlDocument()->RootNode()->ChildNode("fonts");
    m_fillsNode = XmlDocument()->RootNode()->ChildNode("fills");
    m_bordersNode = XmlDocument()->RootNode()->ChildNode("borders");
    m_cellFormatNode = XmlDocument()->RootNode()->ChildNode("cellXfs");
    m_cellStyleNode = XmlDocument()->RootNode()->ChildNode("cellStyles");

    // Read fonts
    auto currentFont = m_fontsNode->ChildNode();
    while (currentFont != nullptr) {
        std::string name = currentFont->ChildNode("name")->Attribute("val")->Value();
        unsigned int size = stoi(currentFont->ChildNode("sz")->Attribute("val")->Value());

        std::string color = "FF000000";
        auto colorAttr = currentFont->ChildNode("color")->Attribute("rgb");
        if (colorAttr != nullptr) color = colorAttr->Value();

        bool bold = false;
        if (currentFont->ChildNode("b") != nullptr) bold = true;

        bool italics = false;
        if (currentFont->ChildNode("i") != nullptr) italics = true;

        bool underline = false;
        if (currentFont->ChildNode("u") != nullptr) underline = true;

        m_fonts.push_back(make_unique<XLFont>(name, size, XLColor(color), bold, italics, underline));

        currentFont = currentFont->NextSibling();
    }


    return true;
}

/**
 * @details
 */
void XLStyles::AddFont(const XLFont &font)
{
    m_fonts.push_back(make_unique<XLFont>(font.m_name,
                                          font.m_size,
                                          font.m_color,
                                          font.m_bold,
                                          font.m_italics,
                                          font.m_underline));

    auto newFont = XmlDocument()->CreateNode("font");

    if (font.m_bold) {
        auto boldNode = XmlDocument()->CreateNode("b");
        newFont->AppendNode(boldNode);
    }

    if (font.m_italics) {
        auto italicsNode = XmlDocument()->CreateNode("i");
        newFont->AppendNode(italicsNode);
    }

    if (font.m_underline) {
        auto underlineNode = XmlDocument()->CreateNode("u");
        newFont->AppendNode(underlineNode);
    }

    auto sizeNode = XmlDocument()->CreateNode("sz");
    auto sizeAttr = XmlDocument()->CreateAttribute("val", to_string(font.m_size));
    sizeNode->AppendAttribute(sizeAttr);
    newFont->AppendNode(sizeNode);

    auto colorNode = XmlDocument()->CreateNode("color");
    auto colorAttr = XmlDocument()->CreateAttribute("rgb", font.m_color.Hex());
    colorNode->AppendAttribute(colorAttr);
    newFont->AppendNode(colorNode);

    auto nameNode = XmlDocument()->CreateNode("name");
    auto nameAttr = XmlDocument()->CreateAttribute("val", font.m_name);
    nameNode->AppendAttribute(nameAttr);
    newFont->AppendNode(nameNode);

    auto familyNode = XmlDocument()->CreateNode("family");
    auto familyAttr = XmlDocument()->CreateAttribute("val", "2");
    familyNode->AppendAttribute(familyAttr);
    newFont->AppendNode(familyNode);

    m_fontsNode->AppendNode(newFont);
}

/**
 * @details
 */
XLFont &XLStyles::Font(unsigned int id)
{
    return *m_fonts.at(id);
}

/**
 * @details
 */
unsigned int XLStyles::FontId(const std::string &font)
{

    unsigned int index = 0;
    unsigned int result = 0;
    for (auto &iter : m_fonts) {
        if (iter->UniqueId() == font) {
            result = index;
            break;
        }

        index++;
    }

    return result + 1;
}


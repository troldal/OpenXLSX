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
    : XLAbstractXMLFile(parent.RootDirectory()->string(), "xl/styles.xml"),
      XLSpreadsheetElement(parent)
{
    LoadXMLData();
}

/**
 * @details
 */
XLStyles::~XLStyles()
{
    SaveXMLData();
}

/**
 * @details
 */
bool XLStyles::ParseXMLData()
{
    m_numberFormatsNode = XmlDocument().rootNode()->childNode("numFmts");
    m_fontsNode = XmlDocument().rootNode()->childNode("fonts");
    m_fillsNode = XmlDocument().rootNode()->childNode("fills");
    m_bordersNode = XmlDocument().rootNode()->childNode("borders");
    m_cellFormatNode = XmlDocument().rootNode()->childNode("cellXfs");
    m_cellStyleNode = XmlDocument().rootNode()->childNode("cellStyles");

    // Read fonts
    auto currentFont = m_fontsNode->childNode();
    while (currentFont != nullptr) {
        std::string name = currentFont->childNode("name")->attribute("val")->value();
        unsigned int size = stoi(currentFont->childNode("sz")->attribute("val")->value());

        std::string color = "FF000000";
        auto colorAttr = currentFont->childNode("color")->attribute("rgb");
        if (colorAttr != nullptr) color = colorAttr->value();

        bool bold = false;
        if (currentFont->childNode("b") != nullptr) bold = true;

        bool italics = false;
        if (currentFont->childNode("i") != nullptr) italics = true;

        bool underline = false;
        if (currentFont->childNode("u") != nullptr) underline = true;

        m_fonts.push_back(make_unique<XLFont>(name, size, XLColor(color), bold, italics, underline));

        currentFont = currentFont->nextSibling();
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

    auto newFont = XmlDocument().createNode("font");

    if (font.m_bold) {
        auto boldNode = XmlDocument().createNode("b");
        newFont->appendNode(boldNode);
    }

    if (font.m_italics) {
        auto italicsNode = XmlDocument().createNode("i");
        newFont->appendNode(italicsNode);
    }

    if (font.m_underline) {
        auto underlineNode = XmlDocument().createNode("u");
        newFont->appendNode(underlineNode);
    }

    auto sizeNode = XmlDocument().createNode("sz");
    auto sizeAttr = XmlDocument().createAttribute("val", to_string(font.m_size));
    sizeNode->appendAttribute(sizeAttr);
    newFont->appendNode(sizeNode);

    auto colorNode = XmlDocument().createNode("color");
    auto colorAttr = XmlDocument().createAttribute("rgb", font.m_color.Hex());
    colorNode->appendAttribute(colorAttr);
    newFont->appendNode(colorNode);

    auto nameNode = XmlDocument().createNode("name");
    auto nameAttr = XmlDocument().createAttribute("val", font.m_name);
    nameNode->appendAttribute(nameAttr);
    newFont->appendNode(nameNode);

    auto familyNode = XmlDocument().createNode("family");
    auto familyAttr = XmlDocument().createAttribute("val", "2");
    familyNode->appendAttribute(familyAttr);
    newFont->appendNode(familyNode);

    m_fontsNode->appendNode(newFont);
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


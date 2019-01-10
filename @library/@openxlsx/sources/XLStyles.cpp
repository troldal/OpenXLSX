//
// Created by Troldal on 02/09/16.
//

#include "XLStyles.h"
#include "XLDocument.h"

#include <pugixml.hpp>

using namespace std;
using namespace OpenXLSX::Impl;

std::map<std::string, XLStyles> XLStyles::s_styles = {};

XMLNode XLStyles::s_numberFormatsNode = XMLNode();
XMLNode XLStyles::s_fontsNode = XMLNode();
XMLNode XLStyles::s_fillsNode = XMLNode();
XMLNode XLStyles::s_bordersNode = XMLNode();
XMLNode XLStyles::s_cellFormatNode = XMLNode();
XMLNode XLStyles::s_cellStyleNode = XMLNode();
XMLNode XLStyles::s_colors = XMLNode();

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
    s_numberFormatsNode = XmlDocument()->child("numFmts");
    s_fontsNode = XmlDocument()->child("fonts");
    s_fillsNode = XmlDocument()->child("fills");
    s_bordersNode = XmlDocument()->child("borders");
    s_cellFormatNode = XmlDocument()->child("cellXfs");
    s_cellStyleNode = XmlDocument()->child("cellStyles");

    // Read fonts
    auto currentFont = s_fontsNode.first_child();
    while (currentFont != nullptr) {
        std::string name = currentFont.child("name").attribute("val").value();
        unsigned int size = stoi(currentFont.child("sz").attribute("val").value());

        std::string color = "FF000000";
        auto colorAttr = currentFont.child("color").attribute("rgb");
        if (colorAttr != nullptr) color = colorAttr.value();

        bool bold = false;
        if (currentFont.child("b").type() != pugi::node_null) bold = true;

        bool italics = false;
        if (currentFont.child("i").type() != pugi::node_null) italics = true;

        bool underline = false;
        if (currentFont.child("u").type() != pugi::node_null) underline = true;

        m_fonts.push_back(make_unique<XLFont>(name, size, XLColor(color), bold, italics, underline));

        currentFont = currentFont.next_sibling();
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

    auto newFont = s_fontsNode.append_child("font");

    if (font.m_bold) newFont.append_child("b");
    if (font.m_italics) newFont.append_child("i");
    if (font.m_underline) newFont.append_child("u");

    newFont.append_child("sz").append_attribute("val").set_value(to_string(font.m_size).c_str());
    newFont.append_child("color").append_attribute("rgb").set_value(font.m_color.Hex().c_str());
    newFont.append_child("name").append_attribute("val").set_value(font.m_name.c_str());
    newFont.append_child("family").append_attribute("val").set_value("2");

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


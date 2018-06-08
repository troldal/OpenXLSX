//
// Created by Kenneth Balslev on 04/06/2017.
//

#include "../headers/XLFont.h"

#include <sstream>

using namespace std;
using namespace OpenXLSX;

std::map<string, XLFont> XLFont::s_fonts = {};

/**
 * @details
 */
XLFont::XLFont(const string &name,
               unsigned int size,
               const XLColor &color,
               bool bold,
               bool italics,
               bool underline)
    : m_fontNode(XMLNode()),
      m_name(name),
      m_size(size),
      m_color(color),
      m_bold(bold),
      m_italics(italics),
      m_underline(underline),
      m_theme(""),
      m_family(""),
      m_scheme("")
{

}

/**
 * @details
 */
std::string OpenXLSX::XLFont::UniqueId() const
{
    stringstream str;

    str << m_name << m_size << m_color.Hex();
    if (m_bold) str << "b";
    if (m_italics) str << "i";
    if (m_underline) str << "u";

    return str.str();
}

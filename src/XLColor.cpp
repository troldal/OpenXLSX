//
// Created by Kenneth Balslev on 04/06/2017.
//

#include "XLColor.h"
#include <sstream>

using namespace std;
using namespace OpenXLSX;

/**
 * @details
 */
XLColor::XLColor(unsigned int red,
                 unsigned int green,
                 unsigned int blue)
    : m_red(std::min(red, 255U)),
      m_green(std::min(green, 255U)),
      m_blue(std::min(blue, 255U))
{

}

/**
 * @details
 */
XLColor::XLColor(const std::string &hexCode)
    : m_red(0),
      m_green(0),
      m_blue(0)
{
    SetColor(hexCode);
}

/**
 * @details
 */
void XLColor::SetColor(unsigned int red,
                       unsigned int green,
                       unsigned int blue)
{
    m_red = std::min(red, 255U);
    m_green = std::min(green, 255U);
    m_blue = std::min(blue, 255U);
}

/**
 * @details
 */
void XLColor::SetColor(const std::string &hexCode)
{

    std::string temp = hexCode;

    if (temp.size() > 6) {
        temp = temp.substr(temp.size() - 6);
    }

    std::string red = temp.substr(0, 2);
    std::string green = temp.substr(2, 2);
    std::string blue = temp.substr(4, 2);

    m_red = stoul(red, 0, 16);
    m_green = stoul(green, 0, 16);
    m_blue = stoul(blue, 0, 16);
}

/**
 * @details
 */
unsigned int XLColor::Red() const
{
    return m_red;
}

/**
 * @details
 */
unsigned int XLColor::Green() const
{
    return m_green;
}

/**
 * @details
 */
unsigned int XLColor::Blue() const
{
    return m_blue;
}

/**
 * @details
 */
std::string XLColor::Hex() const
{

    stringstream str;
    str << "ff";

    if (m_red < 16) str << "0";
    str << std::hex << m_red;

    if (m_green < 16) str << "0";
    str << std::hex << m_green;

    if (m_blue < 16) str << "0";
    str << std::hex << m_blue;

    return (str.str());
}

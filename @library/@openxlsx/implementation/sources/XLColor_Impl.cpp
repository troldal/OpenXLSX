//
// Created by Kenneth Balslev on 04/06/2017.
//

#include "XLColor_Impl.h"
#include <sstream>
#include <algorithm>

using namespace std;
using namespace OpenXLSX;

/**
 * @details
 */
Impl::XLColor::XLColor(unsigned int red, unsigned int green, unsigned int blue)
        : m_red(min(red, 255U)),
          m_green(min(green, 255U)),
          m_blue(min(blue, 255U)) {

}

/**
 * @details
 */
Impl::XLColor::XLColor(const std::string& hexCode)
        : m_red(0),
          m_green(0),
          m_blue(0) {

    SetColor(hexCode);
}

/**
 * @details
 */
void Impl::XLColor::SetColor(unsigned int red, unsigned int green, unsigned int blue) {

    m_red   = min(red, 255U);
    m_green = min(green, 255U);
    m_blue  = min(blue, 255U);
}

/**
 * @details
 */
void Impl::XLColor::SetColor(const std::string& hexCode) {

    std::string temp = hexCode;

    if (temp.size() > 6) {
        temp = temp.substr(temp.size() - 6);
    }

    std::string red   = temp.substr(0, 2);
    std::string green = temp.substr(2, 2);
    std::string blue  = temp.substr(4, 2);

    m_red   = static_cast<unsigned int>(stoul(red, nullptr, 16));
    m_green = static_cast<unsigned int>(stoul(green, nullptr, 16));
    m_blue  = static_cast<unsigned int>(stoul(blue, nullptr, 16));
}

/**
 * @details
 */
unsigned int Impl::XLColor::Red() const {

    return m_red;
}

/**
 * @details
 */
unsigned int Impl::XLColor::Green() const {

    return m_green;
}

/**
 * @details
 */
unsigned int Impl::XLColor::Blue() const {

    return m_blue;
}

/**
 * @details
 */
std::string Impl::XLColor::Hex() const {

    stringstream str;
    str << "ff";

    if (m_red < 16)
        str << "0";
    str << std::hex << m_red;

    if (m_green < 16)
        str << "0";
    str << std::hex << m_green;

    if (m_blue < 16)
        str << "0";
    str << std::hex << m_blue;

    return (str.str());
}

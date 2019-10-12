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
Impl::XLColor::XLColor(uint8_t red, uint8_t green, uint8_t blue)
        : m_red(red),
          m_green(green),
          m_blue(blue) {

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
void Impl::XLColor::SetColor(uint8_t red, uint8_t green, uint8_t blue) {

    m_red = red;
    m_green = green;
    m_blue = blue;
}

/**
 * @details
 */
void Impl::XLColor::SetColor(const std::string& hexCode) {

    std::string temp = hexCode;

    if (temp.size() > 6) {
        temp = temp.substr(temp.size() - 6);
    }

    std::string red = temp.substr(0, 2);
    std::string green = temp.substr(2, 2);
    std::string blue = temp.substr(4, 2);

    m_red = static_cast<uint8_t>(stoul(red, nullptr, 16));
    m_green = static_cast<uint8_t>(stoul(green, nullptr, 16));
    m_blue = static_cast<uint8_t>(stoul(blue, nullptr, 16));
}

/**
 * @details
 */
uint8_t Impl::XLColor::Red() const {

    return m_red;
}

/**
 * @details
 */
uint8_t Impl::XLColor::Green() const {

    return m_green;
}

/**
 * @details
 */
uint8_t Impl::XLColor::Blue() const {

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

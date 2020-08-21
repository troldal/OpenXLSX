/*

   ____                               ____      ___ ____       ____  ____      ___
  6MMMMb                              `MM(      )M' `MM'      6MMMMb\`MM(      )M'
 8P    Y8                              `MM.     d'   MM      6M'    ` `MM.     d'
6M      Mb __ ____     ____  ___  __    `MM.   d'    MM      MM        `MM.   d'
MM      MM `M6MMMMb   6MMMMb `MM 6MMb    `MM. d'     MM      YM.        `MM. d'
MM      MM  MM'  `Mb 6M'  `Mb MMM9 `Mb    `MMd       MM       YMMMMb     `MMd
MM      MM  MM    MM MM    MM MM'   MM     dMM.      MM           `Mb     dMM.
MM      MM  MM    MM MMMMMMMM MM    MM    d'`MM.     MM            MM    d'`MM.
YM      M9  MM    MM MM       MM    MM   d'  `MM.    MM            MM   d'  `MM.
 8b    d8   MM.  ,M9 YM    d9 MM    MM  d'    `MM.   MM    / L    ,M9  d'    `MM.
  YMMMM9    MMYMMM9   YMMMM9 _MM_  _MM_M(_    _)MM_ _MMMMMMM MYMMMM9 _M(_    _)MM_
            MM
            MM
           _MM_

  Copyright (c) 2018, Kenneth Troldal Balslev

  All rights reserved.

  Redistribution and use in source and binary forms, with or without
  modification, are permitted provided that the following conditions are met:
  - Redistributions of source code must retain the above copyright
    notice, this list of conditions and the following disclaimer.
  - Redistributions in binary form must reproduce the above copyright
    notice, this list of conditions and the following disclaimer in the
    documentation and/or other materials provided with the distribution.
  - Neither the name of the author nor the
    names of any contributors may be used to endorse or promote products
    derived from this software without specific prior written permission.

  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
  ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
  WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
  DISCLAIMED. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY
  DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
  (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
  LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
  ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
  (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
  SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

 */

#ifndef OPENXLSX_XLCELL_INL
#define OPENXLSX_XLCELL_INL

/**
 * @details
 * @pre
 * @post
 */
template<typename T,
    typename std::enable_if<std::is_integral<T>::value, int64_t>::type*>
XLCell& XLCell::XLCellValueProxy::operator=(T numberValue)
{
    if constexpr (std::is_same<T, bool>::value) {    // if bool
        m_cell->setBoolean(numberValue);
    }
    else {    // if not bool
        m_cell->setInteger(numberValue);
    }

    return *m_cell;
}

/**
 * @details
 * @pre
 * @post
 */
template<typename T,
    typename std::enable_if<std::is_floating_point<T>::value, long double>::type*>
XLCell& XLCell::XLCellValueProxy::operator=(T numberValue)
{
    m_cell->setFloat(numberValue);
    return *m_cell;
}

/**
 * @details
 * @pre
 * @post
 */
template<typename T,
    typename std::enable_if<!std::is_same<T, XLCellValue>::value && !std::is_same<T, bool>::value &&
        std::is_constructible<T, const char*>::value, T>::type*>
XLCell& XLCell::XLCellValueProxy::operator=(T stringValue)
{
    if constexpr (std::is_same<const char*, typename std::remove_reference<typename std::remove_cv<T>::type>::type>::value)
        m_cell->setString(stringValue);
    else if constexpr (std::is_same<std::string_view, T>::value)
        m_cell->setString(std::string(stringValue).c_str());
    else
        m_cell->setString(stringValue.c_str());
    return *m_cell;
}

/**
 * @details
 * @pre
 * @post
 */
template<typename T, typename std::enable_if<std::is_same<T, XLCellValue>::value, T>::type*>
XLCell& XLCell::XLCellValueProxy::operator=(T value)
{
    switch (value.type()) {
        case XLValueType::Boolean:
            m_cell->setBoolean(value.template get<bool>());
            break;
        case XLValueType::Integer:
            m_cell->setInteger(value.template get<int64_t>());
            break;
        case XLValueType::Float:
            m_cell->setInteger(value.template get<double>());
            break;
        case XLValueType::String:
            m_cell->setString(value.template get<std::string>().c_str());
            break;
        case XLValueType::Empty:
            break;
        default:
            break;
    }

    return *m_cell;
}

/**
 * @details
 * @pre
 * @post
 */
template<typename T, typename std::enable_if<std::is_integral<T>::value, int64_t>::type*>
void XLCell::XLCellValueProxy::set(T numberValue)
{
    *this = numberValue;
}

/**
 * @details
 * @pre
 * @post
 */
template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type*>
void XLCell::XLCellValueProxy::set(T numberValue)
{
    *this = numberValue;
}

/**
 * @details
 * @pre
 * @post
 */
template<typename T,
    typename std::enable_if<!std::is_same<T, XLCellValue>::value && !std::is_same<T, bool>::value &&
        std::is_constructible<T, const char*>::value,
                            T>::type*>
void XLCell::XLCellValueProxy::set(T stringValue)
{
    *this = stringValue;
}

/**
 * @details
 * @pre
 * @post
 */
template<typename T>
T XLCell::XLCellValueProxy::get() const {
    return m_cell->getValue().get<T>();
}

/**
 * @brief
 * @param lhs
 * @param rhs
 * @return
 */
inline bool operator==(const XLCell& lhs, const XLCell& rhs)
{
    return lhs.m_cellNode == rhs.m_cellNode;
}


#endif    // OPENXLSX_XLCELL_INL

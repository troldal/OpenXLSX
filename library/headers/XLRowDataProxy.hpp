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

#ifndef OPENXLSX_XLROWDATAPROXY_HPP
#define OPENXLSX_XLROWDATAPROXY_HPP

#pragma warning(push)
#pragma warning(disable : 4251)
#pragma warning(disable : 4275)

// ===== External Includes ===== //
#include <iterator>
#include <type_traits>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLCellValue.hpp"
#include "XLRowData.hpp"
#include "XLXmlParser.hpp"

// ========== CLASS AND ENUM TYPE DEFINITIONS ========== //
namespace OpenXLSX
{
    class XLRow;

    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLRowDataProxy
    {
        friend class XLRow;

    public:
        /**
         * @brief
         */
        ~XLRowDataProxy();

        /**
         * @brief
         * @param other
         * @return
         */
        XLRowDataProxy& operator=(const XLRowDataProxy& other);

        /**
         * @brief
         * @param values
         * @return
         */
        XLRowDataProxy& operator=(const std::vector<XLCellValue>& values);

        template<typename T,
                 typename std::enable_if<!std::is_same_v<T, XLRowDataProxy> &&
                                             std::is_base_of_v<typename std::forward_iterator_tag,
                                                               typename std::iterator_traits<typename T::iterator>::iterator_category>,
                                         T>::type* = nullptr>
        XLRowDataProxy& operator=(const T& values);

        /**
         * @brief
         * @return
         */
        operator std::vector<XLCellValue>();    // NOLINT

    private:
        //---------- Private Member Functions ---------- //

        /**
         * @brief
         * @param row
         * @param rowNode
         */
        XLRowDataProxy(XLRow* row, XMLNode* rowNode);

        /**
         * @brief
         * @param other
         */
        XLRowDataProxy(const XLRowDataProxy& other);

        /**
         * @brief
         * @param other
         */
        XLRowDataProxy(XLRowDataProxy&& other) noexcept;

        /**
         * @brief
         * @param other
         * @return
         */
        XLRowDataProxy& operator=(XLRowDataProxy&& other) noexcept;

        /**
         * @brief
         * @return
         */
        std::vector<XLCellValue> getValues() const;

        //---------- Private Member Variables ---------- //

        XLRow*   m_row;     /**< */
        XMLNode* m_rowNode; /**< */
    };

}    // namespace OpenXLSX

namespace OpenXLSX
{
    template<typename T,
             typename std::enable_if<!std::is_same_v<T, XLRowDataProxy> &&
                                         std::is_base_of_v<typename std::forward_iterator_tag,
                                                           typename std::iterator_traits<typename T::iterator>::iterator_category>,
                                     T>::type*>
    XLRowDataProxy& XLRowDataProxy::operator=(const T& values)
    {
        auto range = XLRowDataRange(*m_rowNode, 1, 16384, nullptr);

        auto dst = range.begin();
        auto src = values.begin();

        while (true) {
            dst->value() = *src;
            ++src;
            if (src == values.end()) break;
            ++dst;
        }

        return *this;
    }
}    // namespace OpenXLSX

#pragma warning(pop)
#endif    // OPENXLSX_XLROWDATAPROXY_HPP

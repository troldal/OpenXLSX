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

#ifndef OPENXLSX_XLCOLUMN_HPP
#define OPENXLSX_XLCOLUMN_HPP

#pragma GCC diagnostic push
#pragma GCC diagnostic ignored "-Wunknown-pragmas" // disable warning about below #pragma warning being unknown
#ifdef _MSC_VER                                    // additional condition because the previous line does not work on gcc 12.2
#   pragma warning(push)
#   pragma warning(disable : 4251)
#   pragma warning(disable : 4275)
#endif // _MSC_VER
#pragma GCC diagnostic pop

// ===== External Includes ===== //
#include <memory>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLStyles.hpp"          // XLStyleIndex
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLColumn
    {
    public:
        /**
         * @brief Constructor
         * @param columnNode A pointer to the XMLNode for the column.
         */
        explicit XLColumn(const XMLNode& columnNode);

        /**
         * @brief Copy Constructor [deleted]
         */
        XLColumn(const XLColumn& other);

        /**
         * @brief Move Constructor
         * @note The move constructor has been explicitly deleted.
         */
        XLColumn(XLColumn&& other) noexcept;

        /**
         * @brief Destructor
         */
        ~XLColumn();

        /**
         * @brief Copy assignment operator [deleted]
         */
        XLColumn& operator=(const XLColumn& other);

        /**
         * @brief
         * @param other
         * @return
         */
        XLColumn& operator=(XLColumn&& other) noexcept = default;

        /**
         * @brief Get the width of the column.
         * @return The width of the column.
         */
        float width() const;

        /**
         * @brief Set the width of the column
         * @param width The width of the column
         */
        void setWidth(float width);

        /**
         * @brief Is the column hidden?
         * @return The state of the column.
         */
        bool isHidden() const;

        /**
         * @brief Set the column to be shown or hidden.
         * @param state The state of the column.
         */
        void setHidden(bool state);

        /**
         * @brief Get the XMLNode object for the column.
         * @return The XMLNode for the column
         */
        XMLNode& columnNode() const;

        /**
         * @brief Get the array index of xl/styles.xml:<styleSheet>:<cellXfs> for the style assigned to the column.
         *        This value is stored in the col attributes like so: style="2"
         * @returns The index of the applicable format style
         */
        XLStyleIndex format() const;

        /**
         * @brief Set the column style as a reference to the array index of xl/styles.xml:<styleSheet>:<cellXfs>
         * @param cellFormatIndex The style to set, corresponding to the index of XLStyles::cellStyles()
         * @returns true on success, false on failure
         */
        bool setFormat(XLStyleIndex cellFormatIndex);

    private:
        std::unique_ptr<XMLNode> m_columnNode; /**< A pointer to the XMLNode object for the column. */
    };

}    // namespace OpenXLSX

#pragma GCC diagnostic push
#pragma GCC diagnostic ignored "-Wunknown-pragmas" // disable warning about below #pragma warning being unknown
#ifdef _MSC_VER                                    // additional condition because the previous line does not work on gcc 12.2
#   pragma warning(pop)
#endif // _MSC_VER
#pragma GCC diagnostic pop

#endif    // OPENXLSX_XLCOLUMN_HPP

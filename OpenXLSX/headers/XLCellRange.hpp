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

#ifndef OPENXLSX_XLCELLRANGE_HPP
#define OPENXLSX_XLCELLRANGE_HPP

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
#include "XLCell.hpp"
#include "XLCellIterator.hpp"
#include "XLCellReference.hpp"
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    /**
     * @brief This class encapsulates the concept of a cell range, i.e. a square area
     * (or subset) of cells in a spreadsheet.
     */
    class OPENXLSX_EXPORT XLCellRange
    {
        friend class XLCellIterator;

        //----------------------------------------------------------------------------------------------------------------------
        //           Public Member Functions
        //----------------------------------------------------------------------------------------------------------------------

    public:
        /**
         * @brief Default constructor for variable declaration
         */
        XLCellRange();

        /**
         * @brief
         * @param dataNode
         * @param topLeft
         * @param bottomRight
         * @param sharedStrings
         */
        explicit XLCellRange(const XMLNode&         dataNode,
                             const XLCellReference& topLeft,
                             const XLCellReference& bottomRight,
                             const XLSharedStrings& sharedStrings);

        /**
         * @brief Copy constructor
         * @param other The range object to be copied.
         * @note This implements the default copy constructor, i.e. memberwise copying.
         */
        XLCellRange(const XLCellRange& other);

        /**
         * @brief Move constructor
         * @param other The range object to be moved.
         * @note This implements the default move constructor, i.e. memberwise move.
         */
        XLCellRange(XLCellRange&& other) noexcept;

        /**
         * @brief Destructor [default]
         * @note This implements the default destructor.
         */
        ~XLCellRange();

        /**
         * @brief The copy assignment operator [default]
         * @param other The range object to be copied and assigned.
         * @return A reference to the new object.
         * @throws A std::range_error if the source range and destination range are of different size and shape.
         * @note This implements the default copy assignment operator.
         */
        XLCellRange& operator=(const XLCellRange& other);

        /**
         * @brief The move assignment operator [default].
         * @param other The range object to be moved and assigned.
         * @return A reference to the new object.
         * @note This implements the default move assignment operator.
         */
        XLCellRange& operator=(XLCellRange&& other) noexcept;

        /**
         * @brief populate the m_columnStyles
         * @return a const XLCellReference
         */
        void fetchColumnStyles();

        /**
         * @brief get the top left cell
         * @return a const XLCellReference
         */
        const XLCellReference topLeft() const;

        /**
         * @brief get the bottom right cell
         * @return a const XLCellReference
         */
        const XLCellReference bottomRight() const;

        /**
         * @brief get the string reference that corresponds to the represented cell range
         * @return a std::string range reference, e.g. "A2:Z5"
         */
        std::string address() const;

        /**
         * @brief Get the number of rows in the range.
         * @return The number of rows.
         */
        uint32_t numRows() const;

        /**
         * @brief Get the number of columns in the range.
         * @return The number of columns.
         */
        uint16_t numColumns() const;

        /**
         * @brief
         * @return
         */
        XLCellIterator begin() const;

        /**
         * @brief
         * @return
         */
        XLCellIterator end() const;

        /**
         * @brief
         */
        void clear();

        /**
         * @brief Templated assignment operator - assign value to a range of cells
         * @tparam T The type of the value argument.
         * @param value The value.
         * @return A reference to the assigned-to object.
         */
        template<typename T,
                 typename = std::enable_if_t<
                     std::is_integral_v<T> || std::is_floating_point_v<T> || std::is_same_v<std::decay_t<T>, std::string> ||
                     std::is_same_v<std::decay_t<T>, std::string_view> || std::is_same_v<std::decay_t<T>, const char*> ||
                     std::is_same_v<std::decay_t<T>, char*> || std::is_same_v<T, XLDateTime>>>
        XLCellRange& operator=(T value)
        {
            // forward implementation to templated XLCellValue& XLCellValue::operator=(T value)
            for (auto it = begin(); it != end(); ++it) it->value() = value;
            return *this;
        }

        /**
         * @brief Set cell format for a range of cells
         * @param cellFormatIndex The style to set, corresponding to the nidex of XLStyles::cellStyles()
         * @returns true on success, false on failure
         */
        bool setFormat(XLStyleIndex cellFormatIndex);

        //----------------------------------------------------------------------------------------------------------------------
        //           Private Member Variables
        //----------------------------------------------------------------------------------------------------------------------

    private:
        std::unique_ptr<XMLNode>  m_dataNode;      /**< */
        XLCellReference           m_topLeft;       /**< The cell reference of the first cell in the range */
        XLCellReference           m_bottomRight;   /**< The cell reference of the last cell in the range */
        XLSharedStrings           m_sharedStrings;
        std::vector<XLStyleIndex> m_columnStyles;  /**< quick access to column styles in the range - populated by fetchColumnStyles() */
    };
}    // namespace OpenXLSX

#pragma GCC diagnostic push
#pragma GCC diagnostic ignored "-Wunknown-pragmas" // disable warning about below #pragma warning being unknown
#ifdef _MSC_VER                                    // additional condition because the previous line does not work on gcc 12.2
#   pragma warning(pop)
#endif // _MSC_VER
#pragma GCC diagnostic pop

#endif    // OPENXLSX_XLCELLRANGE_HPP

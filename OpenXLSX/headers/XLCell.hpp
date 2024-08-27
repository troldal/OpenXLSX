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

#ifndef OPENXLSX_XLCELL_HPP
#define OPENXLSX_XLCELL_HPP

#pragma GCC diagnostic push
#pragma GCC diagnostic ignored "-Wunknown-pragmas" // disable warning about below #pragma warning being unknown
#ifdef _MSC_VER                                    // additional condition because the previous line does not work on gcc 12.2
#   pragma warning(push)
#   pragma warning(disable : 4251)
#   pragma warning(disable : 4275)
#endif // _MSC_VER
#pragma GCC diagnostic pop

#include <iostream> // std::ostream
#include <ostream>  // std::basic_ostream
#include <memory>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLCellReference.hpp"
#include "XLCellValue.hpp"
#include "XLFormula.hpp"
#include "XLSharedStrings.hpp"
#include "XLStyles.hpp"          // XLStyleIndex

// ========== CLASS AND ENUM TYPE DEFINITIONS ========== //
namespace OpenXLSX
{
    // ===== Flags that can be passed to XLCell::clear as parameter keep, flags can be combined with bitwise OR
    //                                  // Do not clear the cell's:
    constexpr const uint32_t XLKeepCellStyle   =  1; // style (attribute s)
    constexpr const uint32_t XLKeepCellType    =  2; // type (attribute t)
    constexpr const uint32_t XLKeepCellValue   =  4; // value (child node v)
    constexpr const uint32_t XLKeepCellFormula =  8; // formula (child node f)

    class XLCellRange;
    class XLSharedStrings;

    /**
     * @brief An implementation class encapsulating the properties and behaviours of a spreadsheet cell.
     */
    class OPENXLSX_EXPORT XLCell
    {
        friend class XLCellIterator;
        friend class XLCellValueProxy;
        friend class XLRowDataIterator;
        friend bool operator==(const XLCell& lhs, const XLCell& rhs);
        friend bool operator!=(const XLCell& lhs, const XLCell& rhs);

    public:
        //---------- Public Member Functions ----------//

        /**
         * @brief Default constructor. Constructs a null object.
         */
        XLCell();

        /**
         * @brief
         * @param cellNode
         * @param sharedStrings
         */
        XLCell(const XMLNode& cellNode, const XLSharedStrings& sharedStrings);

        /**
         * @brief Copy constructor
         * @param other The XLCell object to be copied.
         * @note The copy constructor has been deleted, as it makes no sense to copy a cell. If the objective is to
         * copy the getValue, create the the target object and then use the copy assignment operator.
         */
        XLCell(const XLCell& other);

        /**
         * @brief Move constructor
         * @param other The XLCell object to be moved
         * @note The move constructor has been deleted, as it makes no sense to move a cell.
         */
        XLCell(XLCell&& other) noexcept;

        /**
         * @brief Destructor
         * @note Using the default destructor
         */
        virtual ~XLCell();

        /**
         * @brief Copy assignment operator
         * @param other The XLCell object to be copy assigned
         * @return A reference to the new object
         * @note Copies only the cell contents, not the pointer to parent worksheet etc.
         */
        virtual XLCell& operator=(const XLCell& other);

        /**
         * @brief Move assignment operator
         * @param other The XLCell object to be move assigned
         * @return A reference to the new object
         * @note The move assignment constructor has been deleted, as it makes no sense to move a cell.
         */
        virtual XLCell& operator=(XLCell&& other) noexcept;

        /**
         * @brief Copy contents of a cell, value & formula
         * @param other The XLCell object from which to copy
         */
        void copyFrom(XLCell const& other);

        /**
         * @brief test if cell object has no (valid) content
         * @return
         */
        bool empty() const;

        /**
         * @brief opposite of empty()
         * @return
         */
        explicit operator bool() const;

        /**
         * @brief clear all cell content and attributes except for the cell reference (attribute r)
         * @param keep do not clear cell properties whose flags are set in keep (XLKeepCellStyle, XLKeepCellType,
         *              XLKeepCellValue, XLKeepCellFormula), flags can be combined with bitwise OR
         */
        void clear(uint32_t keep);

        /**
         * @brief
         * @return
         */
        XLCellValueProxy& value();

        /**
         * @brief
         * @return
         */
        const XLCellValueProxy& value() const;

        /**
         * @brief get the XLCellReference object for the cell.
         * @return A reference to the cells' XLCellReference object.
         */
        XLCellReference cellReference() const;

        /**
         * @brief get the XLCell object from the current cell offset
         * @return A reference to the XLCell object.
         */
        XLCell offset(uint16_t rowOffset, uint16_t colOffset) const;

        /**
         * @brief
         * @return
         */
        bool hasFormula() const;

        /**
         * @brief
         * @return
         */
        XLFormulaProxy& formula();

        /**
         * @brief
         * @return
         */
        const XLFormulaProxy& formula() const;

        /**
         * @brief
         */
        std::string getString() const { return value().getString(); }

        /**
         * @brief Get the array index of xl/styles.xml:<styleSheet>:<cellXfs> for the style used in this cell.
         *        This value is stored in the s attribute of a cell like so: s="2"
         * @returns The index of the applicable format style
         */
        XLStyleIndex cellFormat() const;

        /**
         * @brief Set the cell style (attribute s) with a reference to the array index of xl/styles.xml:<styleSheet>:<cellXfs>
         * @param cellFormatIndex The style to set, corresponding to the nidex of XLStyles::cellStyles()
         * @returns True on success, false on failure
         */
        bool setCellFormat(XLStyleIndex cellFormatIndex);

        /**
         * @brief Print the XML contents of the XLCell using the underlying XMLNode print function
         */
        void print(std::basic_ostream<char>& ostr) const;

    private:
        /**
         * @brief
         * @param lhs
         * @param rhs
         * @return
         */
        static bool isEqual(const XLCell& lhs, const XLCell& rhs);

        //---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_cellNode;      /**< A pointer to the root XMLNode for the cell. */
        XLSharedStrings          m_sharedStrings; /**< */
        XLCellValueProxy         m_valueProxy;    /**< */
        XLFormulaProxy           m_formulaProxy;  /**< */
    };

    class OPENXLSX_EXPORT XLCellAssignable : public XLCell
    {
    public:
        /**
         * @brief Default constructor. Constructs a null object.
         */
        XLCellAssignable() : XLCell() {}

        /**
         * @brief Copy constructor. Constructs an assignable XLCell from an existing cell
         * @param other the cell to construct from
         */
        XLCellAssignable (XLCell const & other);

        /**
         * @brief Move constructor. Constructs an assignable XLCell from a temporary (r)value
         * @param other the cell to construct from
         */
        XLCellAssignable (XLCell && other);

        // /**
        //  * @brief Inherit all constructors with parameters from XLCell
        //  */
        // template<class base>
        // // explicit XLCellAssignable(base b) : XLCell(b)
        // // NOTE: BUG: explicit keyword triggers tons of compiler errors when << operator attempts to use an XLCell (implicit conversion works because << is overloaded for XLCellAssignable)
        // XLCellAssignable(base b) : XLCell(b)
        // {}

        /**
         * @brief Copy assignment operator
         * @param other The XLCell object to be copy assigned
         * @return A reference to the new object
         * @note Copies only the cell contents, not the pointer to parent worksheet etc.
         */
        XLCellAssignable& operator=(const XLCell& other) override;
        XLCellAssignable& operator=(const XLCellAssignable& other);

        /**
         * @brief Move assignment operator -> overrides XLCell copy operator, becomes a copy operator
         * @param other The XLCell object to be copy assigned
         * @return A reference to the new object
         * @note Copies only the cell contents, not the pointer to parent worksheet etc.
         */
        XLCellAssignable& operator=(XLCell&& other) noexcept override;
        XLCellAssignable& operator=(XLCellAssignable&& other) noexcept;

        /**
         * @brief Templated assignment operator.
         * @tparam T The type of the value argument.
         * @param value The value.
         * @return A reference to the assigned-to object.
         */
        template<typename T,
                 typename = std::enable_if_t<
                     std::is_integral_v<T> || std::is_floating_point_v<T> || std::is_same_v<std::decay_t<T>, std::string> ||
                     std::is_same_v<std::decay_t<T>, std::string_view> || std::is_same_v<std::decay_t<T>, const char*> ||
                     std::is_same_v<std::decay_t<T>, char*> || std::is_same_v<T, XLDateTime>>>
        XLCellAssignable& operator=(T value)
        {
            XLCell::value() = value; // forward implementation to templated XLCellValue& XLCellValue::operator=(T value)
            return *this;
        }
    };
}    // namespace OpenXLSX

// ========== FRIEND FUNCTION IMPLEMENTATIONS ========== //
namespace OpenXLSX
{
    /**
     * @brief
     * @param lhs
     * @param rhs
     * @return
     */
    inline bool operator==(const XLCell& lhs, const XLCell& rhs) { return XLCell::isEqual(lhs, rhs); }

    /**
     * @brief
     * @param lhs
     * @param rhs
     * @return
     */
    inline bool operator!=(const XLCell& lhs, const XLCell& rhs) { return !XLCell::isEqual(lhs, rhs); }

    /**
     * @brief      ostream output of XLCell content as string
     * @param os   the ostream destination
     * @param c    the cell to output to the stream
     * @return
     */
    inline std::ostream& operator<<(std::ostream& os, const XLCell& c)
    {
        os << c.getString();
        // TODO: send to stream different data types based on cell data type
        return os;
    }

    /**
     * @brief      ostream output of XLCellAssignable content as string
     * @param os   the ostream destination
     * @param c    the cell to output to the stream
     * @return
     */
    inline std::ostream& operator<<(std::ostream& os, const XLCellAssignable& c)
    {
        os << c.getString();
        // TODO: send to stream different data types based on cell data type
        return os;
    }
}    // namespace OpenXLSX

#pragma GCC diagnostic push
#pragma GCC diagnostic ignored "-Wunknown-pragmas" // disable warning about below #pragma warning being unknown
#ifdef _MSC_VER                                    // additional condition because the previous line does not work on gcc 12.2
#   pragma warning(pop)
#endif // _MSC_VER
#pragma GCC diagnostic pop
#endif    // OPENXLSX_XLCELL_HPP

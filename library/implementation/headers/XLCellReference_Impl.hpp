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

#ifndef OPENXLSX_IMPL_XLCELLREFERENCE_H
#define OPENXLSX_IMPL_XLCELLREFERENCE_H

#include <string>

namespace OpenXLSX::Impl
{

    //======================================================================================================================
    //========== XLCoordinates Alias =======================================================================================
    //======================================================================================================================

    /**
     * @brief
     */
    using XLCoordinates = std::pair<unsigned long, unsigned int>;

    //======================================================================================================================
    //========== XLCellReference Class =====================================================================================
    //======================================================================================================================

    /**
     * @brief
     */
    class XLCellReference final
    {
        friend bool operator==(const XLCellReference& lhs, const XLCellReference& rhs);
        friend bool operator!=(const XLCellReference& lhs, const XLCellReference& rhs);
        friend bool operator<(const XLCellReference& lhs, const XLCellReference& rhs);
        friend bool operator>(const XLCellReference& lhs, const XLCellReference& rhs);
        friend bool operator<=(const XLCellReference& lhs, const XLCellReference& rhs);
        friend bool operator>=(const XLCellReference& lhs, const XLCellReference& rhs);

        //----------------------------------------------------------------------------------------------------------------------
        //           Public Member Functions
        //----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief Constructor taking a cell address as argument.
         * @param cellAddress The address of the cell, e.g. 'A1'.
         */
        explicit XLCellReference(const std::string& cellAddress = "");

        /**
         * @brief Constructor taking the cell coordinates as arguments.
         * @param row The row number of the cell.
         * @param column The column number of the cell.
         */
        explicit XLCellReference(unsigned long row, unsigned int column);

        /**
         * @brief Constructor taking the row number and the column letter as arguments.
         * @param row The row number of the cell.
         * @param column The column letter of the cell.
         */
        explicit XLCellReference(unsigned long row, const std::string& column);

        /**
         * @brief Copy constructor
         * @param other The object to be copied.
         */
        XLCellReference(const XLCellReference& other) = default;

        /**
         * @brief Destructor. Default implementation used.
         */
        ~XLCellReference() = default;

        /**
         * @brief Assignment operator.
         * @param other The object to be copied/assigned.
         * @return A reference to the new object.
         */
        XLCellReference& operator=(const XLCellReference& other) = default;

        /**
         * @brief Get the row number of the XLCellReference.
         * @return The row.
         */
        unsigned long Row() const;

        /**
         * @brief Set the row number for the XLCellReference.
         * @param row The row number.
         */
        void SetRow(unsigned long row);

        /**
         * @brief Get the column number of the XLCellReference.
         * @return The column number.
         */
        unsigned int Column() const;

        /**
         * @brief Set the column number of the XLCellReference.
         * @param column The column number.
         */
        void SetColumn(unsigned int column);

        /**
         * @brief Set both row and column number of the XLCellReference.
         * @param row The row number.
         * @param column The column number.
         */
        void SetRowAndColumn(unsigned long row, unsigned int column);

        /**
         * @brief Get the address of the XLCellReference
         * @return The address, e.g. 'A1'
         */
        std::string Address() const;

        /**
         * @brief Set the address of the XLCellReference
         * @param address The address, e.g. 'A1'
         */
        void SetAddress(const std::string& address);

        //----------------------------------------------------------------------------------------------------------------------
        //           Private Member Functions
        //----------------------------------------------------------------------------------------------------------------------

    private:

        /**
         * @brief
         * @param row
         * @return
         */
        static std::string RowAsString(unsigned long row);

        /**
         * @brief
         * @param row
         * @return
         */
        static unsigned long RowAsNumber(const std::string& row);

        /**
         * @brief Static helper function to convert column number to column letter (e.g. column 1 becomes 'A')
         * @param column The column number.
         * @return The column letter
         * @todo Include a cache of column names generated.
         */
        static std::string ColumnAsString(unsigned int column);

        /**
         * @brief Static helper function to convert column letter to column number (e.g. column 'A' becomes 1)
         * @param column The column letter, e.g. 'A'
         * @return The column number.
         * @todo Include a cache of column names generated.
         */
        static unsigned int ColumnAsNumber(const std::string& column);

        /**
         * @brief Static helper function to convert cell address to coordinates.
         * @param address The address to be converted, e.g. 'A1'
         * @return A std::pair<row, column>
         * @todo Consider including a cache of the coordinates requested,
         */
        static XLCoordinates CoordinatesFromAddress(const std::string& address);

        //----------------------------------------------------------------------------------------------------------------------
        //           Private Member Variables
        //----------------------------------------------------------------------------------------------------------------------

    private:

        unsigned long m_row; /**< The row */
        unsigned int m_column; /**< The column */
        std::string m_cellAddress; /**< The address, e.g. 'A1' */
        bool m_valid /**< Flag indicating if the reference is valid. */;

    };

    //======================================================================================================================
    //========== Global Functions ==========================================================================================
    //======================================================================================================================

    /**
     * @brief Helper function to check equality between two XLCellReferences.
     * @param lhs The first XLCellReference
     * @param rhs The second XLCellReference
     * @return true if equal; otherwise false.
     */
    inline bool operator==(const XLCellReference& lhs, const XLCellReference& rhs) {

        bool result = false;
        if (lhs.Row() == rhs.Row() && lhs.Column() == rhs.Column())
            result = true;
        return result;
    }

    /**
     * @brief Helper function to check for in-equality between two XLCellReferences
     * @param lhs The first XLCellReference
     * @param rhs The second XLCellReference
     * @return false if equal; otherwise true.
     */
    inline bool operator!=(const XLCellReference& lhs, const XLCellReference& rhs) {

        return !(lhs == rhs);
    }

    /**
     * @brief Helper function to check if one XLCellReference is smaller than another.
     * @param lhs The first XLCellReference
     * @param rhs The second XLCellReference
     * @return true if lhs < rhs; otherwise false.
     */
    inline bool operator<(const XLCellReference& lhs, const XLCellReference& rhs) {

        bool result;
        if (lhs.Row() < rhs.Row())
            result = true;
        else if (lhs.Row() > rhs.Row())
            result = false;
        else { // lhs.Row() == rhs.Row()
            result = lhs.Column() < rhs.Column();
        }
        return result;
    }

    /**
     * @brief Helper function to check if one XLCellReference is larger than another.
     * @param lhs The first XLCellReference
     * @param rhs The second XLCellReference
     * @return true if lhs > rhs; otherwise false.
     */
    inline bool operator>(const XLCellReference& lhs, const XLCellReference& rhs) {

        return (rhs < lhs);
    }

    /**
     * @brief Helper function to check if one XLCellReference is smaller than or equal to another.
     * @param lhs The first XLCellReference
     * @param rhs The second XLCellReference
     * @return true if lhs <= rhs; otherwise false
     */
    inline bool operator<=(const XLCellReference& lhs, const XLCellReference& rhs) {

        return !(lhs > rhs);
    }

    /**
     * @brief Helper function to check if one XLCellReference is larger than or equal to another.
     * @param lhs The first XLCellReference
     * @param rhs The second XLCellReference
     * @return true if lhs >= rhs; otherwise false.
     */
    inline bool operator>=(const XLCellReference& lhs, const XLCellReference& rhs) {

        return !(lhs < rhs);
    }
} // namespace OpenXLSX::Impl

#endif //OPENXLSX_IMPL_XLCELLREFERENCE_H

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

    Written by Akira SHIMAHARA

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

#ifndef OPENXLSX_XLTABLE_HPP
#define OPENXLSX_XLTABLE_HPP

#pragma warning(push)
//#pragma warning(disable : 4251)
//#pragma warning(disable : 4275)

// ===== External Includes ===== //
#include <vector>
#include <memory>
// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLAutofilter.hpp"
#include "XLXmlData.hpp"
#include "XLTableColumn.hpp"
#include "XLTableRows.hpp"
#include "XLTableStyle.hpp"
#include "XLSheet.hpp"
namespace OpenXLSX
{
    class OPENXLSX_EXPORT XLTable
    {
        friend class XLTableColumn;
        friend class XLTableStyle;

    public:
        /**
         * @brief The constructor. 
         * @param xmlData from the table file
         */
        explicit XLTable(XLXmlData* xmlData);

        /**
         * @brief The destructor. 
         */
        ~XLTable();

        /**
         * @brief 
         * @return the name of the table
         */
        const std::string name() const;

        /**
         * @brief 
         * @return the reference cells of the table (without the worksheet)
         */
        const std::string ref() const;

        /**
         * @brief 
         * @return A vector containing the columns names
         */
        std::vector<std::string> columnNames() const;

        /**
         * @brief
         * @param name  the name of the requested column index
         * @return The index of the column
         */
        uint16_t columnIndex(const std::string& name) const;

         /**
         * @brief
         * @param name  the name of the requested column
         * @return a pointer to the TableColumn object
         */
        XLTableColumn& column(const std::string& name);

        /**
         * @brief
         * @return A pointer to the worksheet object the table belong to
         */
        XLWorksheet* getWorksheet();

        /**
         * @brief
         * @return an XLCellRange object of the whole table 
         * including header and total row if visibles
         */
        XLCellRange tableRange();

        /**
         * @brief
         * @return a ref to the XLCellRange object of data body range of the table
         * exclude header and total row
         */
        const XLCellRange& dataBodyRange() const;

        /**
         * @brief
         * @return a ref to the XLCellRange object of data body range of the table
         * exclude header and total row
         */
        XLCellRange& dataBodyRange();
    

        /**
         * @brief
         * @return an XLTableRows iterable object (only data body range)
         */
        XLTableRows tableRows();


        /**
         * @brief
         * @return true if header is visible
         */
        bool isHeaderVisible() const;
        

        /**
         * @brief
         * @return true if total row is visible
         */
        bool isTotalVisible() const; 

        /**
         * @brief
         * @param visible
         */
        void setHeaderVisible(bool visible = true);

        /**
         * @brief
         * @param visible
         */
        void setTotalVisible(bool visible = true);

        /**
         * @brief
         * @return the autofilter object
         */
        XLAutofilter autofilter();

        /**
         * @brief return the table style obect
         */
        XLTableStyle tableStyle();

        /**
         * @brief
         * @return number of columns
         */
        uint16_t columnsCount() const;

        /**
         * @brief
         * @return number of data rows, excluding header and total row
         */
        uint32_t rowsCount() const;

        /**
         * @brief set the name of the table
         * @param tableName name of the table
         */
        void setName(const std::string& tableName);

    //----------------------------------------------------------------------------------------------------------------------
    //           Protected
    //----------------------------------------------------------------------------------------------------------------------  
    protected:
        /**
         * @brief set the formulas in the worksheet for the total row
         * the formulas is based on the attribute totalsRowFunction of each col
         */
        void setTotalFormulas();

        /**
         * @brief Adjust the ref according to m_dataBodyRange
         * and the state of visibility of headers and total
         */
        void adjustRef();

    //----------------------------------------------------------------------------------------------------------------------
    //           Private Member Variables
    //----------------------------------------------------------------------------------------------------------------------
    private:
        XLXmlData*      m_pXmlData;
        std::vector<XLTableColumn> m_columns;
        XLWorksheet     m_sheet;
        XLCellRange     m_dataBodyRange;    // without header and total
        // cell range
        // sheet
        // filter
        // sort state
        // tablestyle
    };
}    // namespace OpenXLSX

//#pragma warning(pop)
#endif    // OPENXLSX_XLTABLE_HPP

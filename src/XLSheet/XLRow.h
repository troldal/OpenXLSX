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

#ifndef OPENXLEXE_XLROW_H
#define OPENXLEXE_XLROW_H

#include "../XLWorkbook/XLDocument.h"

namespace OpenXLSX
{
    class XLCell;

//======================================================================================================================
//========== XLRow Class ===============================================================================================
//======================================================================================================================

/**
 * @brief The XLRow class represent a row in an Excel spreadsheet. All cell data are stored by row in the underlying
 * XML file.
 */
    class XLRow
    {
        friend class XLWorksheet;

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:


        /**
         * @brief Constructor
         * @param parent A pointer to the parent XLWorksheet object.
         * @param rowNode A pointer to the XMLNode object for the row.
         */
        explicit XLRow(XLWorksheet &parent,
                       XMLNode rowNode);

        /**
         * @brief Copy Constructor
         * @note The copy constructor is explicitly deleted
         */
        XLRow(const XLRow &other) = delete;

        /**
         * @brief Move Constructor
         * @note The move constructor has been explicitly deleted.
         */
        XLRow(XLRow &&other) = delete;

        /**
         * @brief Destructor
         * @note The destructor has a default implementation.
         */
        virtual ~XLRow() = default;

        /**
         * @brief Copy assignment operator.
         * @note The copy assignment operator is explicitly deleted.
         */
        XLRow &operator=(const XLRow &other) = delete;

        /**
         * @brief Move assignment operator.
         * @note The move assignment operator has been explicitly deleted.
         */
        XLRow &operator=(XLRow &&other) = delete;

        /**
         * @brief Get the height of the row.
         * @return the row height.
         */
        float Height() const;

        /**
         * @brief Set the height of the row.
         * @param height The height of the row.
         */
        void SetHeight(float height);

        /**
         * @brief Get the descent of the row, which is the vertical distance in pixels from the bottom of the cells
         * in the current row to the typographical baseline of the cell content.
         * @return The row descent.
         */
        float Descent() const;

        /**
         * @brief Set the descent of the row, ehich is he vertical distance in pixels from the bottom of the cells
         * in the current row to the typographical baseline of the cell content.
         * @param descent The row descent.
         */
        void SetDescent(float descent);

        /**
         * @brief Is the row hidden?
         * @return The state of the row.
         */
        bool Ishidden() const;

        /**
         * @brief Set the row to be hidden or visible.
         * @param state The state of the row.
         */
        void SetHidden(bool state);

        /**
         * @brief Get the XMLNode object for the row.
         * @return The XMLNode for the object.
         */
        XMLNode RowNode();

        /**
         * @brief Get the XLCell object at a specified column for this row.
         * @param column The column with the XLCell
         * @return A reference to the XLCell object.
         */
        XLCell *Cell(unsigned int column);

        /**
         * @brief Get the XLCell object at a specified column for this row.
         * @param column The column with the XLCell
         * @return A const reference to the XLCell object.
         */
        const XLCell *Cell(unsigned int column) const;

        /**
         * @brief Get the number of cells in the row.
         * @return The number of cells in the row.
         */
        unsigned int CellCount() const;


//----------------------------------------------------------------------------------------------------------------------
//           Protected Member Functions
//----------------------------------------------------------------------------------------------------------------------

    protected:


        /**
         * @brief A static method used to create an entirely new XLRow object (no corresponding node in the XML file).
         * @param worksheet A reference to the worksheet object to which the row is to be added.
         * @param rowNumber The row number to add
         * @return A pointer to the newly created XLRow object.
         */
        static std::unique_ptr<XLRow> CreateRow(XLWorksheet &worksheet,
                                                unsigned long rowNumber);

        /**
         * @brief Resize the row, i.e. change the number of cells in the row.
         * @param cellCount The new number of cells in the row.
         */
        void Resize(unsigned int cellCount);

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Variables
//----------------------------------------------------------------------------------------------------------------------

    private:

        XLWorksheet &m_parentWorksheet; /**< A pointer to the parent XLWorksheet object. */
        XLDocument &m_parentDocument; /**< A pointer to the parent XLDocument object. */

        XMLNode m_rowNode; /**< The XMLNode object for the row. */

        float m_height; /**< The height of the row. */
        float m_descent; /**< The descent of the row. */
        bool m_hidden; /**< The hidden state of the row. */

        unsigned long m_rowNumber; /**< The row number of the current row. */

        std::vector<std::unique_ptr<XLCell>> m_cells; /**< A vector with the XLCell objects. */
    };

}


#endif //OPENXLEXE_XLROW_H

//
// Created by KBA012 on 02/06/2017.
//

#ifndef OPENXLEXE_XLROW_H
#define OPENXLEXE_XLROW_H

#include "XML/XMLNode.h"
#include "XLDocument.h"

namespace RapidXLSX
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
                       XMLNode &rowNode);

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
        XMLNode &RowNode();

        /**
         * @brief Get the XLCell object at a specified column for this row.
         * @param column The column with the XLCell
         * @return A reference to the XLCell object.
         */
        XLCell &Cell(unsigned int column);

        /**
         * @brief Get the XLCell object at a specified column for this row.
         * @param column The column with the XLCell
         * @return A const reference to the XLCell object.
         */
        const XLCell &Cell(unsigned int column) const;

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

        XMLNode &m_rowNode; /**< The XMLNode object for the row. */

        float m_height; /**< The height of the row. */
        float m_descent; /**< The descent of the row. */
        bool m_hidden; /**< The hidden state of the row. */

        unsigned long m_rowNumber; /**< The row number of the current row. */

        std::vector<std::unique_ptr<XLCell>> m_cells; /**< A vector with the XLCell objects. */
    };

}


#endif //OPENXLEXE_XLROW_H

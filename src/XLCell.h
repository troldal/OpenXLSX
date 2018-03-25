//
// Created by Troldal on 02/09/16.
//

#ifndef OPENXL_XLCELL_H
#define OPENXL_XLCELL_H

#include <string>
#include <ostream>
#include "XML/XMLNode.h"
#include "XLCellReference.h"
#include "XLDocument.h"
#include "XLCellType.h"

namespace RapidXLSX
{
    class XLCellRange;
    class XLCellValue;


//======================================================================================================================
//========== XLCell Class ==============================================================================================
//======================================================================================================================

    /**
     * @brief A class encapsulating the properties and behaviours of a spreadsheet cell.
     * @todo Consider using template functions for getting the value, e.g. 'value<Number>...'
     */
    class XLCell: public XLSpreadsheetElement
    {
        friend class XLRow;
        friend class XLCellValue;
        friend class XLValue;
        friend class XLValueNumber;
        friend class XLValueString;

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief Copy constructor
         * @param other The XLCell object to be copied.
         * @note The copy constructor has been deleted, as it makes no sense to copy a cell. If the objective is to
         * copy the value, create the the target object and then use the copy assignment operator.
         */
        XLCell(const XLCell &other) = delete;

        /**
         * @brief Move constructor
         * @param other The XLCell object to be moved
         * @note The move constructor has been deleted, as it makes no sense to move a cell.
         */
        XLCell(XLCell &&other) = delete;

        /**
         * @brief Destructor
         * @note Using the default destructor
         */
        ~XLCell() override = default;

        /**
         * @brief Copy assignment operator
         * @param other The XLCell object to be copy assigned
         * @return A reference to the new object
         * @note Copies only the cell contents, not the pointer to parent worksheet etc.
         */
        XLCell &operator=(const XLCell &other);

        /**
         * @brief This copy assignment operators takes a range as the argument. The purpose is to copy the range to a
         * new location, with the target cell being the top left cell in the range.
         * @param range The range to be copied
         * @return A reference to the new cell object.
         * @note Copies only the cell contents (values).
         * @todo Consider if this is better done as a free function. What should be the return value?
         */
        XLCell &operator=(const XLCellRange &range);

        /**
         * @brief Move assignment operator [deleted]
         * @param other The XLCell object to be move assigned
         * @return A reference to the new object
         * @note The move assignment constructor has been deleted, as it makes no sense to move a cell.
         */
        XLCell &operator=(XLCell &&other) = delete;

        /**
         * @brief Get a reference to the XLCellValue object for the cell.
         * @return A reference to an XLCellValue object.
         */
        XLCellValue &Value();

        /**
         * @brief Get a const reference to the XLCellValue object for the cell.
         * @return A const reference to an XLCellValue object.
         */
        const XLCellValue &Value() const;

        /**
         * @brief Set the cell to a 'dirty' state, i.e. the worksheet needs to be saved.
         */
        void SetModified();

        /**
         * @brief
         * @return
         */
        XLValueType ValueType() const;

        /**
         * @brief get the XLCellReference object for the cell.
         * @return A reference to the cells' XLCellReference object.
         */
        const XLCellReference &CellReference() const;

//----------------------------------------------------------------------------------------------------------------------
//           Protected Member Functions
//----------------------------------------------------------------------------------------------------------------------

    protected:

        /**
         * @brief Constructor
         * @param parent A pointer to the parent XLWorksheet object. Must not be nullptr.
         * @param cellNode A pointer to the XMLNode with the cell data. Must not be nullptr.
         */
        explicit XLCell(XLWorksheet &parent,
                        XMLNode &cellNode);

        /**
         * @brief Factory method for creating a new cell
         * @param parent A pointer to the parent XLWorksheet object. Must not be nullptr.
         * @param cellNode A pointer to the XMLNode with the cell data. Must not be nullptr.
         * @return A std::unique_ptr to the newly created object
         */
        static std::unique_ptr<XLCell> CreateCell(XLWorksheet &parent,
                                                  XMLNode &cellNode);

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Functions
//----------------------------------------------------------------------------------------------------------------------

    private:

        /**
         * @brief
         * @return
         */
        XLWorksheet &ParentWorksheet();

        /**
         * @brief
         * @return
         */
        const XLWorksheet &ParentWorksheet() const;

        /**
         * @brief
         * @return
         */
        XMLDocument &XmlDocument();

        /**
         * @brief
         * @return
         */
        const XMLDocument &XmlDocument() const;

        /**
         * @brief
         * @return
         */
        XMLNode *CellNode();

        /**
         * @brief
         * @return
         */
        const XMLNode *CellNode() const;

        /**
         * @brief 
         * @return 
         */
        bool HasTypeAttribute() const;
        
        /**
         * @brief
         * @return
         */
        const XMLAttribute *TypeAttribute() const;

        /**
         * @brief
         * @param typeString
         */
        void SetTypeAttribute(const std::string &typeString = "");

        /**
         * @brief
         */
        void DeleteTypeAttribute();

        /**
         * @brief
         */
        XMLNode *CreateValueNode();


//----------------------------------------------------------------------------------------------------------------------
//           Private Member Variables
//----------------------------------------------------------------------------------------------------------------------

    private:

        XLDocument *m_parentDocument; /**< A pointer to the parent XLDocument object. */
        const XLWorkbook *m_parentWorkbook; /**< A pointer to the parent XLWorkbook object. */
        XLWorksheet *m_parentWorksheet; /**< A pointer to the parent XLWorksheet object. */

        XLCellReference m_cellReference; /**< The cell reference variable. */

        XMLNode *m_rowNode; /**< A pointer to the row node to which the cell belongs. */
        XMLNode *m_cellNode; /**< A pointer to the root XMLNode for the cell. */
        XMLNode *m_formulaNode; /**< A pointer to the formula XMLNode, a child node of the m_cellNode. */

        std::unique_ptr<XLCellValue> m_value; /**<  */

    };
}

#endif //OPENXL_XLCELL_H
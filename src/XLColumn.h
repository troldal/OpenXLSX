//
// Created by Kenneth Balslev on 02/06/2017.
//

#ifndef OPENXLEXE_XLCOLUMN_H
#define OPENXLEXE_XLCOLUMN_H

#include "XML/XMLNode.h"
#include "XLDocument.h"

namespace RapidXLSX
{

//======================================================================================================================
//========== XLWorksheet Class =========================================================================================
//======================================================================================================================

    /**
     * @brief
     */
    class XLColumn
    {

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief Constructor
         * @param parent A pointer to the parent XLWorksheet object.
         * @param columnNode A pointer to the XMLNode for the column.
         */
        explicit XLColumn(XLWorksheet &parent,
                          XMLNode &columnNode);

        /**
         * @brief Copy Constructor [deleted]
         */
        XLColumn(const XLColumn &other) = delete;

        /**
         * @brief Destructor
         */
        virtual ~XLColumn() = default;

        /**
         * @brief Copy assignment operator [deleted]
         */
        XLColumn &operator=(const XLColumn &other) = delete;

        /**
         * @brief Get the width of the column.
         * @return The width of the column.
         */
        float Width() const;

        /**
         * @brief Set the width of the column
         * @param width The width of the column
         */
        void SetWidth(float width);

        /**
         * @brief Is the column hidden?
         * @return The state of the column.
         */
        bool Ishidden() const;

        /**
         * @brief Set the column to be shown or hidden.
         * @param state The state of the column.
         */
        void SetHidden(bool state);

        /**
         * @brief Get the XMLNode object for the column.
         * @return The XMLNode for the column
         */
        XMLNode &ColumnNode();

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Variables
//----------------------------------------------------------------------------------------------------------------------

    private:
        XLWorksheet *m_parentWorksheet; /**< A pointer to the parent XLWorksheet object. */
        XLDocument *m_parentDocument; /**< A pointer to the parent XLDocument object. */

        XMLNode *m_columnNode; /**< A pointer to the XMLNode object for the column. */

        float m_width; /**< The width of the column */
        bool m_hidden; /**< The hidden state of the column */

        unsigned long m_column; /**< The column number for the column */

    };

}


#endif //OPENXLEXE_XLCOLUMN_H

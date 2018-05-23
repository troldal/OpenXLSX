//
// Created by KBA012 on 18-12-2017.
//

#include "XLSpreadsheetElement.h"
#include "XLDocument.h"

namespace OpenXLSX
{
    /**
     * @details
     */
    XLSpreadsheetElement::XLSpreadsheetElement(XLDocument &parent)
        : m_document(&parent),
          m_workbook(parent.Workbook())
    {

    }

    /**
     * @details
     */
    XLWorkbook *XLSpreadsheetElement::ParentWorkbook()
    {
        return const_cast<XLWorkbook *>(static_cast<const XLSpreadsheetElement *>(this)->ParentWorkbook());
    }

    /**
     * @details
     */
    const XLWorkbook * XLSpreadsheetElement::ParentWorkbook() const
    {
        return m_workbook;
    }

    /**
     * @details
     */
    XLDocument * XLSpreadsheetElement::ParentDocument()
    {
        return const_cast<XLDocument *>(static_cast<const XLSpreadsheetElement *>(this)->ParentDocument());
    }

    /**
     * @details
     */
    const XLDocument * XLSpreadsheetElement::ParentDocument() const
    {
        return m_document;
    }
}
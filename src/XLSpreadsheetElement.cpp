//
// Created by KBA012 on 18-12-2017.
//

#include "XLSpreadsheetElement.h"
#include "XLDocument.h"

namespace OpenXLSX
{
    XLSpreadsheetElement::XLSpreadsheetElement(XLDocument &parent)
        : m_document(parent),
          m_workbook(parent.Workbook())
    {

    }

    XLWorkbook &XLSpreadsheetElement::ParentWorkbook()
    {
        return const_cast<XLWorkbook &>(static_cast<const XLSpreadsheetElement *>(this)->ParentWorkbook());
    }

    const XLWorkbook &XLSpreadsheetElement::ParentWorkbook() const
    {
        return m_workbook;
    }

    XLDocument &XLSpreadsheetElement::ParentDocument()
    {
        return const_cast<XLDocument &>(static_cast<const XLSpreadsheetElement *>(this)->ParentDocument());
    }

    const XLDocument &XLSpreadsheetElement::ParentDocument() const
    {
        return m_document;
    }
}
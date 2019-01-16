//
// Created by KBA012 on 18-12-2017.
//

#include "XLSpreadsheetElement_Impl.h"
#include "XLDocument_Impl.h"

using namespace OpenXLSX;

/**
 * @details
 */
Impl::XLSpreadsheetElement::XLSpreadsheetElement(XLDocument& parent)
        : m_document(&parent),
          m_workbook(parent.Workbook()) {

}

/**
 * @details
 */
Impl::XLWorkbook* Impl::XLSpreadsheetElement::ParentWorkbook() {

    return const_cast<XLWorkbook*>(static_cast<const XLSpreadsheetElement*>(this)->ParentWorkbook());
}

/**
 * @details
 */
const Impl::XLWorkbook* Impl::XLSpreadsheetElement::ParentWorkbook() const {

    return m_workbook;
}

/**
 * @details
 */
Impl::XLDocument* Impl::XLSpreadsheetElement::ParentDocument() {

    return const_cast<XLDocument*>(static_cast<const XLSpreadsheetElement*>(this)->ParentDocument());
}

/**
 * @details
 */
const Impl::XLDocument* Impl::XLSpreadsheetElement::ParentDocument() const {

    return m_document;
}

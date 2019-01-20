//
// Created by Troldal on 2019-01-10.
//

#include <XLDocument.h>
#include <XLDocument_Impl.h>

using namespace OpenXLSX;

/**
 * @details
 */
XLDocument::XLDocument()
        : m_document(std::make_shared<Impl::XLDocument>()) {
}

/**
 * @details
 */
XLDocument::XLDocument(const std::string& name)
        : m_document(std::make_shared<Impl::XLDocument>(name)) {
}

/**
 * @details
 */
void XLDocument::OpenDocument(const std::string& fileName) {

    m_document = std::make_shared<Impl::XLDocument>();
    m_document->OpenDocument(fileName);
}

/**
 * @details
 */
void XLDocument::CreateDocument(const std::string& fileName) {

    m_document->CreateDocument(fileName);
}

/**
 * @details
 */
void XLDocument::CloseDocument() {

    m_document->CloseDocument();
    //m_document = nullptr;
}

/**
 * @details
 */
bool XLDocument::SaveDocument() {

    return m_document->SaveDocument();
}

/**
 * @details
 */
bool XLDocument::SaveDocumentAs(const std::string& fileName) {

    return m_document->SaveDocumentAs(fileName);
}

/**
 * @details
 */
std::string XLDocument::DocumentName() const {

    return m_document->DocumentName();
}

/**
 * @details
 */
std::string XLDocument::DocumentPath() const {

    return m_document->DocumentPath();
}

/**
 * @details
 */
XLWorkbook XLDocument::Workbook() {

    return XLWorkbook(*m_document->Workbook());
}

/**
 * @details
 */
const XLWorkbook XLDocument::Workbook() const {

    return XLWorkbook(*m_document->Workbook());
}

/**
 * @details
 */
std::string XLDocument::GetProperty(XLProperty theProperty) const {

    return m_document->GetProperty(theProperty);
}

/**
 * @details
 */
void XLDocument::SetProperty(XLProperty theProperty,
                             const std::string& value) {

    m_document->SetProperty(theProperty,
                            value);
}

/**
 * @details
 */
void XLDocument::DeleteProperty(XLProperty theProperty) {

    m_document->DeleteProperty(theProperty);
}


//
// Created by Troldal on 2019-01-10.
//

#include <XLDocument.hpp>
#include <XLDocument_Impl.hpp>

using namespace OpenXLSX;

/**********************************************************************************************************************/
XLDocument::XLDocument()
        : m_document(std::make_shared<Impl::XLDocument>()) {
}

/**********************************************************************************************************************/
XLDocument::XLDocument(const std::string& name)
        : m_document(std::make_shared<Impl::XLDocument>(name)) {
}

/**********************************************************************************************************************/
void XLDocument::OpenDocument(const std::string& fileName) {

    m_document = std::make_shared<Impl::XLDocument>();
    m_document->OpenDocument(fileName);
}

/**********************************************************************************************************************/
void XLDocument::CreateDocument(const std::string& fileName) {

    if (!m_document)
        m_document = std::make_shared<Impl::XLDocument>();
    m_document->CreateDocument(fileName);
}

/**********************************************************************************************************************/
void XLDocument::CloseDocument() {

    if (!m_document)
        throw XLException("Invalid XLDocument object!");
    m_document->CloseDocument();
    m_document = nullptr;
}

/**********************************************************************************************************************/
bool XLDocument::SaveDocument() {

    if (!m_document)
        throw XLException("Invalid XLDocument object!");
    return m_document->SaveDocument();
}

/**********************************************************************************************************************/
bool XLDocument::SaveDocumentAs(const std::string& fileName) {

    if (!m_document)
        throw XLException("Invalid XLDocument object!");
    return m_document->SaveDocumentAs(fileName);
}

/**********************************************************************************************************************/
const std::string& XLDocument::DocumentName() const {

    if (!m_document)
        throw XLException("Invalid XLDocument object!");
    return m_document->DocumentName();
}

/**********************************************************************************************************************/
const std::string& XLDocument::DocumentPath() const {

    if (!m_document)
        throw XLException("Invalid XLDocument object!");
    return m_document->DocumentPath();
}

/**********************************************************************************************************************/
XLWorkbook XLDocument::Workbook() {

    if (!m_document)
        throw XLException("Invalid XLDocument object!");
    return XLWorkbook(*m_document->Workbook());
}

/**********************************************************************************************************************/
const XLWorkbook XLDocument::Workbook() const {

    if (!m_document)
        throw XLException("Invalid XLDocument object!");
    return static_cast<const XLWorkbook>(*m_document->Workbook());
}

/**********************************************************************************************************************/
std::string XLDocument::GetProperty(XLProperty theProperty) const {

    if (!m_document)
        throw XLException("Invalid XLDocument object!");
    return m_document->GetProperty(theProperty);
}

/**********************************************************************************************************************/
void XLDocument::SetProperty(XLProperty theProperty, const std::string& value) {

    if (!m_document)
        throw XLException("Invalid XLDocument object!");
    m_document->SetProperty(theProperty, value);
}

/**********************************************************************************************************************/
void XLDocument::DeleteProperty(XLProperty theProperty) {

    if (!m_document)
        throw XLException("Invalid XLDocument object!");
    m_document->DeleteProperty(theProperty);
}

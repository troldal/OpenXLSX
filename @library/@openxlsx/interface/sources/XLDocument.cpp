//
// Created by Troldal on 2019-01-10.
//

#include <XLDocument.h>
#include <XLDocument_Impl.h>

using namespace OpenXLSX;

XLDocument::XLDocument()
        : m_document(std::make_shared<Impl::XLDocument>()) {
}

XLDocument::XLDocument(const std::string& name)
        : m_document(std::make_shared<Impl::XLDocument>(name)) {
}

void XLDocument::OpenDocument(const std::string& fileName) {
    m_document = std::make_shared<Impl::XLDocument>();
    m_document->OpenDocument(fileName);
}

void XLDocument::CreateDocument(const std::string& fileName) {
    m_document->CreateDocument(fileName);
}

void XLDocument::CloseDocument() {
    m_document->CloseDocument();
    m_document = nullptr;
}

bool XLDocument::SaveDocument() {
    return m_document->SaveDocument();
}

bool XLDocument::SaveDocumentAs(const std::string& fileName) {
    return m_document->SaveDocumentAs(fileName);
}

std::string XLDocument::DocumentName() const {
    return m_document->DocumentName();
}

std::string XLDocument::DocumentPath() const {
    return m_document->DocumentPath();
}

XLWorkbook XLDocument::Workbook() {
    return XLWorkbook(*m_document->Workbook());
}

const XLWorkbook XLDocument::Workbook() const {
    return XLWorkbook(*m_document->Workbook());
}

std::string XLDocument::GetProperty(XLProperty theProperty) const {
    return m_document->GetProperty(theProperty);
}

void XLDocument::SetProperty(XLProperty theProperty,
                             const std::string& value) {
    m_document->SetProperty(theProperty,
                            value);
}

void XLDocument::DeleteProperty(XLProperty theProperty) {
    m_document->DeleteProperty(theProperty);
}


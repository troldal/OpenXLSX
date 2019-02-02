//
// Created by Troldal on 2019-01-31.
//

#include "XLDocument_Impl.h"
#include "XLProperty.h"

extern "C" {
#include "XL_Document.h"
}

static OpenXLSX::XLProperty ConvertProperty(XL_Property property) {

    switch (property) {
        case XL_Property::Application :
            return OpenXLSX::XLProperty::Application;

        case XL_Property::AppVersion :
            return OpenXLSX::XLProperty::AppVersion;

        case XL_Property::Category :
            return OpenXLSX::XLProperty::Category;

        case XL_Property::Company :
            return OpenXLSX::XLProperty::Company;

        case XL_Property::CreationDate :
            return OpenXLSX::XLProperty::CreationDate;

        case XL_Property::Creator :
            return OpenXLSX::XLProperty::Creator;

        case XL_Property::Description :
            return OpenXLSX::XLProperty::Description;

        case XL_Property::DocSecurity :
            return OpenXLSX::XLProperty::DocSecurity;

        case XL_Property::HyperlinkBase :
            return OpenXLSX::XLProperty::HyperlinkBase;

        case XL_Property::HyperlinksChanged :
            return OpenXLSX::XLProperty::HyperlinksChanged;

        case XL_Property::Keywords :
            return OpenXLSX::XLProperty::Keywords;

        case XL_Property::LastModifiedBy :
            return OpenXLSX::XLProperty::LastModifiedBy;

        case XL_Property::LastPrinted :
            return OpenXLSX::XLProperty::LastPrinted;

        case XL_Property::LinksUpToDate :
            return OpenXLSX::XLProperty::LinksUpToDate;

        case XL_Property::Manager :
            return OpenXLSX::XLProperty::Manager;

        case XL_Property::ModificationDate :
            return OpenXLSX::XLProperty::ModificationDate;

        case XL_Property::ScaleCrop :
            return OpenXLSX::XLProperty::ScaleCrop;

        case XL_Property::SharedDoc :
            return OpenXLSX::XLProperty::SharedDoc;

        case XL_Property::Subject :
            return OpenXLSX::XLProperty::Subject;

        case XL_Property::Title :
            return OpenXLSX::XLProperty::Title;
    }
}

XL_Document* XL_CreateDocument(const char *name) {

    auto result = new OpenXLSX::Impl::XLDocument();
    result->CreateDocument(name);
    return reinterpret_cast<XL_Document*>(result);
}

XL_Document* XL_OpenDocument(const char *name) {

    auto result = new OpenXLSX::Impl::XLDocument();
    result->OpenDocument(name);
    return reinterpret_cast<XL_Document*>(result);
}

void XL_DestroyDocument(XL_Document* doc) {

    delete reinterpret_cast<OpenXLSX::Impl::XLDocument*>(doc);
}

void XL_CloseDocument(XL_Document* doc) {

    reinterpret_cast<OpenXLSX::Impl::XLDocument*>(doc)->CloseDocument();
}

void XL_SaveDocument(XL_Document* doc) {

    reinterpret_cast<OpenXLSX::Impl::XLDocument*>(doc)->SaveDocument();
}

void XL_SaveDocumentAs(XL_Document* doc, const char* name) {

    reinterpret_cast<OpenXLSX::Impl::XLDocument*>(doc)->SaveDocumentAs(name);
}

const char * XL_DocumentName(XL_Document *doc)
{
    return reinterpret_cast<OpenXLSX::Impl::XLDocument*>(doc)->DocumentName().c_str();
}

const char* XL_DocumentPath(XL_Document* doc) {

    return reinterpret_cast<OpenXLSX::Impl::XLDocument*>(doc)->DocumentPath().c_str();
}

//XL_Workbook* XL_GetWorkbook(XL_Document* doc);

const char* XL_GetDocumentProperty(XL_Document* doc, XL_Property property) {

    return reinterpret_cast<OpenXLSX::Impl::XLDocument*>(doc)->GetProperty(ConvertProperty(property)).c_str();
}

void XL_SetDocumentProperty(XL_Document* doc, XL_Property property, const char* value) {

    reinterpret_cast<OpenXLSX::Impl::XLDocument*>(doc)->SetProperty(ConvertProperty(property), value);
}

void XL_DeleteDocumentProperty(XL_Document* doc, XL_Property property) {

    reinterpret_cast<OpenXLSX::Impl::XLDocument*>(doc)->DeleteProperty(ConvertProperty(property));
}
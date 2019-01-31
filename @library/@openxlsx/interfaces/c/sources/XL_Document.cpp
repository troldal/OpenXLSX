//
// Created by Troldal on 2019-01-31.
//

#include "XLDocument.h"

extern "C" {
#include "XL_Document.h"
}

XL_Document* XL_DocumentCreate(const char *name)
{
    auto result = new OpenXLSX::XLDocument();
    result->CreateDocument(name);
    return reinterpret_cast<XL_Document*>(result);
}

const char * XL_DocumentPath(XL_Document *doc)
{
    return reinterpret_cast<OpenXLSX::XLDocument*>(doc)->DocumentPath().c_str();
}

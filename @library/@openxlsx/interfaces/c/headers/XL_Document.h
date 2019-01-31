//
// Created by Troldal on 2019-01-31.
//

#ifndef OPENXLSX_XL_DOCUMENT_H
#define OPENXLSX_XL_DOCUMENT_H

struct XL_Document;
typedef struct XL_Document XL_Document;

XL_Document* XL_DocumentCreate(const char* name);
const char* XL_DocumentPath(XL_Document* doc);

#endif //OPENXLSX_XL_DOCUMENT_H

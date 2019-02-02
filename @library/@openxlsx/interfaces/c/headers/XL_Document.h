//
// Created by Troldal on 2019-01-31.
//

#ifndef OPENXLSX_XL_DOCUMENT_H
#define OPENXLSX_XL_DOCUMENT_H

#include "XL_Property.h"

struct XL_Document;
typedef struct XL_Document XL_Document;

/**
 * @brief
 * @param name
 * @return
 */
XL_Document* XL_CreateDocument(const char *name);

/**
 * @brief
 * @param name
 * @return
 */
XL_Document* XL_OpenDocument(const char *name);

void XL_DestroyDocument(XL_Document* doc);

void XL_CloseDocument(XL_Document* doc);

void XL_SaveDocument(XL_Document* doc);

void XL_SaveDocumentAs(XL_Document* doc, const char* name);

/**
 * @brief
 * @param doc
 * @return
 */
const char* XL_DocumentName(XL_Document *doc);

const char* XL_DocumentPath(XL_Document* doc);

//XL_Workbook* XL_GetWorkbook(XL_Document* doc);

const char* XL_GetDocumentProperty(XL_Document* doc, XL_Property property);

void XL_SetDocumentProperty(XL_Document* doc, XL_Property property, const char* value);

void XL_DeleteDocumentProperty(XL_Document* doc, XL_Property property);


#endif //OPENXLSX_XL_DOCUMENT_H

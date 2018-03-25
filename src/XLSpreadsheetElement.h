//
// Created by KBA012 on 18-12-2017.
//

#ifndef OPENXLEXE_XLWORKBOOKCHILD_H
#define OPENXLEXE_XLWORKBOOKCHILD_H

namespace RapidXLSX
{

    class XLWorkbook;
    class XLDocument;

    class XLSpreadsheetElement
    {
    public:
        explicit XLSpreadsheetElement(XLDocument &parent);
        XLSpreadsheetElement(const XLSpreadsheetElement &other) = default;
        XLSpreadsheetElement(XLSpreadsheetElement &&other) = default;
        virtual ~XLSpreadsheetElement() = default;

        XLSpreadsheetElement &operator=(const XLSpreadsheetElement &other) = delete;
        XLSpreadsheetElement &operator=(XLSpreadsheetElement &&other) = delete;

        virtual XLWorkbook &ParentWorkbook();
        virtual const XLWorkbook &ParentWorkbook() const;

        virtual XLDocument &ParentDocument();
        virtual const XLDocument &ParentDocument() const;

    private:
        XLDocument &m_document;
        XLWorkbook &m_workbook;

    };

}

#endif //OPENXLEXE_XLWORKBOOKCHILD_H

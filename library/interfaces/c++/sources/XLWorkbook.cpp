//
// Created by Troldal on 2019-01-16.
//

#include <XLWorkbook.hpp>
#include <XLWorkbook_Impl.hpp>
#include "XLSheet_Impl.hpp"
#include "XLWorksheet_Impl.hpp"
#include "XLChartsheet_Impl.hpp"

using namespace OpenXLSX;

/**********************************************************************************************************************/
XLWorkbook::XLWorkbook(Impl::XLWorkbook& workbook)
        : m_workbook(&workbook) {
}

/**********************************************************************************************************************/
XLSheet XLWorkbook::Sheet(unsigned int index) {

    return XLSheet(*m_workbook->Sheet(index));
}

/**********************************************************************************************************************/
const XLSheet XLWorkbook::Sheet(unsigned int index) const {

    return static_cast<const XLSheet>(*m_workbook->Sheet(index));
}

/**********************************************************************************************************************/
XLSheet XLWorkbook::Sheet(const std::string& sheetName) {

    return XLSheet(*m_workbook->Sheet(sheetName));
}

/**********************************************************************************************************************/
const XLSheet XLWorkbook::Sheet(const std::string& sheetName) const {

    return static_cast<const XLSheet>(*m_workbook->Sheet(sheetName));
}

/**********************************************************************************************************************/
XLWorksheet XLWorkbook::Worksheet(const std::string& sheetName) {

    return XLWorksheet(*m_workbook->Worksheet(sheetName));
}

/**********************************************************************************************************************/
const XLWorksheet XLWorkbook::Worksheet(const std::string& sheetName) const {

    return static_cast<const XLWorksheet>(*m_workbook->Worksheet(sheetName));
}

/**********************************************************************************************************************/
XLChartsheet XLWorkbook::Chartsheet(const std::string& sheetName) {

    return XLChartsheet(*m_workbook->Chartsheet(sheetName));
}

/**********************************************************************************************************************/
const XLChartsheet XLWorkbook::Chartsheet(const std::string& sheetName) const {

    return static_cast<const XLChartsheet>(*m_workbook->Chartsheet(sheetName));
}

/**********************************************************************************************************************/
void XLWorkbook::DeleteSheet(const std::string& sheetName) {

    m_workbook->DeleteSheet(sheetName);
}

/**********************************************************************************************************************/
void XLWorkbook::AddWorksheet(const std::string& sheetName, unsigned int index) {

    m_workbook->AddWorksheet(sheetName, index);
}

/**********************************************************************************************************************/
void XLWorkbook::CloneWorksheet(const std::string& extName, const std::string& newName, unsigned int index) {

    m_workbook->CloneWorksheet(extName, newName, index);
}

/**********************************************************************************************************************/
void XLWorkbook::AddChartsheet(const std::string& sheetName, unsigned int index) {

    m_workbook->AddChartsheet(sheetName, index);
}

/**********************************************************************************************************************/
void XLWorkbook::MoveSheet(const std::string& sheetName, unsigned int index) {

    m_workbook->MoveSheet(sheetName, index);
}

/**********************************************************************************************************************/
unsigned int XLWorkbook::IndexOfSheet(const std::string& sheetName) const {

    return m_workbook->IndexOfSheet(sheetName);
}

/**********************************************************************************************************************/
XLSheetType XLWorkbook::TypeOfSheet(const std::string& sheetName) const {

    return m_workbook->TypeOfSheet(sheetName);
}

/**********************************************************************************************************************/
XLSheetType XLWorkbook::TypeOfSheet(unsigned int index) const {

    return m_workbook->TypeOfSheet(index);
}

/**********************************************************************************************************************/
unsigned int XLWorkbook::SheetCount() const {

    return m_workbook->SheetCount();
}

/**********************************************************************************************************************/
unsigned int XLWorkbook::WorksheetCount() const {

    return m_workbook->WorksheetCount();
}

/**********************************************************************************************************************/
unsigned int XLWorkbook::ChartsheetCount() const {

    return m_workbook->ChartsheetCount();
}

/**********************************************************************************************************************/
std::vector<std::string> XLWorkbook::SheetNames() const {

    return m_workbook->SheetNames();
}

/**********************************************************************************************************************/
std::vector<std::string> XLWorkbook::WorksheetNames() const {

    return m_workbook->WorksheetNames();
}

/**********************************************************************************************************************/
std::vector<std::string> XLWorkbook::ChartsheetNames() const {

    return m_workbook->ChartsheetNames();
}

/**********************************************************************************************************************/
bool XLWorkbook::SheetExists(const std::string& sheetName) const {

    return m_workbook->SheetExists(sheetName);
}

/**********************************************************************************************************************/
bool XLWorkbook::WorksheetExists(const std::string& sheetName) const {

    return m_workbook->WorksheetExists(sheetName);
}

/**********************************************************************************************************************/
bool XLWorkbook::ChartsheetExists(const std::string& sheetName) const {

    return m_workbook->ChartsheetExists(sheetName);
}

/**********************************************************************************************************************/
void XLWorkbook::DeleteNamedRanges() {

    m_workbook->DeleteNamedRanges();
}
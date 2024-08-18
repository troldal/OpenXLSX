//
// Created by Kenneth Balslev on 24/08/2020.
//

#ifndef OPENXLSX_XLUTILITIES_HPP
#define OPENXLSX_XLUTILITIES_HPP

#include <fstream>
#include <pugixml.hpp>
#include <string>    // 2024-04-25 needed for xml_node_type_string

#include "XLConstants.hpp"        // 2024-05-28 OpenXLSX::MAX_ROWS
#include "XLCellReference.hpp"
#include "XLCellValue.hpp"        // OpenXLSX::XLValueType
#include "XLContentTypes.hpp"     // OpenXLSX::XLContentType
#include "XLRelationships.hpp"    // OpenXLSX::XLRelationshipType
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    /**
     * @brief Get rid of compiler warnings about unused variables (-Wunused-variable) or unused parameters (-Wunusued-parameter)
     * @param (unnamed) the variable / parameter which is intentionally unused (e.g. for function stubs)
     */
    template<class T> void ignore( const T& ) {}

    /**
     * @brief Get a string representation of pugi::xml_node_type
     * @param t the pugi::xml_node_type of a node
     * @return a std::string containing the descriptive name of the node type
     */
    inline std::string XLValueTypeString( XLValueType t )
    {
        using namespace std::literals::string_literals;
        switch( t ) {
            case XLValueType::Empty      : return "Empty"s;
            case XLValueType::Boolean    : return "Boolean"s;
            case XLValueType::Integer    : return "Integer"s;
            case XLValueType::Float      : return "Float"s;
            case XLValueType::Error      : return "Error"s;
            case XLValueType::String     : return "String"s;
        }
        throw XLInternalError(__func__ + "Invalid XLValueType."s);
    }

    /**
     * @brief Get a string representation of pugi::xml_node_type
     * @param t the pugi::xml_node_type of a node
     * @return a std::string containing the descriptive name of the node type
     */
    inline std::string xml_node_type_string( pugi::xml_node_type t )
    {
        using namespace std::literals::string_literals;
        switch( t ) {
            case pugi::node_null        : return "node_null"s;
            case pugi::node_document    : return "node_document"s;
            case pugi::node_element     : return "node_element"s;
            case pugi::node_pcdata      : return "node_pcdata"s;
            case pugi::node_cdata       : return "node_cdata"s;
            case pugi::node_comment     : return "node_comment"s;
            case pugi::node_pi          : return "node_pi"s;
            case pugi::node_declaration : return "node_declaration"s;
            case pugi::node_doctype     : return "node_doctype"s;
        }
        throw XLInternalError("Invalid XML node type.");
    }

    /**
     * @brief Get a string representation of OpenXLSX::XLContentType
     * @param t an OpenXLSX::XLContentType value
     * @return a std::string containing the descriptive name of the content type
     */
    inline std::string XLContentTypeString( OpenXLSX::XLContentType const & t )
    {
        using namespace OpenXLSX;
        switch (t) {
            case XLContentType::Workbook: return "Workbook";
            case XLContentType::WorkbookMacroEnabled: return "WorkbookMacroEnabled";
            case XLContentType::Worksheet: return "Worksheet";
            case XLContentType::Chartsheet: return "Chartsheet";
            case XLContentType::ExternalLink: return "ExternalLink";
            case XLContentType::Theme: return "Theme";
            case XLContentType::Styles: return "Styles";
            case XLContentType::SharedStrings: return "SharedStrings";
            case XLContentType::Drawing: return "Drawing";
            case XLContentType::Chart: return "Chart";
            case XLContentType::ChartStyle: return "ChartStyle";
            case XLContentType::ChartColorStyle: return "ChartColorStyle";
            case XLContentType::ControlProperties: return "ControlProperties";
            case XLContentType::CalculationChain: return "CalculationChain";
            case XLContentType::VBAProject: return "VBAProject";
            case XLContentType::CoreProperties: return "CoreProperties";
            case XLContentType::ExtendedProperties: return "ExtendedProperties";
            case XLContentType::CustomProperties: return "CustomProperties";
            case XLContentType::Comments: return "Comments";
            case XLContentType::Table: return "Table";
            case XLContentType::VMLDrawing: return "VMLDrawing";
            case XLContentType::Unknown: return "Unknown";
        }
        return "invalid";
    }

    /**
     * @brief Get a string representation of OpenXLSX::XLRelationshipType
     * @param t an OpenXLSX::XLRelationshipType value
     * @return std::string containing the descriptive name of the relationship type
     */
    inline std::string XLRelationshipTypeString( OpenXLSX::XLRelationshipType const & t )
    {
        using namespace OpenXLSX;
        switch (t) {
            case XLRelationshipType::CoreProperties: return "CoreProperties";
            case XLRelationshipType::ExtendedProperties: return "ExtendedProperties";
            case XLRelationshipType::CustomProperties: return "CustomProperties";
            case XLRelationshipType::Workbook: return "Workbook";
            case XLRelationshipType::Worksheet: return "Worksheet";
            case XLRelationshipType::Chartsheet: return "Chartsheet";
            case XLRelationshipType::Dialogsheet: return "Dialogsheet";
            case XLRelationshipType::Macrosheet: return "Macrosheet";
            case XLRelationshipType::CalculationChain: return "CalculationChain";
            case XLRelationshipType::ExternalLink: return "ExternalLink";
            case XLRelationshipType::ExternalLinkPath: return "ExternalLinkPath";
            case XLRelationshipType::Theme: return "Theme";
            case XLRelationshipType::Styles: return "Styles";
            case XLRelationshipType::Chart: return "Chart";
            case XLRelationshipType::ChartStyle: return "ChartStyle";
            case XLRelationshipType::ChartColorStyle: return "ChartColorStyle";
            case XLRelationshipType::Image: return "Image";
            case XLRelationshipType::Drawing: return "Drawing";
            case XLRelationshipType::VMLDrawing: return "VMLDrawing";
            case XLRelationshipType::SharedStrings: return "SharedStrings";
            case XLRelationshipType::PrinterSettings: return "PrinterSettings";
            case XLRelationshipType::VBAProject: return "VBAProject";
            case XLRelationshipType::ControlProperties: return "ControlProperties";
            case XLRelationshipType::Comments: return "Comments";
            case XLRelationshipType::Unknown: return "Unknown";
        }
        return "invalid";
    }

    /**
     * @details
     */
    inline XMLNode getRowNode(XMLNode sheetDataNode, uint32_t rowNumber)
    {
        if (rowNumber < 1 || rowNumber > OpenXLSX::MAX_ROWS) {    // 2024-05-28: added range check
            using namespace std::literals::string_literals;
            throw XLCellAddressError("rowNumber "s + std::to_string( rowNumber ) + " is outside valid range [1;"s + std::to_string(OpenXLSX::MAX_ROWS) + "]"s);
        }

        // ===== Get the last child of sheetDataNode that is of type node_element.
        XMLNode result = sheetDataNode.last_child_of_type(pugi::node_element);

        // ===== If there are now rows in the worksheet, or the requested row is beyond the current max row, append a new row to the end.
        if (result.empty() || (rowNumber > result.attribute("r").as_ullong())) {
            result = sheetDataNode.append_child("row");
            result.append_attribute("r") = rowNumber;
            //            result.append_attribute("x14ac:dyDescent") = "0.2";
            //            result.append_attribute("spans")           = "1:1";
        }

        // ===== If the requested node is closest to the end, start from the end and search backwards.
        else if (result.attribute("r").as_ullong() - rowNumber < rowNumber) {
            while (not result.empty() && (result.attribute("r").as_ullong() > rowNumber)) result = result.previous_sibling_of_type(pugi::node_element);
            // ===== If the backwards search failed to locate the requested row
            if (result.empty() || (result.attribute("r").as_ullong() != rowNumber)) {
                if (result.empty())
                    result = sheetDataNode.prepend_child("row"); // insert a new row node at datasheet begin. When saving, this will keep whitespace formatting towards next row node
                else
                    result = sheetDataNode.insert_child_after("row", result);
                result.append_attribute("r") = rowNumber;
                //                result.append_attribute("x14ac:dyDescent") = "0.2";
                //                result.append_attribute("spans")           = "1:1";
            }
        }

        // ===== Otherwise, start from the beginning
        else {
            // ===== At this point, it is guaranteed that there is at least one node_element in the row that is not empty.
            result = sheetDataNode.first_child_of_type(pugi::node_element);

            // ===== It has been verified above that the requested rowNumber is <= the row number of the last node_element, therefore this loop will halt.
            while (result.attribute("r").as_ullong() < rowNumber) result = result.next_sibling_of_type(pugi::node_element);
            // ===== If the forwards search failed to locate the requested row
            if (result.attribute("r").as_ullong() > rowNumber) {
                result = sheetDataNode.insert_child_before("row", result);

                result.append_attribute("r") = rowNumber;
                //                result.append_attribute("x14ac:dyDescent") = "0.2";
                //                result.append_attribute("spans")           = "1:1";
            }
        }

        return result;
    }

    /**
     * @brief get the style attribute s for the indicated column, if any is set
     * @param rowNode the row node from which to obtain the parent that should hold the <cols> node
     * @param colNo the column for which to obtain the style
     * @return the XLStyleIndex stored for the column, or XLDefaultCellFormat if none
     */
    inline XLStyleIndex getColumnStyle(XMLNode rowNode, uint16_t colNo)
    {
        XMLNode cols = rowNode.parent().parent().child("cols");
        if (not cols.empty()) {
            XMLNode col = cols.first_child_of_type(pugi::node_element);
            while (not col.empty()) {
                if (col.attribute("min").as_int(MAX_COLS+1) <= colNo && col.attribute("max").as_int(0) >= colNo) // found
                    return col.attribute("style").as_uint(XLDefaultCellFormat);
               col = col.next_sibling_of_type(pugi::node_element);
            }
        }
        return XLDefaultCellFormat; // if no col style was found
    }

    /**
     * @brief set the cell reference, and a default cell style attribute if and only if row or column style is != XLDefaultCellFormat
     * @note row style takes precedence over column style
     * @param cellNode the cell XML node
     * @param cellRef the cell reference (attribute r) to set
     * @param rowNode the row node for the cell
     * @param colNo the column number for this cell (to try and fetch a column style)
     * @param colStyles an optional std::vector<XLStyleIndex> that contains all pre-evaluated column styles,
     *         and can be used to avoid performance impact from lookup
     */
    inline void setDefaultCellAttributes(XMLNode cellNode, const std::string & cellRef, XMLNode rowNode, uint16_t colNo,
    /**/                                 std::vector<XLStyleIndex> const & colStyles = {})
    {
        cellNode.append_attribute("r").set_value(cellRef.c_str());
        XMLAttribute rowStyle = rowNode.attribute("s");
        XLStyleIndex cellStyle;    // scope declaration
        if (rowStyle.empty()) {    // if no explicit row style is set
            if (colStyles.size() > 0)    // if colStyles were provided
                cellStyle = (colNo > colStyles.size() ? XLDefaultCellFormat : colStyles[ colNo - 1 ]);    // determine cellStyle from colStyles
            else                         // else: no colStyles provided
                cellStyle = getColumnStyle(rowNode, colNo);                                               // so perform a lookup
        }
        else                       // else: an explicit row style is set
            cellStyle = rowStyle.as_uint(XLDefaultCellFormat);    // use the row style

        if (cellStyle != XLDefaultCellFormat) // if cellStyle was determined as not the default style (no point in setting that)
            cellNode.append_attribute("s").set_value(cellStyle);
    }

    /**
     * @brief Retrieve the xml node representing the cell at the given row and column. If the node doesn't
     * exist, it will be created.
     * @param rowNode The row node under which to find the cell.
     * @param columnNumber The column at which to find the cell.
     * @param rowNumber (optional) row number of the row node, if already known, defaults to 0
     * @param colStyles an optional std::vector<XLStyleIndex> that contains all pre-evaluated column styles,
     *         and can be used to avoid performance impact from lookup
     * @return The xml node representing the requested cell.
     */
    inline XMLNode getCellNode(XMLNode rowNode, uint16_t columnNumber, uint32_t rowNumber = 0,
    /**/                       std::vector<XLStyleIndex> const & colStyles = {})
    {
        if (columnNumber < 1 || columnNumber > OpenXLSX::MAX_COLS) {    // 2024-08-05: added range check
            using namespace std::literals::string_literals;
            throw XLException("XLWorksheet::column: columnNumber "s + std::to_string(columnNumber) + " is outside allowed range [1;"s + std::to_string(MAX_COLS) + "]"s);
        }
        if (rowNode.empty()) return XMLNode{};    // 2024-05-28: return an empty node in case of empty rowNode

        XMLNode cellNode = rowNode.last_child_of_type(pugi::node_element);
        if (!rowNumber) rowNumber = rowNode.attribute("r").as_uint(); // if not provided, determine from rowNode
        auto cellRef  = XLCellReference(rowNumber, columnNumber);

        // ===== If there are no cells in the current row, or the requested cell is beyond the last cell in the row...
        if (cellNode.empty() || (XLCellReference(cellNode.attribute("r").value()).column() < columnNumber)) {
            // ===== append a new node to the end.
            cellNode = rowNode.append_child("c");
            setDefaultCellAttributes(cellNode, cellRef.address(), rowNode, columnNumber, colStyles);
        }
        // ===== If the requested node is closest to the end, start from the end and search backwards...
        else if (XLCellReference(cellNode.attribute("r").value()).column() - columnNumber < columnNumber) {
            while (not cellNode.empty() && (XLCellReference(cellNode.attribute("r").value()).column() > columnNumber))
                cellNode = cellNode.previous_sibling_of_type(pugi::node_element);
            // ===== If the backwards search failed to locate the requested cell
            if (cellNode.empty() || (XLCellReference(cellNode.attribute("r").value()).column() < columnNumber)) {
                if (cellNode.empty()) // If between row begin and higher column number, only non-element nodes exist
                    cellNode = rowNode.prepend_child("c"); // insert a new cell node at row begin. When saving, this will keep whitespace formatting towards next cell node
                else
                    cellNode = rowNode.insert_child_after("c", cellNode);
                setDefaultCellAttributes(cellNode, cellRef.address(), rowNode, columnNumber, colStyles);
            }
        }
        // ===== Otherwise, start from the beginning
        else {
            // ===== At this point, it is guaranteed that there is at least one node_element in the row that is not empty.
            cellNode = rowNode.first_child_of_type(pugi::node_element);

            // ===== It has been verified above that the requested columnNumber is <= the column number of the last node_element, therefore this loop will halt:
            while (XLCellReference(cellNode.attribute("r").value()).column() < columnNumber)
                cellNode = cellNode.next_sibling_of_type(pugi::node_element);
            // ===== If the forwards search failed to locate the requested cell
            if (XLCellReference(cellNode.attribute("r").value()).column() > columnNumber) {
                cellNode = rowNode.insert_child_before("c", cellNode);
                setDefaultCellAttributes(cellNode, cellRef.address(), rowNode, columnNumber, colStyles);
            }
        }
        return cellNode;
    }
}    // namespace OpenXLSX

#endif    // OPENXLSX_XLUTILITIES_HPP

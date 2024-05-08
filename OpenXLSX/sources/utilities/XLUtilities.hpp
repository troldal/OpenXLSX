//
// Created by Kenneth Balslev on 24/08/2020.
//

#ifndef OPENXLSX_XLUTILITIES_HPP
#define OPENXLSX_XLUTILITIES_HPP

#include <fstream>
#include <pugixml.hpp>
#include <string>   // 2024-04-25 needed for xml_node_type_string

#include "XLCellReference.hpp"
#include "XLCellValue.hpp"    // OpenXLSX::XLValueType
#include "XLContentTypes.hpp" // OpenXLSX::XLContentType
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
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
     * @details
     */
    inline XMLNode getRowNode(XMLNode sheetDataNode, uint32_t rowNumber)
    {
        // ===== Get the last child of sheetDataNode that is of type node_element.
        auto result = XMLNode();
        result = sheetDataNode.last_child_of_type(pugi::node_element);

        // ===== If there are now rows in the worksheet, or the requested row is beyond the current max row, append a new row to the end.
        if (result.empty() || (rowNumber > result.attribute("r").as_ullong())) {
            result = sheetDataNode.append_child("row");
            result.append_attribute("r") = rowNumber;
            //            result.append_attribute("x14ac:dyDescent") = "0.2";
            //            result.append_attribute("spans")           = "1:1";
        }

        // ===== If the requested node is closest to the end, start from the end and search backwards.
        else if (result.attribute("r").as_ullong() - rowNumber < rowNumber) {
            while (!result.empty() && (result.attribute("r").as_ullong() > rowNumber)) result = result.previous_sibling_of_type(pugi::node_element);
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
            if ((result.type() != pugi::node_element) || (result.attribute("r").as_ullong() > rowNumber)) {
                result = sheetDataNode.insert_child_before("row", result);

                result.append_attribute("r") = rowNumber;
                //                result.append_attribute("x14ac:dyDescent") = "0.2";
                //                result.append_attribute("spans")           = "1:1";
            }
        }

        return result;
    }

    /**
     * @brief Retrieve the xml node representing the cell at the given row and column. If the node doesn't
     * exist, it will be created.
     * @param rowNode The row node under which to find the cell.
     * @param columnNumber The column at which to find the cell.
     * @param rowNumber (optional) row number of the row node, if already known, defaults to 0
     * @return The xml node representing the requested cell.
     */
    inline XMLNode getCellNode(XMLNode rowNode, uint16_t columnNumber, uint32_t rowNumber = 0 )
    {
        auto cellNode = XMLNode();
        cellNode = rowNode.last_child_of_type(pugi::node_element);
        if( !rowNumber ) rowNumber = rowNode.attribute("r").as_uint(); // if not provided, determine from rowNode
        auto cellRef  = XLCellReference(rowNumber, columnNumber);

        // ===== If there are no cells in the current row, or the requested cell is beyond the last cell in the row...
        if (cellNode.empty() || (XLCellReference(cellNode.attribute("r").value()).column() < columnNumber)) {
            // ===== append a new node to the end.
            rowNode.append_child("c").append_attribute("r").set_value(cellRef.address().c_str());
            cellNode = rowNode.last_child();
        }
        // ===== If the requested node is closest to the end, start from the end and search backwards...
        else if (XLCellReference(cellNode.attribute("r").value()).column() - columnNumber < columnNumber) {
            while (!cellNode.empty() && (XLCellReference(cellNode.attribute("r").value()).column() > columnNumber)) cellNode = cellNode.previous_sibling_of_type(pugi::node_element);
            // ===== If the backwards search failed to locate the requested cell
            if (cellNode.empty() || (XLCellReference(cellNode.attribute("r").value()).column() < columnNumber)) {
                if (cellNode.empty()) // If between row begin and higher column number, only non-element nodes exist
                    cellNode = rowNode.prepend_child("c"); // insert a new cell node at row begin. When saving, this will keep whitespace formatting towards next cell node
                else
                    cellNode = rowNode.insert_child_after("c", cellNode);
                cellNode.append_attribute("r").set_value(cellRef.address().c_str());
            }
        }
        // ===== Otherwise, start from the beginning
        else {
            // ===== At this point, it is guaranteed that there is at least one node_element in the row that is not empty.
            cellNode = rowNode.first_child_of_type(pugi::node_element);

            // ===== It has been verified above that the requested columnNumber is <= the column number of the last node_element, therefore this loop will halt:
            while (XLCellReference(cellNode.attribute("r").value()).column() < columnNumber) cellNode = cellNode.next_sibling_of_type(pugi::node_element);
            // ===== If the forwards search failed to locate the requested cell
            if (XLCellReference(cellNode.attribute("r").value()).column() > columnNumber) {
                cellNode = rowNode.insert_child_before("c", cellNode);
                cellNode.append_attribute("r").set_value(cellRef.address().c_str());
            }
        }
        return cellNode;
    }

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLUTILITIES_HPP

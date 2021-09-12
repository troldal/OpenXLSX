//
// Created by Kenneth Balslev on 24/08/2020.
//

#ifndef OPENXLSX_XLUTILITIES_HPP
#define OPENXLSX_XLUTILITIES_HPP

#include <fstream>
#include <pugixml.hpp>

#include "XLCellReference.hpp"
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    /**
     * @details
     */
    inline XMLNode getRowNode(XMLNode sheetDataNode, uint32_t rowNumber)
    {
        // ===== If the requested node is beyond the current max node, append a new node to the end.
        auto result = XMLNode();
        if (!sheetDataNode.last_child() || rowNumber > sheetDataNode.last_child().attribute("r").as_ullong()) {
            result = sheetDataNode.append_child("row");

            result.append_attribute("r") = rowNumber;
            //            result.append_attribute("x14ac:dyDescent") = "0.2";
            //            result.append_attribute("spans")           = "1:1";
        }

        // ===== If the requested node is closest to the end, start from the end and search backwards
        else if (sheetDataNode.last_child().attribute("r").as_ullong() - rowNumber < rowNumber) {
            result = sheetDataNode.last_child();
            while (result.attribute("r").as_ullong() > rowNumber) result = result.previous_sibling();
            if (result.attribute("r").as_ullong() < rowNumber) {
                result = sheetDataNode.insert_child_after("row", result);

                result.append_attribute("r") = rowNumber;
                //                result.append_attribute("x14ac:dyDescent") = "0.2";
                //                result.append_attribute("spans")           = "1:1";
            }
        }

        // ===== Otherwise, start from the beginning
        else {
            result = sheetDataNode.first_child();
            while (result.attribute("r").as_ullong() < rowNumber) result = result.next_sibling();
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
     * @brief Retrieve the xml node representing the cell at the given row and column. If the node doesn't
     * exist, it will be created.
     * @param rowNode The row node under which to find the cell.
     * @param columnNumber The column at which to find the cell.
     * @return The xml node representing the requested cell.
     */
    inline XMLNode getCellNode(XMLNode rowNode, uint16_t columnNumber)
    {
        auto cellNode = XMLNode();
        auto cellRef  = XLCellReference(rowNode.attribute("r").as_uint(), columnNumber);

        // ===== If there are no cells in the current row, or the requested cell is beyond the last cell in the row...
        if (rowNode.last_child().empty() || XLCellReference(rowNode.last_child().attribute("r").value()).column() < columnNumber) {
            rowNode.append_child("c").append_attribute("r").set_value(cellRef.address().c_str());
            cellNode = rowNode.last_child();
        }
        // ===== If the requested node is closest to the end, start from the end and search backwards...
        else if (XLCellReference(rowNode.last_child().attribute("r").value()).column() - columnNumber < columnNumber) {
            cellNode = rowNode.last_child();
            while (XLCellReference(cellNode.attribute("r").value()).column() > columnNumber) cellNode = cellNode.previous_sibling();
            if (XLCellReference(cellNode.attribute("r").value()).column() < columnNumber) {
                cellNode = rowNode.insert_child_after("c", cellNode);
                cellNode.append_attribute("r").set_value(cellRef.address().c_str());
            }
        }
        // ===== Otherwise, start from the beginning
        else {
            cellNode = rowNode.first_child();
            while (XLCellReference(cellNode.attribute("r").value()).column() < columnNumber) cellNode = cellNode.next_sibling();
            if (XLCellReference(cellNode.attribute("r").value()).column() > columnNumber) {
                cellNode = rowNode.insert_child_before("c", cellNode);
                cellNode.append_attribute("r").set_value(cellRef.address().c_str());
            }
        }
        return cellNode;
    }

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLUTILITIES_HPP

/*

   ____                               ____      ___ ____       ____  ____      ___
  6MMMMb                              `MM(      )M' `MM'      6MMMMb\`MM(      )M'
 8P    Y8                              `MM.     d'   MM      6M'    ` `MM.     d'
6M      Mb __ ____     ____  ___  __    `MM.   d'    MM      MM        `MM.   d'
MM      MM `M6MMMMb   6MMMMb `MM 6MMb    `MM. d'     MM      YM.        `MM. d'
MM      MM  MM'  `Mb 6M'  `Mb MMM9 `Mb    `MMd       MM       YMMMMb     `MMd
MM      MM  MM    MM MM    MM MM'   MM     dMM.      MM           `Mb     dMM.
MM      MM  MM    MM MMMMMMMM MM    MM    d'`MM.     MM            MM    d'`MM.
YM      M9  MM    MM MM       MM    MM   d'  `MM.    MM            MM   d'  `MM.
 8b    d8   MM.  ,M9 YM    d9 MM    MM  d'    `MM.   MM    / L    ,M9  d'    `MM.
  YMMMM9    MMYMMM9   YMMMM9 _MM_  _MM_M(_    _)MM_ _MMMMMMM MYMMMM9 _M(_    _)MM_
            MM
            MM
           _MM_

  Copyright (c) 2018, Kenneth Troldal Balslev

  All rights reserved.

  Redistribution and use in source and binary forms, with or without
  modification, are permitted provided that the following conditions are met:
  - Redistributions of source code must retain the above copyright
    notice, this list of conditions and the following disclaimer.
  - Redistributions in binary form must reproduce the above copyright
    notice, this list of conditions and the following disclaimer in the
    documentation and/or other materials provided with the distribution.
  - Neither the name of the author nor the
    names of any contributors may be used to endorse or promote products
    derived from this software without specific prior written permission.

  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
  ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
  WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
  DISCLAIMED. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY
  DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
  (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
  LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
  ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
  (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
  SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

 */

// ===== Standard Library Includes ===== //
#include <algorithm>
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLCellRange.hpp"
#include "XLDocument.hpp"
#include "XLSheet.hpp"

using namespace OpenXLSX;

namespace
{
    /**
     * @brief
     * @param sheetDataNode
     * @param rowNumber
     * @return
     */
    inline XMLNode getRowNode(XMLNode sheetDataNode, uint32_t rowNumber)
    {
        // ===== If the requested node is beyond the current max node, append a new node to the end.
        auto result = XMLNode();
        if (rowNumber > sheetDataNode.last_child().attribute("r").as_ullong()) {
            result = sheetDataNode.append_child("row");

            result.append_attribute("r")               = rowNumber;
            result.append_attribute("x14ac:dyDescent") = "0.2";
            result.append_attribute("spans")           = "1:1";
        }

        // ===== If the requested node is closest to the end, start from the end and search backwards
        else if (sheetDataNode.last_child().attribute("r").as_ullong() - rowNumber < rowNumber) {
            result = sheetDataNode.last_child();
            while (result.attribute("r").as_ullong() > rowNumber) result = result.previous_sibling();
            if (result.attribute("r").as_ullong() < rowNumber) {
                result                                     = sheetDataNode.insert_child_after("row", result);

                result.append_attribute("r")               = rowNumber;
                result.append_attribute("x14ac:dyDescent") = "0.2";
                result.append_attribute("spans")           = "1:1";
            }
        }

        // ===== Otherwise, start from the beginning
        else {
            result = sheetDataNode.first_child();
            while (result.attribute("r").as_ullong() < rowNumber) result = result.next_sibling();
            if (result.attribute("r").as_ullong() > rowNumber) {
                result                                     = sheetDataNode.insert_child_before("row", result);

                result.append_attribute("r")               = rowNumber;
                result.append_attribute("x14ac:dyDescent") = "0.2";
                result.append_attribute("spans")           = "1:1";
            }
        }

        return result;
    }
}    // namespace

/**
 * @details The constructor initializes the member variables and calls the loadXMLData from the
 * XLAbstractXMLFile base class.
 */
XLWorksheet::XLWorksheet(XLXmlData* xmlData) : XLSheetBase(xmlData)
{
    // ===== Read the dimensions of the Sheet and set data members accordingly.
    std::string dimensions = xmlDocument().document_element().child("dimension").attribute("ref").value();
    if (dimensions.find(':') == std::string::npos)
        xmlDocument().document_element().child("dimension").set_value("A1");
    else
        xmlDocument().document_element().child("dimension").set_value(dimensions.substr(dimensions.find(':') + 1).c_str());

    // If Column properties are grouped, divide them into properties for individual Columns.
    if (xmlDocument().first_child().child("cols").type() != pugi::node_null) {
        auto currentNode = xmlDocument().first_child().child("cols").first_child();
        while (currentNode != nullptr) {
            int min = std::stoi(currentNode.attribute("min").value());
            int max = std::stoi(currentNode.attribute("max").value());
            if (min != max) {
                currentNode.attribute("min").set_value(max);
                for (int i = min; i < max; i++) {
                    auto newnode = xmlDocument().first_child().child("cols").insert_child_before("col", currentNode);
                    auto attr    = currentNode.first_attribute();
                    while (attr != nullptr) {
                        newnode.append_attribute(attr.name()) = attr.value();
                        attr                                  = attr.next_attribute();
                    }
                    newnode.attribute("min") = i;
                    newnode.attribute("max") = i;
                }
            }
            currentNode = currentNode.next_sibling();
        }
    }
}

/**
 * @details
 */
XLCell XLWorksheet::cell(const std::string& ref)
{
    return cell(XLCellReference(ref));
}

/**
 * @details
 */
XLCell XLWorksheet::cell(const std::string& ref) const
{
    return cell(XLCellReference(ref));
}

/**
 * @details
 */
XLCell XLWorksheet::cell(const XLCellReference& ref)
{
    return cell(ref.row(), ref.column());
}

/**
 * @details
 */
XLCell XLWorksheet::cell(const XLCellReference& ref) const
{
    return cell(ref.row(), ref.column());
}

/**
 * @details This function returns a pointer to an XLCell object in the worksheet. This particular overload
 * also serves as the main function, called by the other overloads.
 */
XLCell XLWorksheet::cell(uint32_t rowNumber, uint16_t columnNumber)
{
    auto cellNode = XMLNode();
    auto cellRef  = XLCellReference(rowNumber, columnNumber);
    auto rowNode  = getRowNode(xmlDocument().first_child().child("sheetData"), rowNumber);

    // ===== If there are no cells in the current row, or the requested cell is beyond the last cell in the row...
    if (rowNode.last_child().empty() || XLCellReference(rowNode.last_child().attribute("r").value()).column() < columnNumber) {
        // if (rowNode.last_child().empty() ||
        // XLCellReference::CoordinatesFromAddress(rowNode.last_child().attribute("r").value()).second < columnNumber) {
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

    return XLCell(cellNode, parentDoc().executeQuery(XLQuerySharedStrings()).sharedStrings());
}

/**
 * @details
 * @throw XLException if rowNumber exceeds rowCount
 */
XLCell XLWorksheet::cell(uint32_t rowNumber, uint16_t columnNumber) const
{
    //    if (rowNumber > RowCount())
    //        throw XLException("Row " + to_string(rowNumber) + " does not exist!");
    //
    //    return Row(rowNumber)->Cell(columnNumber);
}

/**
 * @brief
 * @return
 * @todo The returned object is not const.
 */
XLCellRange XLWorksheet::range() const
{
    return range(XLCellReference("A1"), lastCell());
}

/**
 * @details
 */
XLCellRange XLWorksheet::range(const XLCellReference& topLeft, const XLCellReference& bottomRight) const
{
    return XLCellRange(xmlDocument().first_child().child("sheetData"), topLeft, bottomRight);
}

/**
 * @details Get the XLRow object corresponding to the given row number. In the XML file, all cell data are stored under
 * the corresponding row, and all rows have to be ordered in ascending order. If a row have no data, there may not be a
 * node for that row. In RapidXLSX,the vector with all rows are initialized to a fixed size (the maximum size),
 * but most elements will be nullptr.
 */
XLRow XLWorksheet::row(uint32_t rowNumber)
{
    return XLRow(getRowNode(xmlDocument().first_child().child("sheetData"), rowNumber));
}

/**
 * @details
 */
const XLRow* XLWorksheet::row(uint32_t rowNumber) const
{
    return nullptr;
}

/**
 * @details Get the XLColumn object corresponding to the given column number. In the underlying XML data structure,
 * column nodes do not hold any cell data. Columns are used solely to hold data regarding column formatting.
 */
XLColumn XLWorksheet::column(uint16_t columnNumber) const
{
    // If no columns exists, create the <cols> node in the XML document.
    if (!xmlDocument().first_child().child("cols"))
        xmlDocument().root().insert_child_before("cols", xmlDocument().first_child().child("sheetData"));

    auto columnNode =
        xmlDocument().first_child().child("cols").find_child([&](XMLNode node) { return node.attribute("min").as_int() >= columnNumber; });

    if (!columnNumber || columnNode.attribute("min").as_int() > columnNumber) {
        if (columnNode.attribute("min").as_int() > columnNumber) {
            columnNode = xmlDocument().first_child().child("cols").insert_child_before("col", columnNode);
        }
        else {
            columnNode = xmlDocument().first_child().child("cols").append_child("col");
        }

        columnNode.append_attribute("min")         = columnNumber;
        columnNode.append_attribute("max")         = columnNumber;
        columnNode.append_attribute("width")       = 10;
        columnNode.append_attribute("customWidth") = 1;
    }

    return XLColumn(columnNode);
}

/**
 * @details Returns the value of the m_lastCell member variable.
 * @pre The m_lastCell member variable must have a valid value.
 * @post The object must remain unmodified.
 */
XLCellReference XLWorksheet::lastCell() const noexcept
{
    std::string dim = xmlDocument().document_element().child("dimension").value();
    return XLCellReference(dim.substr(dim.find(':') + 1));
}

/**
 * @details Returns the column() value from the last (bottom right) cell of the worksheet.
 * @pre LastCell() must return a valid XLCellReference object, which accurately represents the worksheet size.
 * @post Object must remain unmodified.
 */
uint16_t XLWorksheet::columnCount() const noexcept
{
    return lastCell().column();
}

/**
 * @details Returns the row() value from the last (bottom right) cell of the worksheet.
 * @pre LastCell() must return a valid XLCellReference object, which accurately represents the worksheet size.
 * @post Object must remain unmodified.
 */
uint32_t XLWorksheet::rowCount() const noexcept
{
    return xmlDocument().document_element().child("sheetData").last_child().attribute("r").as_ullong();
}

/**
 * @details
 */
void XLWorksheet::updateSheetName(const std::string& oldName, const std::string& newName)
{
    // ===== Set up temporary variables
    std::string oldNameTemp = oldName;
    std::string newNameTemp = newName;
    std::string formula;

    // ===== If the sheet name contains spaces, it should be enclosed in single quotes (')
    if (oldName.find(' ') != std::string::npos) oldNameTemp = "\'" + oldName + "\'";
    if (newName.find(' ') != std::string::npos) newNameTemp = "\'" + newName + "\'";

    // ===== Ensure only sheet names are replaced (references to sheets always ends with a '!')
    oldNameTemp += '!';
    newNameTemp += '!';

    // ===== Iterate through all defined names
    for (auto& row : xmlDocument().document_element().child("sheetData")) {
        for (auto& cell : row.children()) {
            if (!XLCell(cell, nullptr).hasFormula()) continue;

            formula = XLCell(cell, nullptr).formula();

            // ===== Skip if formula contains a '[' and ']' (means that the defined refers to external workbook)
            if (formula.find('[') == std::string::npos && formula.find(']') == std::string::npos) {
                // ===== For all instances of the old sheet name in the formula, replace with the new name.
                while (formula.find(oldNameTemp) != std::string::npos) {
                    formula.replace(formula.find(oldNameTemp), oldNameTemp.length(), newNameTemp);
                }
                XLCell(cell, nullptr).setFormula(formula);
            }
        }
    }
}

/**
 * @details
 */
XLChartsheet::XLChartsheet(XLXmlData* xmlData) : XLSheetBase(xmlData) {}

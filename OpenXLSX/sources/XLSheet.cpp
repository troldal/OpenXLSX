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

// ===== External Includes ===== //
#include <algorithm>
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLCellRange.hpp"
#include "XLDocument.hpp"
#include "XLRelationships.hpp"
#include "XLSheet.hpp"

using namespace OpenXLSX;

namespace OpenXLSX
{
    // Forward declaration. Implementation is in the XLUtilities.hpp file
    XMLNode getRowNode(XMLNode sheetDataNode, uint32_t rowNumber);

    /**
     * @brief Function for setting tab color.
     * @param xmlDocument XMLDocument object
     * @param color Thr color to set
     */
    void setTabColor(const XMLDocument& xmlDocument, const XLColor& color) {

        if (!xmlDocument.document_element().child("sheetPr")) xmlDocument.document_element().prepend_child("sheetPr");

        if (!xmlDocument.document_element().child("sheetPr").child("tabColor"))
            xmlDocument.document_element().child("sheetPr").prepend_child("tabColor");

        auto colorNode = xmlDocument.document_element().child("sheetPr").child("tabColor");
        for (auto attr : colorNode.attributes()) colorNode.remove_attribute(attr);

        colorNode.prepend_attribute("rgb").set_value(color.hex().c_str());
    }

    /**
     * @brief Function for checking if the tab is selected.
     * @param xmlDocument
     * @param selected
     */
    void setTabSelected(const XMLDocument& xmlDocument, bool selected) {
        unsigned int value = (selected ? 1 : 0);
        xmlDocument.first_child().child("sheetViews").first_child().attribute("tabSelected").set_value(value);
    }

    /**
     * @brief Set the tab selected property to true.
     * @param xmlDocument
     * @return
     */
    bool tabIsSelected(const XMLDocument& xmlDocument) {
        return xmlDocument.first_child().child("sheetViews").first_child().attribute("tabSelected").value();
    }

}    // namespace OpenXLSX

// ========== XLSheet Member Functions

/**
 * @details The constructor begins by constructing an instance of its superclass, XLAbstractXMLFile. The default
 * sheet type is WorkSheet and the default sheet state is Visible.
 */
XLSheet::XLSheet(XLXmlData* xmlData) : XLXmlFile(xmlData)
{
    if (xmlData->getXmlType() == XLContentType::Worksheet)
        m_sheet = XLWorksheet(xmlData);
    else if (xmlData->getXmlType() == XLContentType::Chartsheet)
        m_sheet = XLChartsheet(xmlData);
    else
        throw XLInternalError("Invalid XML data.");
}

/**
 * @details This method uses visitor pattern to return the name() member variable of the underlying
 * sheet object (XLWorksheet or XLChartsheet).
 */
std::string XLSheet::name() const
{
    return std::visit([](auto&& arg) { return arg.name(); }, m_sheet);
}

/**
 * @details This method sets the name of the sheet to a new name, by calling the setName()
 * member function of the underlying sheet object (XLWorksheet or XLChartsheet).
 */
void XLSheet::setName(const std::string& name)
{
    std::visit([&](auto&& arg) { return arg.setName(name); }, m_sheet);
}

/**
 * @details This method uses visitor pattern to return the visibility() member variable of the underlying
 * sheet object (XLWorksheet or XLChartsheet).
 */
XLSheetState XLSheet::visibility() const
{
    return std::visit([](auto&& arg) { return arg.visibility(); }, m_sheet);
}

/**
* @details This method sets the visibility state of the sheet, by calling the setVisibility()
* member function of the underlying sheet object (XLWorksheet or XLChartsheet).
 */
void XLSheet::setVisibility(XLSheetState state)
{
    std::visit([&](auto&& arg) { return arg.setVisibility(state); }, m_sheet);
}

/**
* @details This method uses visitor pattern to return the color() member variable of the underlying
* sheet object (XLWorksheet or XLChartsheet).
 */
XLColor XLSheet::color() const
{
    return std::visit([](auto&& arg) { return arg.color(); }, m_sheet);
}

/**
* @details This method sets the color of the sheet, by calling the setColor()
* member function of the underlying sheet object (XLWorksheet or XLChartsheet).
 */
void XLSheet::setColor(const XLColor& color)
{
    std::visit([&](auto&& arg) { return arg.setColor(color); }, m_sheet);
}

/**
* @details This method sets the selection state of the sheet, by calling the setSelected()
* member function of the underlying sheet object (XLWorksheet or XLChartsheet).
 */
void XLSheet::setSelected(bool selected)
{
    std::visit([&](auto&& arg) { return arg.setSelected(selected); }, m_sheet);
}

/**
 * @details Clones the sheet by calling the clone() method in the underlying sheet object
 * (XLWorksheet or XLChartsheet), using the visitor pattern.
 */
void XLSheet::clone(const std::string& newName)
{
    std::visit([&](auto&& arg) { arg.clone(newName); }, m_sheet);
}

/**
 * @details Get the index of the sheet, by calling the index() method in the underlying
 * sheet object (XLWorksheet or XLChartsheet), using the visitor pattern.
 */
uint16_t XLSheet::index() const
{
    return std::visit([](auto&& arg) { return arg.index(); }, m_sheet);
}

/**
* @details This method sets the index of the sheet (i.e. move the sheet), by calling the setIndex()
* member function of the underlying sheet object (XLWorksheet or XLChartsheet).
 */
void XLSheet::setIndex(uint16_t index)
{
    std::visit([&](auto&& arg) { arg.setIndex(index); }, m_sheet);
}

/**
 * @details Implicit conversion operator to XLWorksheet. Calls the get<>() member function, with XLWorksheet as
 * the template argument.
 */
XLSheet::operator XLWorksheet() const
{
    return this->get<XLWorksheet>();
}

/**
* @details Implicit conversion operator to XLChartsheet. Calls the get<>() member function, with XLChartsheet as
* the template argument.
 */
XLSheet::operator XLChartsheet() const
{
    return this->get<XLChartsheet>();
}


// ========== XLWorksheet Member Functions

/**
 * @details The constructor does some slight reconfiguration of the XML file, in order to make parsing easier.
 * For example, columns with identical formatting are by default grouped under the same node. However, this makes it more difficult to
 * parse, so the constructor reconfigures it so each column has it's own formatting.
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
                for (int i = min; i < max; i++) { // NOLINT
                    auto newnode = xmlDocument().first_child().child("cols").insert_child_before("col", currentNode);
                    auto attr    = currentNode.first_attribute();
                    while (attr != nullptr) { // NOLINT
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
 * @details Destructor. Default implementation.
 */
XLWorksheet::~XLWorksheet() = default;

/**
 * @details
 */
XLColor XLWorksheet::getColor_impl() const
{
    return XLColor(xmlDocument().document_element().child("sheetPr").child("tabColor").attribute("rgb").value());
}

/**
 * @details
 */
void XLWorksheet::setColor_impl(const XLColor& color)
{
    setTabColor(xmlDocument(), color);
}

/**
 * @details
 */
bool XLWorksheet::isSelected_impl() const
{
    return tabIsSelected(xmlDocument());
}

/**
 * @details
 */
void XLWorksheet::setSelected_impl(bool selected)
{
    setTabSelected(xmlDocument(), selected);
}

/**
 * @details
 */
bool XLWorksheet::isActive_impl() const
{
    return parentDoc().execQuery(XLQuery(XLQueryType::QuerySheetIsActive).setParam("sheetID", relationshipID())).result<bool>();
}

/**
 * @details
 */
void XLWorksheet::setActive_impl()
{
        parentDoc().execCommand(XLCommand(XLCommandType::SetSheetActive).setParam("sheetID", relationshipID()));
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
XLCell XLWorksheet::cell(const XLCellReference& ref) const
{
    return cell(ref.row(), ref.column());
}

/**
 * @details This function returns a pointer to an XLCell object in the worksheet. This particular overload
 * also serves as the main function, called by the other overloads.
 */
XLCell XLWorksheet::cell(uint32_t rowNumber, uint16_t columnNumber) const
{
    auto cellNode = XMLNode();
    auto cellRef  = XLCellReference(rowNumber, columnNumber);
    auto rowNode  = getRowNode(xmlDocument().first_child().child("sheetData"), rowNumber);

    // ===== If there are no cells in the current row, or the requested cell is beyond the last cell in the row...
    if (rowNode.last_child().empty() || XLCellReference(rowNode.last_child().attribute("r").value()).column() < columnNumber) {
        // if (rowNode.last_child().empty() ||
        // XLCellReference::CoordinatesFromAddress(rowNode.last_child().attribute("r").getValue()).second < columnNumber) {
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

    return XLCell{cellNode, parentDoc().execQuery(XLQuery(XLQueryType::QuerySharedStrings)).result<XLSharedStrings>()};
}

/**
 * @details
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
    return XLCellRange(xmlDocument().first_child().child("sheetData"),
                       topLeft,
                       bottomRight,
                       parentDoc().execQuery(XLQuery(XLQueryType::QuerySharedStrings)).result<XLSharedStrings>());
}

/**
 * @details
 * @pre
 * @post
 */
XLRowRange XLWorksheet::rows() const
{
    auto sheetDataNode = xmlDocument().first_child().child("sheetData");
    return XLRowRange(sheetDataNode,
                      1,
                      (sheetDataNode.last_child()
                           ? static_cast<uint32_t>(sheetDataNode.last_child().attribute("r").as_ullong())
                           : 1),
                      parentDoc().execQuery(XLQuery(XLQueryType::QuerySharedStrings)).result<XLSharedStrings>());
}

/**
 * @details
 * @pre
 * @post
 */
XLRowRange XLWorksheet::rows(uint32_t rowCount) const
{
    return XLRowRange(xmlDocument().first_child().child("sheetData"),
                      1,
                      rowCount,
                      parentDoc().execQuery(XLQuery(XLQueryType::QuerySharedStrings)).result<XLSharedStrings>());
}

/**
 * @details
 * @pre
 * @post
 */
XLRowRange XLWorksheet::rows(uint32_t firstRow, uint32_t lastRow) const
{
    return XLRowRange(xmlDocument().first_child().child("sheetData"),
                      firstRow,
                      lastRow,
                      parentDoc().execQuery(XLQuery(XLQueryType::QuerySharedStrings)).result<XLSharedStrings>());
}

/**
 * @details Get the XLRow object corresponding to the given row number. In the XML file, all cell data are stored under
 * the corresponding row, and all rows have to be ordered in ascending order. If a row have no data, there may not be a
 * node for that row.
 */
XLRow XLWorksheet::row(uint32_t rowNumber) const
{
    return XLRow{getRowNode(xmlDocument().first_child().child("sheetData"), rowNumber),
                   parentDoc().execQuery(XLQuery(XLQueryType::QuerySharedStrings)).result<XLSharedStrings>()};
}

/**
 * @details Get the XLColumn object corresponding to the given column number. In the underlying XML data structure,
 * column nodes do not hold any cell data. Columns are used solely to hold data regarding column formatting.
 * @todo Consider simplifying this function. Can any standard algorithms be used?
 */
XLColumn XLWorksheet::column(uint16_t columnNumber) const
{
    // If no columns exists, create the <cols> node in the XML document.
    if (!xmlDocument().first_child().child("cols"))
        xmlDocument().first_child().insert_child_before("cols", xmlDocument().first_child().child("sheetData"));

    // ===== Find the column node, if it exists
    auto columnNode =
        xmlDocument().first_child().child("cols").find_child([&](XMLNode node) {
            return (columnNumber >= node.attribute("min").as_int() && columnNumber <= node.attribute("max").as_int()) || node.attribute("min").as_int() > columnNumber; });

    // ===== If the node exists for the column, and only spans that column, then continue...
    if (columnNode && columnNode.attribute("min").as_int() == columnNumber && columnNode.attribute("max").as_int() == columnNumber) {}

    // ===== If the node exists for the column, but spans several columns, split it into individual nodes, and set columnNode to the right one...
    else if (columnNode && columnNode.attribute("min").as_int() != columnNode.attribute("max").as_int()) {
        // ===== Split the node in individual columns...
        for (int i = columnNode.attribute("min").as_int(); i < columnNode.attribute("max").as_int(); ++i) {
            auto node = xmlDocument().first_child().child("cols").insert_copy_before(columnNode, columnNode);
            node.attribute("min").set_value(i);
            node.attribute("max").set_value(i);
        }
        // ===== Delete the original node
        columnNode = columnNode.previous_sibling();
        xmlDocument().first_child().child("cols").remove_child(columnNode.next_sibling());
        // ===== Find the node corresponding to the column number
        while (true) {
            if (columnNode.attribute("min").as_int() == columnNumber) break;
            columnNode = columnNode.previous_sibling();
        }
    }

    // ===== If a node for the column does NOT exist, but a node for a higher column exist...
    else if (columnNode && columnNode.attribute("min").as_int() > columnNumber) {
        columnNode = xmlDocument().first_child().child("cols").insert_child_before("col", columnNode);
        columnNode.append_attribute("min")         = columnNumber;
        columnNode.append_attribute("max")         = columnNumber;
        columnNode.append_attribute("width")       = 9.8; // NOLINT
        columnNode.append_attribute("customWidth") = 0;
    }

    // ===== Otherwise, the end of the list is reached, and a new node is appended
    else if (!columnNode) {
        columnNode = xmlDocument().first_child().child("cols").append_child("col");
        columnNode.append_attribute("min")         = columnNumber;
        columnNode.append_attribute("max")         = columnNumber;
        columnNode.append_attribute("width")       = 9.8; // NOLINT
        columnNode.append_attribute("customWidth") = 0;
    }


//    if (!columnNode || columnNode.attribute("min").as_int() > columnNumber) {
//        if (columnNode.attribute("min").as_int() > columnNumber) {
//            columnNode = xmlDocument().first_child().child("cols").insert_child_before("col", columnNode);
//        }
//        else {
//            columnNode = xmlDocument().first_child().child("cols").append_child("col");
//        }
//
//        columnNode.append_attribute("min")         = columnNumber;
//        columnNode.append_attribute("max")         = columnNumber;
//        columnNode.append_attribute("width")       = 10; // NOLINT
//        columnNode.append_attribute("customWidth") = 1;
//    }

    return XLColumn(columnNode);
}

/**
 * @details Returns an XLCellReference to the last cell using rowCount() and columnCount() methods.
 */
XLCellReference XLWorksheet::lastCell() const noexcept
{
        return {rowCount(), columnCount()};
}

/**
 * @details Iterates through the rows and finds the maximum number of cells.
 */
uint16_t XLWorksheet::columnCount() const noexcept
{
        std::vector<int16_t> counts;
        for (const auto& row : rows()) {
            counts.emplace_back(row.cellCount());
        }

        return std::max(static_cast<uint16_t>(1), static_cast<uint16_t>(*std::max_element(counts.begin(), counts.end())));
}

/**
 * @details Finds the last row (node) and returns the row number.
 */
uint32_t XLWorksheet::rowCount() const noexcept
{
    return static_cast<uint32_t>(xmlDocument().document_element().child("sheetData").last_child().attribute("r").as_ullong());
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
            if (!XLCell(cell, XLSharedStrings()).hasFormula()) continue;

            formula = XLCell(cell, XLSharedStrings()).formula().get();

            // ===== Skip if formula contains a '[' and ']' (means that the defined refers to external workbook)
            if (formula.find('[') == std::string::npos && formula.find(']') == std::string::npos) {
                // ===== For all instances of the old sheet name in the formula, replace with the new name.
                while (formula.find(oldNameTemp) != std::string::npos) { // NOLINT
                    formula.replace(formula.find(oldNameTemp), oldNameTemp.length(), newNameTemp);
                }
                XLCell(cell, XLSharedStrings()).formula() = formula;
            }
        }
    }
}

/**
 * @details Constructor
 */
XLChartsheet::XLChartsheet(XLXmlData* xmlData) : XLSheetBase(xmlData) {}

/**
 * @details Destructor. Default implementation used.
 */
XLChartsheet::~XLChartsheet() = default;

/**
 * @details
 */
XLColor XLChartsheet::getColor_impl() const
{
    return XLColor(xmlDocument().document_element().child("sheetPr").child("tabColor").attribute("rgb").value());
}

/**
 * @details Calls the setTabColor() free function.
 */
void XLChartsheet::setColor_impl(const XLColor& color)
{
    setTabColor(xmlDocument(), color);
}

/**
 * @details Calls the tabIsSelected() free function.
 */
bool XLChartsheet::isSelected_impl() const
{
    return tabIsSelected(xmlDocument());
}

/**
 * @details Calls the setTabSelected() free function.
 */
void XLChartsheet::setSelected_impl(bool selected)
{
    setTabSelected(xmlDocument(), selected);
}
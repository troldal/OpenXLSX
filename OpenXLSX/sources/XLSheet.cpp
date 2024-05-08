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
#include "XLSheet.hpp"
#include "utilities/XLUtilities.hpp"

using namespace OpenXLSX;

namespace OpenXLSX
{
    // // Forward declaration. Implementation is in the XLUtilities.hpp file
    // XMLNode getRowNode(XMLNode sheetDataNode, uint32_t rowNumber);

    /**
     * @brief Function for setting tab color.
     * @param xmlDocument XMLDocument object
     * @param color Thr color to set
     */
    void setTabColor(const XMLDocument& xmlDocument, const XLColor& color)
    {
        if (!xmlDocument.document_element().child("sheetPr")) xmlDocument.document_element().prepend_child("sheetPr");

        if (!xmlDocument.document_element().child("sheetPr").child("tabColor"))
            xmlDocument.document_element().child("sheetPr").prepend_child("tabColor");

        auto colorNode = xmlDocument.document_element().child("sheetPr").child("tabColor");
        for (auto attr : colorNode.attributes()) colorNode.remove_attribute(attr);

        colorNode.prepend_attribute("rgb").set_value(color.hex().c_str());
    }

    /**
     * @brief Set the tab selected property to desired value
     * @param xmlDocument
     * @param selected
     */
    void setTabSelected(const XMLDocument& xmlDocument, bool selected)
    {    // 2024-04-30: whitespace support
        unsigned int value       = (selected ? 1 : 0);
        XMLNode      sheetView   = xmlDocument.document_element().child("sheetViews").first_child_of_type(pugi::node_element);
        XMLAttribute tabSelected = sheetView.attribute("tabSelected");
        if (tabSelected.empty())
            sheetView.prepend_attribute("tabSelected");    // BUGFIX 2024-05-01: create tabSelected attribute if it does not exist
        tabSelected.set_value(value);
    }

    /**
     * @brief Function for checking if the tab is selected.
     * @param xmlDocument
     * @return
     * @note pugi::xml_attribute::as_bool defaults to false for a non-existing (= empty) attribute
     */
    bool tabIsSelected(const XMLDocument& xmlDocument)
    {    // 2024-04-30: whitespace support
        return xmlDocument.document_element()
            .child("sheetViews")
            .first_child_of_type(pugi::node_element)
            .attribute("tabSelected")
            .as_bool();    // BUGFIX 2024-05-01: .value() "0" was evaluating to true
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
 * @details This reports the selection state of the sheet, by calling the isSelected()
 * member function of the underlying sheet object (XLWorksheet or XLChartsheet).
 */
bool XLSheet::isSelected() const
{
    return std::visit([](auto&& arg) { return arg.isSelected(); }, m_sheet);
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
XLSheet::operator XLWorksheet() const { return this->get<XLWorksheet>(); }

/**
 * @details Implicit conversion operator to XLChartsheet. Calls the get<>() member function, with XLChartsheet as
 * the template argument.
 */
XLSheet::operator XLChartsheet() const { return this->get<XLChartsheet>(); }

/**
 * @details Print the underlying XML using pugixml::xml_node::print
 */
void XLSheet::print(std::basic_ostream<char, std::char_traits<char>>& ostr) { xmlDocument().document_element().print(ostr); }

// ========== XLWorksheet Member Functions

/**
 * @details The constructor does some slight reconfiguration of the XML file, in order to make parsing easier.
 * For example, columns with identical formatting are by default grouped under the same node. However, this makes it more difficult to
 * parse, so the constructor reconfigures it so each column has it's own formatting.
 */
XLWorksheet::XLWorksheet(XLXmlData* xmlData) : XLSheetBase(xmlData)
{
    // ===== Read the dimensions of the Sheet and set data members accordingly.
    if (const std::string dimensions = xmlDocument().document_element().child("dimension").attribute("ref").value();
        dimensions.find(':') == std::string::npos)
        xmlDocument().document_element().child("dimension").set_value("A1");
    else
        xmlDocument().document_element().child("dimension").set_value(dimensions.substr(dimensions.find(':') + 1).c_str());

    // If Column properties are grouped, divide them into properties for individual Columns.
    if (xmlDocument().document_element().child("cols").type() != pugi::node_null) {
        auto currentNode = xmlDocument().document_element().child("cols").first_child_of_type(pugi::node_element);
        while (!currentNode.empty()) {
            uint16_t min {};
            uint16_t max {};
            try {
                min = std::stoi(currentNode.attribute("min").value());
                max = std::stoi(currentNode.attribute("max").value());
            }
            catch (...) {
                throw XLInternalError("Worksheet column min and/or max attributes are invalid.");
            }
            if (min != max) {
                currentNode.attribute("min").set_value(max);
                for (uint16_t i = min; i < max; i++) {    // NOLINT
                    auto newnode = xmlDocument().document_element().child("cols").insert_child_before("col", currentNode);
                    auto attr    = currentNode.first_attribute();
                    while (!attr.empty()) {    // NOLINT
                        newnode.append_attribute(attr.name()) = attr.value();
                        attr                                  = attr.next_attribute();
                    }
                    newnode.attribute("min") = i;
                    newnode.attribute("max") = i;
                }
            }
            currentNode = currentNode.next_sibling_of_type(pugi::node_element);
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
void XLWorksheet::setColor_impl(const XLColor& color) { setTabColor(xmlDocument(), color); }

/**
 * @details
 */
bool XLWorksheet::isSelected_impl() const { return tabIsSelected(xmlDocument()); }

/**
 * @details
 */
void XLWorksheet::setSelected_impl(bool selected) { setTabSelected(xmlDocument(), selected); }

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
bool XLWorksheet::setActive_impl()
{
    return parentDoc().execCommand(XLCommand(XLCommandType::SetSheetActive).setParam("sheetID", relationshipID()));
}

/**
 * @details
 */
XLCellAssignable XLWorksheet::cell(const std::string& ref) const
{
    return static_cast<XLCellAssignable>(
        cell(XLCellReference(ref)));    // TODO TBD if this is defined behavior: XLCellAssignable adds only methods, no member variables
}

/**
 * @details
 */
XLCell XLWorksheet::cell(const XLCellReference& ref) const { return cell(ref.row(), ref.column()); }

/**
 * @details This function returns a pointer to an XLCell object in the worksheet. This particular overload
 * also serves as the main function, called by the other overloads.
 */
XLCell XLWorksheet::cell(uint32_t rowNumber, uint16_t columnNumber) const
{
    const XMLNode rowNode  = getRowNode(xmlDocument().document_element().child("sheetData"), rowNumber);
    const XMLNode cellNode = getCellNode(rowNode, columnNumber, rowNumber);
    return XLCell { cellNode, parentDoc().execQuery(XLQuery(XLQueryType::QuerySharedStrings)).result<XLSharedStrings>() };

    /*** BEGIN obsolete section: below code is identical to XLUtilities.hpp::getCellNode ***/
    //     // ===== Get the last child of rowNode that is of type node_element, if any.
    //     auto cellNode = XMLNode();
    //     cellNode = rowNode.last_child_of_type(pugi::node_element);
    //     auto cellRef  = XLCellReference(rowNumber, columnNumber);  // only distinction from getCellNode: rowNumber is available directly
    //     as uint16_t
    //
    //     // ===== If there are no cells in the current row, or the requested cell is beyond the last cell in the row...
    //     if (cellNode.empty() || (XLCellReference(cellNode.attribute("r").value()).column() < columnNumber)) {
    //         // ===== append a new node to the end.
    //         rowNode.append_child("c").append_attribute("r").set_value(cellRef.address().c_str());
    //         cellNode = rowNode.last_child();
    //     }
    //     // ===== If the requested node is closest to the end, start from the end and search backwards...
    //     else if (XLCellReference(cellNode.attribute("r").value()).column() - columnNumber < columnNumber) {
    //         while (!cellNode.empty() && (XLCellReference(cellNode.attribute("r").value()).column() > columnNumber)) cellNode =
    //         cellNode.previous_sibling_of_type(pugi::node_element);
    //         // ===== If the backwards search failed to locate the requested cell
    //         if (cellNode.empty() || (XLCellReference(cellNode.attribute("r").value()).column() < columnNumber)) {
    //             if (cellNode.empty()) // If between row begin and higher column number, only non-element nodes exist
    //                 cellNode = rowNode.prepend_child("c"); // insert a new cell node at row begin. When saving, this will keep whitespace
    //                 formatting towards next cell node
    //             else
    //                 cellNode = rowNode.insert_child_after("c", cellNode);
    //             cellNode.append_attribute("r").set_value(cellRef.address().c_str());
    //         }
    //     }
    //     // ===== Otherwise, start from the beginning
    //     else {
    //         // ===== At this point, it is guaranteed that there is at least one node_element in the row that is not empty.
    //         cellNode = rowNode.first_child_of_type(pugi::node_element);
    //
    //         // ===== It has been verified above that the requested columnNumber is <= the column number of the last node_element,
    //         therefore this loop will halt: while (XLCellReference(cellNode.attribute("r").value()).column() < columnNumber) cellNode =
    //         cellNode.next_sibling_of_type(pugi::node_element);
    //         // ===== If the forwards search failed to locate the requested cell
    //         if (XLCellReference(cellNode.attribute("r").value()).column() > columnNumber) {
    //             cellNode = rowNode.insert_child_before("c", cellNode);
    //             cellNode.append_attribute("r").set_value(cellRef.address().c_str());
    //         }
    //     }
    //
    //     return XLCell{cellNode, parentDoc().execQuery(XLQuery(XLQueryType::QuerySharedStrings)).result<XLSharedStrings>()};
    /*** END obsolete section ***/
}

/**
 * @details
 */
XLCellRange XLWorksheet::range() const { return range(XLCellReference("A1"), lastCell()); }

/**
 * @details
 */
XLCellRange XLWorksheet::range(const XLCellReference& topLeft, const XLCellReference& bottomRight) const
{
    return XLCellRange(xmlDocument().document_element().child("sheetData"),
                       topLeft,
                       bottomRight,
                       parentDoc().execQuery(XLQuery(XLQueryType::QuerySharedStrings)).result<XLSharedStrings>());
}

/**
 * @details
 * @pre
 * @post
 */
XLRowRange XLWorksheet::rows() const    // 2024-04-29: patched for whitespace
{
    const auto sheetDataNode = xmlDocument().document_element().child("sheetData");
    return XLRowRange(sheetDataNode,
                      1,
                      (sheetDataNode.last_child_of_type(pugi::node_element).empty()
                           ? 1
                           : static_cast<uint32_t>(sheetDataNode.last_child_of_type(pugi::node_element).attribute("r").as_ullong())),
                      parentDoc().execQuery(XLQuery(XLQueryType::QuerySharedStrings)).result<XLSharedStrings>());
}

/**
 * @details
 * @pre
 * @post
 */
XLRowRange XLWorksheet::rows(uint32_t rowCount) const
{
    return XLRowRange(xmlDocument().document_element().child("sheetData"),
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
    return XLRowRange(xmlDocument().document_element().child("sheetData"),
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
    return XLRow { getRowNode(xmlDocument().document_element().child("sheetData"), rowNumber),
                   parentDoc().execQuery(XLQuery(XLQueryType::QuerySharedStrings)).result<XLSharedStrings>() };
}

/**
 * @details Get the XLColumn object corresponding to the given column number. In the underlying XML data structure,
 * column nodes do not hold any cell data. Columns are used solely to hold data regarding column formatting.
 * @todo Consider simplifying this function. Can any standard algorithms be used?
 */
XLColumn XLWorksheet::column(uint16_t columnNumber) const
{
    using namespace std::literals::string_literals;

    // If no columns exists, create the <cols> node in the XML document.
    if (xmlDocument().document_element().child("cols").empty())
        xmlDocument().document_element().insert_child_before("cols", xmlDocument().document_element().child("sheetData"));

    // ===== Find the column node, if it exists
    auto columnNode = xmlDocument().document_element().child("cols").find_child([&](const XMLNode node) {
        return (columnNumber >= node.attribute("min").as_int() && columnNumber <= node.attribute("max").as_int()) ||
               node.attribute("min").as_int() > columnNumber;
    });

    uint16_t minColumn {};
    uint16_t maxColumn {};
    if (!columnNode.empty()) {
        minColumn = columnNode.attribute("min").as_int();    // only look it up once for multiple access
        maxColumn = columnNode.attribute("max").as_int();    //   "
    }

    // ===== If the node exists for the column, and only spans that column, then continue...
    if (!columnNode.empty() && (minColumn == columnNumber) && (maxColumn == columnNumber)) {
    }

    // ===== If the node exists for the column, but spans several columns, split it into individual nodes, and set columnNode to the right
    // one... BUGFIX 2024-04-27 - old if condition would split a multi-column setting even if columnNumber is < minColumn (see lambda return
    // value above) NOTE 2024-04-27: the column splitting for loaded files is already handled in the constructor, technically this code is
    // not necessary here
    else if (!columnNode.empty() && (columnNumber >= minColumn) && (minColumn != maxColumn)) {
        // ===== Split the node in individual columns...
        columnNode.attribute("min").set_value(maxColumn);    // Limit the original node to a single column
        for (int i = minColumn; i < maxColumn; ++i) {
            auto node = xmlDocument().document_element().child("cols").insert_copy_before(columnNode, columnNode);
            node.attribute("min").set_value(i);
            node.attribute("max").set_value(i);
        }
        // BUGFIX 2024-04-27: the original node should not be deleted, but - in line with XLWorksheet constructor - is now limited above to
        // min = max attribute
        // // ===== Delete the original node
        // columnNode = columnNode.previous_sibling_of_type(pugi::node_element); // due to insert loop, previous node should be guaranteed
        // to be an element node xmlDocument().document_element().child("cols").remove_child(columnNode.next_sibling());

        // ===== Find the node corresponding to the column number - BUGFIX 2024-04-27: loop should abort on empty node
        while (not columnNode.empty() && columnNode.attribute("min").as_int() != columnNumber)
            columnNode = columnNode.previous_sibling_of_type(pugi::node_element);
        if (columnNode.empty())
            throw XLInternalError("XLWorksheet::"s + __func__ + ": column node for index "s + std::to_string(columnNumber) +
                                  "not found after splitting column nodes"s);
    }

    // ===== If a node for the column does NOT exist, but a node for a higher column exist...
    else if (!columnNode.empty() && minColumn > columnNumber) {
        columnNode                                 = xmlDocument().document_element().child("cols").insert_child_before("col", columnNode);
        columnNode.append_attribute("min")         = columnNumber;
        columnNode.append_attribute("max")         = columnNumber;
        columnNode.append_attribute("width")       = 9.8;    // NOLINT
        columnNode.append_attribute("customWidth") = 0;
    }

    // ===== Otherwise, the end of the list is reached, and a new node is appended
    else if (columnNode.empty()) {
        columnNode                                 = xmlDocument().document_element().child("cols").append_child("col");
        columnNode.append_attribute("min")         = columnNumber;
        columnNode.append_attribute("max")         = columnNumber;
        columnNode.append_attribute("width")       = 9.8;    // NOLINT
        columnNode.append_attribute("customWidth") = 0;
    }

    using namespace std::literals::string_literals;
    if (columnNode.empty())
        throw XLInternalError("XLWorksheet::"s + __func__ + ": was unable to find or create node for column "s +
                              std::to_string(columnNumber));
    return XLColumn(columnNode);
}

/**
 * @details Returns an XLCellReference to the last cell using rowCount() and columnCount() methods.
 */
XLCellReference XLWorksheet::lastCell() const noexcept { return { rowCount(), columnCount() }; }

/**
 * @details Iterates through the rows and finds the maximum number of cells.
 */
uint16_t XLWorksheet::columnCount() const noexcept
{
    uint16_t maxCount = 0;    // Pull request: Update XLSheet.cpp with correct type #176, Explicitely cast to unsigned short int #163
    for (const auto& row : rows()) {
        uint16_t cellCount = row.cellCount();
        maxCount = std::max(cellCount, maxCount);
    }
    return maxCount;

    // std::vector<uint16_t> counts; // Pull request: Update XLSheet.cpp with correct type #176, Explicitely cast to unsigned short int #163
    // for (const auto& row : rows()) {
    // counts.emplace_back(row.cellCount());
    // }
    // return std::max(static_cast<uint16_t>(1), *std::max_element(counts.begin(), counts.end()));
}

/**
 * @details Finds the last row (node) and returns the row number.
 */
uint32_t XLWorksheet::rowCount() const noexcept
{
    return static_cast<uint32_t>(
        xmlDocument().document_element().child("sheetData").last_child_of_type(pugi::node_element).attribute("r").as_ullong());
}

/**
 * @details
 */
void XLWorksheet::updateSheetName(const std::string& oldName, const std::string& newName)    // 2024-04-29: patched for whitespace
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
    XMLNode row = xmlDocument().document_element().child("sheetData").first_child_of_type(pugi::node_element);
    for (; not row.empty(); row = row.next_sibling_of_type(pugi::node_element)) {
        for (XMLNode cell = row.first_child_of_type(pugi::node_element); not cell.empty();
             cell         = cell.next_sibling_of_type(pugi::node_element))
        {
            if (!XLCell(cell, XLSharedStrings()).hasFormula()) continue;

            formula = XLCell(cell, XLSharedStrings()).formula().get();

            // ===== Skip if formula contains a '[' and ']' (means that the defined refers to external workbook)
            if (formula.find('[') == std::string::npos && formula.find(']') == std::string::npos) {
                // ===== For all instances of the old sheet name in the formula, replace with the new name.
                while (formula.find(oldNameTemp) != std::string::npos) {    // NOLINT
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
void XLChartsheet::setColor_impl(const XLColor& color) { setTabColor(xmlDocument(), color); }

/**
 * @details Calls the tabIsSelected() free function.
 */
bool XLChartsheet::isSelected_impl() const { return tabIsSelected(xmlDocument()); }

/**
 * @details Calls the setTabSelected() free function.
 */
void XLChartsheet::setSelected_impl(bool selected) { setTabSelected(xmlDocument(), selected); }

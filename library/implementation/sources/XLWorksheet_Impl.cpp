//
// Created by Troldal on 03/09/16.
//

#include "XLWorksheet_Impl.hpp"
#include "XLCellRange_Impl.hpp"
#include "XLSheet_Impl.hpp"
#include <algorithm>
#include <pugixml.hpp>

using namespace std;
using namespace OpenXLSX;

namespace {

    /**
     * @brief
     * @param sheetDataNode
     * @param rowNumber
     * @return
     */
    inline XMLNode getRowNode(XMLNode sheetDataNode, uint32_t rowNumber) {

        auto result = XMLNode();
        if (rowNumber > sheetDataNode.last_child().attribute("r").as_ullong()) {
            result = sheetDataNode.append_child("row");

            result.append_attribute("r") = rowNumber;
            result.append_attribute("x14ac:dyDescent") = 0.2;
            result.append_attribute("spans") = "1:1";

            //TODO: Use rowNode.append_copy instead of the above;
//        for (auto& attribute : rowNode.previous_sibling().attributes()) {
//            rowNode.append_copy(attribute);
//        }
        }

        // ===== If the requested node is closest to the end, start from the end and search backwards
        else if (sheetDataNode.last_child().attribute("r").as_ullong() - rowNumber < rowNumber ) {
            result = sheetDataNode.last_child();
            while (result.attribute("r").as_ullong() > rowNumber) result = result.previous_sibling();
            if (result.attribute("r").as_ullong() < rowNumber) {
                result = sheetDataNode.insert_child_after("row", result);
                result.append_attribute("r") = rowNumber;
                result.append_attribute("x14ac:dyDescent") = 0.2;
                result.append_attribute("spans") = "1:1";
            }
        }

        // ===== Otherwise, start from the beginning
        else {
            result = sheetDataNode.first_child();
            while (result.attribute("r").as_ullong() < rowNumber) result = result.next_sibling();
            if (result.attribute("r").as_ullong() > rowNumber) {
                result = sheetDataNode.insert_child_before("row", result);
                result.append_attribute("r") = rowNumber;
                result.append_attribute("x14ac:dyDescent") = 0.2;
                result.append_attribute("spans") = "1:1";
            }
        }

        return result;
    }

}  // namespace


/**
 * @details The constructor initializes the member variables and calls the loadXMLData from the
 * XLAbstractXMLFile base class.
 */
Impl::XLWorksheet::XLWorksheet(XLDocument& parent, const std::string& sheetRID, XMLAttribute name, const std::string& filePath,
                               const std::string& xmlData)

        : XLSheet(parent, sheetRID, name, filePath, xmlData) {

    // Call the 'LoadXMLData' method in the XLAbstractXMLFile base class
    ParseXMLData();
}

/**
 * @details This function reads the .xml file for the worksheet and populates the data into the internal
 * datastructure of the class. This function is not called directly, but is called via the loadXMLData
 * function in the base class.
 */
bool Impl::XLWorksheet::ParseXMLData() {

    // Read the dimensions of the Sheet and set data members accordingly.
    string dimensions = XmlDocument()->document_element().child("dimension").attribute("ref").value();
    if (dimensions.find(':') == string::npos)
        XmlDocument()->document_element().child("dimension").set_value("A1");
    else
        XmlDocument()->document_element().child("dimension").set_value(dimensions.substr(dimensions.find(':') + 1).c_str());

    // If Column properties are grouped, divide them into properties for individual Columns.
    if (XmlDocument()->first_child().child("cols").type() != pugi::node_null) {
        auto currentNode = XmlDocument()->first_child().child("cols").first_child();
        while (currentNode != nullptr) {
            int min = stoi(currentNode.attribute("min").value());
            int max = stoi(currentNode.attribute("max").value());
            if (min != max) {
                currentNode.attribute("min").set_value(max);
                for (int i = min; i < max; i++) {
                    auto newnode = XmlDocument()->first_child().child("cols").insert_child_before("col", currentNode);
                    auto attr = currentNode.first_attribute();
                    while (attr != nullptr) {
                        newnode.append_attribute(attr.name()) = attr.value();
                        attr = attr.next_attribute();
                    }
                    newnode.attribute("min") = i;
                    newnode.attribute("max") = i;
                }
            }
            currentNode = currentNode.next_sibling();
        }
    }

    return true;
}

/**
 * @details Creates an identical clone of the worksheet. All references internally in the spreadsheet are
 * handled automatically by the clone function.
 */
Impl::XLWorksheet* Impl::XLWorksheet::Clone(const std::string& newName) {

    ParentDoc().Workbook()->CloneWorksheet(Name(), newName);
    return ParentDoc().Workbook()->Worksheet(newName);
}

Impl::XLCell Impl::XLWorksheet::Cell(const XLCellReference& ref) {

    return Cell(ref.Row(), ref.Column());
}

/**
 * @details
 */
Impl::XLCell Impl::XLWorksheet::Cell(const XLCellReference& ref) const {

    return Cell(ref.Row(), ref.Column());
}

/**
 * @details This function returns a pointer to an XLCell object in the worksheet. This particular overload
 * also serves as the main function, called by the other overloads.
 */
Impl::XLCell Impl::XLWorksheet::Cell(uint32_t rowNumber,
                                      uint16_t columnNumber) {

    auto cellNode = XMLNode();
    auto cellRef  = XLCellReference(rowNumber, columnNumber);
    auto rowNode  = getRowNode(XmlDocument()->first_child().child("sheetData"), rowNumber);

        // ===== If there are no cells in the current row, or the requested cell is beyond the last cell in the row...
    if (rowNode.last_child().empty() || XLCellReference(rowNode.last_child().attribute("r").value()).Column() < columnNumber) {
    //if (rowNode.last_child().empty() || XLCellReference::CoordinatesFromAddress(rowNode.last_child().attribute("r").value()).second < columnNumber) {
        rowNode.append_child("c").append_attribute("r").set_value(cellRef.Address().c_str());
        cellNode = rowNode.last_child();
    }
        // ===== If the requested node is closest to the end, start from the end and search backwards...
    else if (XLCellReference(rowNode.last_child().attribute("r").value()).Column() - columnNumber < columnNumber) {
        cellNode = rowNode.last_child();
        while (XLCellReference(cellNode.attribute("r").value()).Column() > columnNumber)
            cellNode = cellNode.previous_sibling();
        if (XLCellReference(cellNode.attribute("r").value()).Column() < columnNumber) {
            cellNode = rowNode.insert_child_after("c", cellNode);
            cellNode.append_attribute("r").set_value(cellRef.Address().c_str());
        }
    }
        // ===== Otherwise, start from the beginning
    else {
        cellNode = rowNode.first_child();
        while (XLCellReference(cellNode.attribute("r").value()).Column() < columnNumber)
            cellNode = cellNode.next_sibling();
        if (XLCellReference(cellNode.attribute("r").value()).Column() > columnNumber) {
            cellNode = rowNode.insert_child_before("c", cellNode);
            cellNode.append_attribute("r").set_value(cellRef.Address().c_str());
        }
    }

//    m_currentCell->reset(cellNode);
//    return *m_currentCell;
    return XLCell(*this, cellNode);
}

/**
 * @details
 * @throw XLException if rowNumber exceeds rowCount
 */
 Impl::XLCell Impl::XLWorksheet::Cell(uint32_t rowNumber, uint16_t columnNumber) const {

//    if (rowNumber > RowCount())
//        throw XLException("Row " + to_string(rowNumber) + " does not exist!");
//
//    return Row(rowNumber)->Cell(columnNumber);

}

/**
 * @details Get an XLCellRange object, encompassing all valid cells in the worksheet. This is a convenience function
 * that calls a different function overload.
 */
Impl::XLCellRange Impl::XLWorksheet::Range() {

    return Range(XLCellReference("A1"), LastCell());
}

/**
 * @brief
 * @return
 * @todo The returned object is not const.
 */
const Impl::XLCellRange Impl::XLWorksheet::Range() const {

    return Range(XLCellReference("A1"), LastCell());
}

/**
 * @details Get an XLCellRange object with the given coordinates.
 */
Impl::XLCellRange Impl::XLWorksheet::Range(const XLCellReference& topLeft, const XLCellReference& bottomRight) {
    // Set the last Cell to some value, in order to create all objects in Range.
    if (Cell(bottomRight).ValueType() == XLValueType::Empty)
        Cell(bottomRight).Value().Clear();

    return XLCellRange(*this, topLeft, bottomRight);
}

/**
 * @details
 */
const Impl::XLCellRange Impl::XLWorksheet::Range(const XLCellReference& topLeft,
                                                 const XLCellReference& bottomRight) const {
    // Set the last Cell to some ValueAsString, in order to create all objects in Range.
    //if (Cell(bottomRight)->CellType() == XLCellType::Empty) Cell(bottomRight)->SetEmptyValue();

    return XLCellRange(*this, topLeft, bottomRight);
}

/**
 * @details Get the XLRow object corresponding to the given row number. In the XML file, all cell data are stored under
 * the corresponding row, and all rows have to be ordered in ascending order. If a row have no data, there may not be a
 * node for that row. In RapidXLSX,the vector with all rows are initialized to a fixed size (the maximum size),
 * but most elements will be nullptr.
 */
Impl::XLRow Impl::XLWorksheet::Row(uint32_t rowNumber) {

    return XLRow(*this, getRowNode(XmlDocument()->first_child().child("sheetData"), rowNumber));
}

/**
 * @details
 */
const Impl::XLRow* Impl::XLWorksheet::Row(uint32_t rowNumber) const {

    return nullptr;

}

/**
 * @details Get the XLColumn object corresponding to the given column number. In the underlying XML data structure,
 * column nodes do not hold any cell data. Columns are used solely to hold data regarding column formatting.
 */
Impl::XLColumn Impl::XLWorksheet::Column(uint16_t columnNumber) const {

    // If no columns exists, create the <cols> node in the XML document.
    if (!XmlDocument()->first_child().child("cols"))
        XmlDocument()->root().insert_child_before("cols", XmlDocument()->first_child().child("sheetData"));

    auto columnNode = XmlDocument()->first_child().child("cols").find_child([&](XMLNode node) {
        return node.attribute("min").as_int() >= columnNumber;
    });

    if (!columnNumber || columnNode.attribute("min").as_int() > columnNumber) {

        if (columnNode.attribute("min").as_int() > columnNumber) {
            columnNode = XmlDocument()->first_child().child("cols").insert_child_before("col", columnNode);
        }
        else {
            columnNode = XmlDocument()->first_child().child("cols").append_child("col");
        }

        columnNode.append_attribute("min") = columnNumber;
        columnNode.append_attribute("max") = columnNumber;
        columnNode.append_attribute("width") = 10;
        columnNode.append_attribute("customWidth") = 1;
    }

    return XLColumn(columnNode);
}

/**
 * @details Returns the value of the m_lastCell member variable.
 * @pre The m_lastCell member variable must have a valid value.
 * @post The object must remain unmodified.
 */
Impl::XLCellReference Impl::XLWorksheet::LastCell() const noexcept {

    std::string dim = XmlDocument()->document_element().child("dimension").value();
    return XLCellReference(dim.substr(dim.find(':') + 1));
}

/**
 * @details Returns the column() value from the last (bottom right) cell of the worksheet.
 * @pre LastCell() must return a valid XLCellReference object, which accurately represents the worksheet size.
 * @post Object must remain unmodified.
 */
uint16_t Impl::XLWorksheet::ColumnCount() const noexcept {

    return LastCell().Column();
}

/**
 * @details Returns the row() value from the last (bottom right) cell of the worksheet.
 * @pre LastCell() must return a valid XLCellReference object, which accurately represents the worksheet size.
 * @post Object must remain unmodified.
 */
uint32_t Impl::XLWorksheet::RowCount() const noexcept {

    return XmlDocument()->document_element().child("sheetData").last_child().attribute("r").as_ullong();
}

/**
 * @details
 */
void Impl::XLWorksheet::UpdateSheetName(const std::string& oldName, const std::string& newName) {

    // ===== Set up temporary variables
    std::string oldNameTemp = oldName;
    std::string newNameTemp = newName;
    std::string formula;

    // ===== If the sheet name contains spaces, it should be enclosed in single quotes (')
    if (oldName.find(' ') != std::string::npos)
        oldNameTemp = "\'" + oldName + "\'";
    if (newName.find(' ') != std::string::npos)
        newNameTemp = "\'" + newName + "\'";

    // ===== Ensure only sheet names are replaced (references to sheets always ends with a '!')
    oldNameTemp += '!';
    newNameTemp += '!';

    // ===== Iterate through all defined names
    for (auto& row : XmlDocument()->document_element().child("sheetData")) {
        for (auto& cell : row.children()) {
            if (!XLCell(*this, cell).HasFormula())
                continue;

            formula = XLCell(*this, cell).GetFormula();

            // ===== Skip if formula contains a '[' and ']' (means that the defined refers to external workbook)
            if (formula.find('[') == string::npos && formula.find(']') == string::npos) {

                // ===== For all instances of the old sheet name in the formula, replace with the new name.
                while (formula.find(oldNameTemp) != string::npos) {
                    formula.replace(formula.find(oldNameTemp), oldNameTemp.length(), newNameTemp);
                }
                XLCell(*this, cell).SetFormula(formula);
            }
        }
    }
}

std::string Impl::XLWorksheet::GetXmlData() const {

    return XLAbstractXMLFile::GetXmlData();

}

Impl::XLSheetType Impl::XLWorksheet::Type() const {
    return XLSheetType::WorkSheet;
}

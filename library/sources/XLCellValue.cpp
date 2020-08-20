//
// Created by Kenneth Balslev on 19/08/2020.
//

#include "XLCellValue.hpp"
#include "XLException.hpp"
#include <pugixml.hpp>

OpenXLSX::XLCellValue::XLCellValue(const XMLNode& cellNode, const XLSharedStrings& sharedStrings)
{
    // ===== If neither a Type attribute or a value node is present, the cell is empty.
    if (!cellNode.attribute("t") && !cellNode.child("v")) {
        m_type  = XLValueType::Empty;
        m_value = "";
    }

    // ===== If a Type attribute is not present, but a value node is, the cell contains a number.
    else if ((!cellNode.attribute("t") || (strcmp(cellNode.attribute("t").value(), "n") == 0 && cellNode.child("v") != nullptr))) {
        std::string numberString = cellNode.child("v").text().get();
        if (numberString.find('.') != std::string::npos || numberString.find("E-") != std::string::npos ||
            numberString.find("e-") != std::string::npos)
        {
            m_type  = XLValueType::Float;
            m_value = cellNode.child("v").text().as_double();
        }
        else {
            m_type  = XLValueType::Integer;
            m_value = cellNode.child("v").text().as_llong();
        }
    }

    // ===== If the cell is of type "s", the cell contains a shared string.
    else if (cellNode.attribute("t") != nullptr && strcmp(cellNode.attribute("t").value(), "s") == 0) {
        m_type  = XLValueType::String;
        m_value = sharedStrings.getString(static_cast<uint32_t>(cellNode.child("v").text().as_ullong()));
    }

    // ===== If the cell is of type "inlineStr", the cell contains an inline string.
    else if (cellNode.attribute("t") != nullptr && strcmp(cellNode.attribute("t").value(), "inlineStr") == 0) {
        m_type = XLValueType::String;
        throw XLException("Unknown string type");
    }

    // ===== If the cell is of type "str", the cell contains an ordinary string.
    else if (cellNode.attribute("t") != nullptr && strcmp(cellNode.attribute("t").value(), "str") == 0) {
        m_type  = XLValueType::String;
        m_value = cellNode.child("v").text().get();
    }

    // ===== If the cell is of type "b", the cell contains a boolean.
    else if (cellNode.attribute("t") != nullptr && strcmp(cellNode.attribute("t").value(), "b") == 0) {
        m_type  = XLValueType::Boolean;
        m_value = cellNode.child("v").text().as_bool();
    }

    // ===== Otherwise, the cell contains an error.
    else {
        m_type  = XLValueType::Error;    // the m_typeAttribute has the ValueAsString "e"
        m_value = "";
    }
}

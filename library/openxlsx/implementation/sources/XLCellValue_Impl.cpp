//
// Created by KBA012 on 16-12-2017.
//

#include <cassert>
#include <pugixml.hpp>

#include "XLCellValue_Impl.h"
#include "XLCell_Impl.h"
#include "XLWorksheet_Impl.h"

using namespace OpenXLSX;
using namespace std;

/**
 * @details The constructor sets the m_parentCell to the value of the input parameter. The m_parentCell member variable
 * is a reference to the parent XLCell object, in order to avoid null objects.
 * @pre The parent input parameter is a valid XLCell object.
 * @post A valid XLCellValue object has been constructed.
 */
Impl::XLCellValue::XLCellValue(XLCell& parent) noexcept
        : m_parentCell(parent) {
}

/**
 * @details The copy constructor and copy assignment operator works differently for XLCellValue objects.
 * The copy constructor creates an exact copy of the object, with the same parent XLCell object. The copy
 * assignment operator only copies the underlying cell value and type attribute to the target object.
 * @pre Both the lhs and rhs are valid objects.
 * @post Successful assignment to the target object.
 */
Impl::XLCellValue& Impl::XLCellValue::operator=(const Impl::XLCellValue& other) {

    if (&other != this) {
        SetValueNode(!other.ValueNode() ? "" : other.ValueNode().text().get());
        ValueNode().attribute("xml:space").set_value(other.ValueNode().attribute("xml:space").value());
        SetTypeAttribute(!other.TypeAttribute() ? "" : other.TypeAttribute().value());
    }

    return *this;
}

/**
 * @details The assignment operator taking a std::string object as a parameter, calls the corresponding Set method
 * and returns the resulting object.
 * @pre The XLCellValue object and the stringValue parameter are both valid.
 * @post The underlying cell xml data has been set to the value of the input parameter and the type attribute has been
 * set to 'str'
 */
Impl::XLCellValue& Impl::XLCellValue::operator=(const std::string& stringValue) {

    Set(stringValue);
    return *this;
}

/**
 * @details The assignment operator taking a char* string as a parameter, calls the corresponding Set method
 * and returns the resulting object.
 * @pre The XLCellValue object and the stringValue parameter are both valid.
 * @post The underlying cell xml data has been set to the value of the input parameter and the type attribute has been
 * set to 'str'
 */
Impl::XLCellValue& Impl::XLCellValue::operator=(const char* stringValue) {

    Set(stringValue);
    return *this;
}

/**
 * @details If the value type is already a String, the value will be set to the new value. Otherwise, the m_value
 * member variable will be set to an XLString object with the given value.
 * @pre The XLCellValue object and the stringValue parameter are both valid.
 * @post The underlying cell xml data has been set to the value of the input parameter and the type attribute has been
 * set to 'str'
 */
void Impl::XLCellValue::Set(const string& stringValue) {

    Set(stringValue.c_str());
}

/**
 * @details Converts the char* parameter to a std::string and calls the corresponding Set method.
 * @pre The XLCellValue object and the stringValue parameter are both valid.
 * @post The underlying cell xml data has been set to the value of the input parameter and the type attribute has been
 * set to 'str'
 */
void Impl::XLCellValue::Set(const char* stringValue) {

    TypeAttribute().set_value("str");
    ValueNode().text().set(stringValue);
    ValueNode().attribute("xml:space").set_value("preserve");

}

/**
 * @details Deletes the value node and type attribute if they exists.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post The value node and the type attribute in the underlying xml data has been deleted.
 */
void Impl::XLCellValue::Clear() {

    DeleteValueNode();
    DeleteTypeAttribute();
}

/**
 * @details Return the cell value as a string, by calling the AsString method of the m_value member variable.
 * @pre The m_value member variable is valid.
 * @post The current object, and any associated objects, are unchanged.
 */
std::string Impl::XLCellValue::AsString() const {

    if (string_view(TypeAttribute().value()) == "b") {
        if (string_view(ValueNode().text().get()) == "0")
            return "FALSE";

        return "TRUE";
    }

    if (string_view(TypeAttribute().value()) == "s")
        return string(
                Cell()->Worksheet()->Workbook()->SharedStrings()->GetStringNode(ValueNode().text().as_ullong()).text()
                      .get());

    return ValueNode().text().get();
}

/**
 * @details Determine the value type, based on the cell type, and return the corresponding XLValueType object.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post The current object, and any associated objects, are unchanged.
 */
XLValueType Impl::XLCellValue::ValueType() const {

    switch (CellType()) {
        case XLCellType::Empty:
            return XLValueType::Empty;

        case XLCellType::Error:
            return XLValueType::Error;

        case XLCellType::Number: {
            if (DetermineNumberType(ValueNode().text().get()) == XLNumberType::Integer)
                return XLValueType::Integer;

            return XLValueType::Float;
        }

        case XLCellType::Boolean:
            return XLValueType::Boolean;

        default:
            return XLValueType::String;
    }
}

/**
 * @details Determine the cell type, based on the contents of the underlying XML file, and returns the corresponding
 * XLCellType object.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post The current object, and any associated objects, are unchanged.
 */
Impl::XLCellType Impl::XLCellValue::CellType() const {

    // ===== If neither a Type attribute or a value node is present, the cell is empty.
    if (!m_parentCell.HasTypeAttribute() && !HasValueNode()) {
        return XLCellType::Empty;
    }

    // ===== If a Type attribute is not present, but a value node is, the cell contains a number.
    if ((!m_parentCell.HasTypeAttribute() || string(m_parentCell.TypeAttribute().value()) == "n") && HasValueNode()) {
        return XLCellType::Number;
    }

    // ===== If the cell is of type "s", the cell contains a shared string.
    if (m_parentCell.HasTypeAttribute() && string(m_parentCell.TypeAttribute().value()) == "s") {
        return XLCellType::String;
    }

    // ===== If the cell is of type "inlineStr", the cell contains an inline string.
    if (m_parentCell.HasTypeAttribute() && string(m_parentCell.TypeAttribute().value()) == "inlineStr") {
        return XLCellType::String;
    }

    // ===== If the cell is of type "str", the cell contains an ordinary string.
    if (m_parentCell.HasTypeAttribute() && string(m_parentCell.TypeAttribute().value()) == "str") {
        return XLCellType::String;
    }

    // ===== If the cell is of type "b", the cell contains a boolean.
    if (m_parentCell.HasTypeAttribute() && string(m_parentCell.TypeAttribute().value()) == "b") {
        return XLCellType::Boolean;
    }

    // ===== Otherwise, the cell contains an error.
    return XLCellType::Error; //the m_typeAttribute has the ValueAsString "e"
}

/**
 * @details Returns the value of the m_parentCell member variable.
 * @pre The parent XLCell object is valid.
 * @post The current object, and any associated objects, are unchanged.
 */
Impl::XLCell* Impl::XLCellValue::Cell() {

    return &m_parentCell;
}

/**
 * @details Returns the value of the m_parentCell member variable.
 * @pre The parent XLCell object is valid.
 * @post The current object, and any associated objects, are unchanged.
 */
const Impl::XLCell* Impl::XLCellValue::Cell() const {

    return &m_parentCell;
}

/**
 * @details If the input value is an empty string, object will be set to empty. Otherwise, a value node will be
 * created (if it doesn't exists) and the value set to the input value.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post The cell is either set to empty, or a value node exists and has been set to the input value.
 * @note If the input string is empty, the cell type will be set to empty.
 */
void Impl::XLCellValue::SetValueNode(string_view value) {

    if (value.empty())
        Clear();
    else
        ValueNode().text().set(string(value).c_str());
}

/**
 * @details If a value node exists, it will be deleted.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post No value node exists for the cell, in the underlying XML file.
 */
void Impl::XLCellValue::DeleteValueNode() {

    if (HasValueNode())
        ValueNode().parent().remove_child(ValueNode());
}

/**
 * @details If a value node doesn't exist, one will be created and returned. Otherwise, the existing one will be
 * returned.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post A value node exists for the cell in the underlying XML file.
 */
XMLNode Impl::XLCellValue::CreateValueNode() {

    if (!HasValueNode())
        Cell()->CreateValueNode();
    return ValueNode();
}

/**
 * @details If the type string is empty, the type attribute will be deleted. Otherwise, the type attribute will be
 * created and set to the value of the type string.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post The type attribute is either deleted or has been set to the value of the input type string.
 */
void Impl::XLCellValue::SetTypeAttribute(const std::string& typeString) {

    if (typeString.empty())
        DeleteTypeAttribute();
    else
        TypeAttribute().set_value(typeString.c_str());
}

/**
 * @details If a type attribute exists, it will be deleted from the XML file.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post No type attribute for the cell exists in the XML file.
 */
void Impl::XLCellValue::DeleteTypeAttribute() {

    if (HasTypeAttribute())
        Cell()->CellNode().remove_attribute("t");
}

/**
 * @details If a type attribute does not exist, a new one will be created, added to the XML document and returned.
 * Otherwise a pointer to the existing type attribute will be returned.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post An empty type attribute has been added to the XML file.
 * @note An empty type attribute may not be valid in Excel. Therefore, the value must be set immediately afterwards.
 */
XMLAttribute Impl::XLCellValue::CreateTypeAttribute() {

    if (!HasTypeAttribute())
        Cell()->CellNode().append_attribute("t") = "";

    return Cell()->CellNode().attribute("t");
}

/**
 * @details Obtain a pointer to the corresponding XMLNode object, requesting child node "v" from the cell node. If
 * no type node exists, a nullptr will be returned.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post The current object, and any associated objects, are unchanged.
 */
XMLNode Impl::XLCellValue::ValueNode() {

    if (!HasValueNode())
        CreateValueNode();
    return Cell()->CellNode().child("v");
}

/**
 * @details Obtain a pointer to the corresponding XMLNode object, requesting child node "v" from the cell node. If
 * no type node exists, a nullptr will be returned.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post The current object, and any associated objects, are unchanged.
 */
const XMLNode Impl::XLCellValue::ValueNode() const {

    return Cell()->CellNode().child("v");
}

/**
 * @details Determines the existance of the value node, by checking if a nullptr is returned.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post The current object, and any associated objects, are unchanged.
 */
bool Impl::XLCellValue::HasValueNode() const {

    return Cell()->CellNode().child("v") != nullptr;
}

/**
 * @details Obtain a pointer to the corresponding XMLAttribute object, requesting attribute "t" from the cell node. If
 * no type attribute exists, a nullptr will be returned.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post The current object, and any associated objects, are unchanged.
 */
XMLAttribute Impl::XLCellValue::TypeAttribute() {

    if (!HasTypeAttribute())
        CreateTypeAttribute();
    return Cell()->CellNode().attribute("t");
}

/**
 * @details Obtain a pointer to the corresponding XMLAttribute object, requesting attribute "t" from the cell node. If
 * no type attribute exists, a nullptr will be returned.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post The current object, and any associated objects, are unchanged.
 */
const XMLAttribute Impl::XLCellValue::TypeAttribute() const {

    return Cell()->CellNode().attribute("t");
}

/**
 * @details Determines the existance of the type attribute, by checking if a nullptr is returned.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post The current object, and any associated objects, are unchanged.
 */
bool Impl::XLCellValue::HasTypeAttribute() const {

    return Cell()->CellNode().attribute("t") != nullptr;
}

/**
 * @details The number type (integer or floating point) is determined simply by identifying whether or not a decimal
 * point is present in the input string. If present, the number type is floating point.
 */
Impl::XLNumberType Impl::XLCellValue::DetermineNumberType(const string& numberString) const {

    if (numberString.find('.') != string::npos)
        return XLNumberType::Float;

    return XLNumberType::Integer;
}

/**
 * @details Returns the shared string node at the given index, by accessing the shared strings via the parent workbook.
 */
XMLNode Impl::XLCellValue::SharedStringNode(unsigned long index) const {

    return Cell()->Worksheet()->Workbook()->SharedStrings()->GetStringNode(index);
}

void Impl::XLCellValue::SetInteger(long long int numberValue) {

    ValueNode().remove_attribute("t");
    ValueNode().text().set(numberValue);
    ValueNode().attribute("xml:space").set_value("default");
}

void Impl::XLCellValue::SetBoolean(bool numberValue) {

    TypeAttribute().set_value("b");
    ValueNode().text().set(numberValue ? 1 : 0);
    ValueNode().attribute("xml:space").set_value("default");
}

void Impl::XLCellValue::SetFloat(long double numberValue) {

    ValueNode().remove_attribute("t");
    ValueNode().text().set(std::to_string(numberValue).c_str());
    ValueNode().attribute("xml:space").set_value("default");
}

long long int Impl::XLCellValue::GetInteger() const {

    if (ValueType() != XLValueType::Integer)
        throw XLException("Cell value is not Integer");
    return ValueNode().text().as_llong();
}

bool Impl::XLCellValue::GetBoolean() const {

    if (ValueType() != XLValueType::Boolean)
        throw XLException("Cell value is not Boolean");
    return !(std::string_view(ValueNode().text().get()) == "0");
}

long double Impl::XLCellValue::GetFloat() const {

    if (ValueType() != XLValueType::Float)
        throw XLException("Cell value is not Float");
    return std::stold(ValueNode().text().get());
}

const char* Impl::XLCellValue::GetString() const {

    if (ValueType() != XLValueType::String)
        throw XLException("Cell value is not String");
    if (std::string_view(TypeAttribute().value()) == "str") // ordinary string
        return ValueNode().text().get();
    if (std::string_view(TypeAttribute().value()) == "s") // shared string
        return SharedStringNode(ValueNode().text().as_ullong()).text().get();

    throw XLException("Unknown string type");
}

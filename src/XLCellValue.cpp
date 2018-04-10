//
// Created by KBA012 on 16-12-2017.
//

#include "XLCellValue.h"
#include "XLException.h"
#include "XLWorksheet.h"

using namespace OpenXLSX;
using namespace std;


/**
 * @details The constructor sets the m_value to nullptr and the m_parentCell to the value of the input parameter.
 * Subsequently, the Initialize member function is called, which sets up the m_value variable.
 * @pre The parent input parameter is a valid XLCell object.
 * @post A valid XLCellValue object has been constructed.
 */
XLCellValue::XLCellValue(XLCell &parent)
    : m_value(nullptr),
      m_parentCell(parent)
{
    Initialize();
}

/**
 * @details Initializes the object, by creating the right value object, based on the cell type, which is ascertained
 * based on the content of the underlying XML file.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post The object has been initialized with the correct value type.
 */
void XLCellValue::Initialize()
{
    switch (CellType()) {
        case XLCellType::Empty:
            m_value = make_unique<XLValueEmpty>(*this);
            break;

        case XLCellType::Number:
            m_value = make_unique<XLValueNumber>(*this);
            break;

        case XLCellType::String:
            m_value = make_unique<XLValueString>(*this);
            break;

        case XLCellType::Boolean:
            m_value = make_unique<XLValueBoolean>(*this);
            break;

        case XLCellType::Error:
            m_value = make_unique<XLValueError>(*this);
            break;
    }
}

/**
 * @details The copy assignment operator sets the value node and type attribute to the corresponding items in
 * the object being assigned from. Subsequently, the Initialize method is called, which initializes the object
 * based on the underlying XML file.
 * @pre
 * @post
 */
XLCellValue &XLCellValue::operator=(const XLCellValue &other)
{
    SetValueNode(other.ValueNode() == nullptr ? "" : other.ValueNode()->Value());
    SetTypeAttribute(other.TypeString());

    Initialize();
    ParentCell()->SetModified();
    return *this;
}

/**
 * @details The copy assignment operator sets the value node and type attribute to the corresponding items in
 * the object being assigned from. Subsequently, the Initialize method is called, which initializes the object
 * based on the underlying XML file. This is identical to the copy assignment operator.
 * @pre
 * @post
 * @todo Implement a real move assignment operator
 */
XLCellValue &XLCellValue::operator=(XLCellValue &&other) noexcept
{
    SetValueNode(other.ValueNode() == nullptr ? "" : other.ValueNode()->Value());
    SetTypeAttribute(other.TypeString());

    Initialize();
    ParentCell()->SetModified();
    return *this;
}

/**
 * @details The assignment operator taking a XLBool object as a parameter, calls the corresponding Set method
 * and returns the resulting object.
 * @pre
 * @post
 */
XLCellValue &XLCellValue::operator=(XLBool boolValue)
{
    Set(boolValue);
    return *this;
}

/**
 * @details The assignment operator taking a std::tring object as a parameter, calls the corresponding Set method
 * and returns the resulting object.
 * @pre
 * @post
 */
XLCellValue &XLCellValue::operator=(const std::string &stringValue)
{
    Set(stringValue);
    return *this;
}

/**
 * @details The assignment operator taking a char* string as a parameter, calls the corresponding Set method
 * and returns the resulting object.
 * @pre
 * @post
 */
XLCellValue &XLCellValue::operator=(const char *stringValue)
{
    Set(stringValue);
    return *this;
}

/**
 * @details If the value type is already a boolean, the value will be set to the new value. Otherwise, the m_value
 * member variable will be set to an XLBoolean object with the given value.
 * @pre
 * @post
 */
void XLCellValue::Set(XLBool boolValue)
{
    if (ValueType() == XLValueType::Boolean)
        static_cast<XLValueBoolean &>(*m_value).Set(boolValue);
    else
        m_value = make_unique<XLValueBoolean>(boolValue, *this);

    ParentCell()->SetModified();
}

/**
 * @details If the value type is already a String, the value will be set to the new value. Otherwise, the m_value
 * member variable will be set to an XLString object with the given value.
 * @pre
 * @post
 */
void XLCellValue::Set(const std::string &stringValue)
{
    if (ValueType() == XLValueType::String)
        static_cast<XLValueString &>(*m_value).Set(stringValue);
    else
        m_value = make_unique<XLValueString>(stringValue, *this);

    ParentCell()->SetModified();
}

/**
 * @details Converts the char* parameter to a std::string and calls the corresponding Set method.
 * @pre
 * @post
 */
void XLCellValue::Set(const char *stringValue)
{
    Set(std::string(stringValue));
}

/**
 * @details Sets the cell value and type to empty and deletes the value node and type attribute if they exists.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post The m_value holds an XLEmpty object and the value node and type attributes have been deleted.
 */
void XLCellValue::SetEmpty()
{
    m_value = make_unique<XLValueEmpty>(*this);
    DeleteValueNode();
    DeleteTypeAttribute();
    ParentCell()->SetModified();
}

/**
 * @details Returns the boolean value of the cell, provided the value type is 'Boolean'.
 * @exception XLException If the value type is not 'Boolean', an XLException will be thrown.
 * @pre The m_value member variable is valid and the value type is 'Boolean'
 * @post The current object, and any associated objects, are unchanged.
 */
XLBool XLCellValue::Boolean() const
{
    if (ValueType() != XLValueType::Boolean) throw XLException("Cell value is not Boolean");
    return static_cast<XLValueBoolean &>(*m_value).Boolean();
}

/**
 * @details Returns the float value of the cell, provided the value type is 'Float'.
 * @exception XLException If the value type is not 'Float', an XLException will be thrown.
 * @pre The m_value member variable is valid and the value type is 'Float'
 * @post The current object, and any associated objects, are unchanged.
 */
long double XLCellValue::Float() const
{
    if (ValueType() != XLValueType::Float) throw XLException("Cell value is not Float");
    return static_cast<XLValueNumber &>(*m_value).Float();
}

/**
 * @details Returns the integer value of the cell, provided the value type is 'Integer'.
 * @exception XLException If the value type is not 'Integer', an XLException will be thrown.
 * @pre The m_value member variable is valid and the value type is 'Integer'
 * @post The current object, and any associated objects, are unchanged.
 */
long long int XLCellValue::Integer() const
{
    if (ValueType() != XLValueType::Integer) throw XLException("Cell value is not Integer");
    return static_cast<XLValueNumber &>(*m_value).Integer();
}

/**
 * @details Returns the string value of the cell, provided the value type is 'String'.
 * @exception XLException If the value type is not 'String', an XLException will be thrown.
 * @pre The m_value member variable is valid and the value type is 'String'
 * @post The current object, and any associated objects, are unchanged.
 */
const std::string &XLCellValue::String() const
{
    if (ValueType() != XLValueType::String) throw XLException("Cell value is not String");
    return static_cast<XLValueString &>(*m_value).String();
}

/**
 * @details Return the cell value as a string, by calling the AsString method of the m_value member variable.
 * @pre The m_value member variable is valid.
 * @post The current object, and any associated objects, are unchanged.
 */
std::string XLCellValue::AsString() const
{
    return (*m_value).AsString();
}

/**
 * @details Determine the value type, based on the cell type, and return the corresponding XLValueType object.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post The current object, and any associated objects, are unchanged.
 */
XLValueType XLCellValue::ValueType() const
{
    switch (CellType()) {
        case XLCellType::Empty:
            return XLValueType::Empty;

        case XLCellType::Error:
            return XLValueType::Error;

        case XLCellType::Number: {
            if (dynamic_cast<XLValueNumber &>(*m_value).NumberType() == XLNumberType::Integer)
                return XLValueType::Integer;
            else
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
XLCellType XLCellValue::CellType() const
{
    // Determine the type of the Cell
    // If neither a Type attribute or a value node is present, the cell is empty.
    if (!m_parentCell.HasTypeAttribute() && !HasValueNode()) {
        return XLCellType::Empty;
    }

        // If a Type attribute is not present, but a value node is, the cell contains a number.
    else if (!m_parentCell.HasTypeAttribute() && HasValueNode()) {
        return XLCellType::Number;
    }

        // If the cell is of type "s", the cell contains a shared string.
    else if (m_parentCell.HasTypeAttribute() && m_parentCell.TypeAttribute()->Value() == "s") {
        return XLCellType::String;
    }

        // If the cell is of type "inlineStr", the cell contains an inline string.
    else if (m_parentCell.HasTypeAttribute() && m_parentCell.TypeAttribute()->Value() == "inlineStr") {
        return XLCellType::String;
    }

        // If the cell is of type "str", the cell contains an ordinary string.
    else if (m_parentCell.HasTypeAttribute() && m_parentCell.TypeAttribute()->Value() == "str") {
        return XLCellType::String;
    }

        // If the cell is of type "b", the cell contains a boolean.
    else if (m_parentCell.HasTypeAttribute() && m_parentCell.TypeAttribute()->Value() == "b") {
        return XLCellType::Boolean;
    }

        // Otherwise, the cell contains an error.
    else
        return XLCellType::Error; //the m_typeAttribute has the ValueAsString "e"
}

/**
 * @details Returns the type string by calling the TypeString method of the m_value member variable.
 * @pre The m_value member variable is valid.
 * @post The current object, and any associated objects, are unchanged.
 */
std::string XLCellValue::TypeString() const
{
    return (*m_value).TypeString();
}

/**
 * @details Returns the value of the m_parentCell member variable.
 * @pre The parent XLCell object is valid.
 * @post The current object, and any associated objects, are unchanged.
 */
XLCell *XLCellValue::ParentCell()
{
    return &m_parentCell;
}

/**
 * @details Returns the value of the m_parentCell member variable.
 * @pre The parent XLCell object is valid.
 * @post The current object, and any associated objects, are unchanged.
 */
const XLCell *XLCellValue::ParentCell() const
{
    return &m_parentCell;
}

/**
 * @details If the input value is an empty string, object will be set to empty. Otherwise, a value node will be
 * created (if it doesn't exists) and the value set to the input value.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post The cell is either set to empty, or a value node exists and has been set to the input value.
 * @note If the input string is empty, the cell type will be set to empty.
 */
void XLCellValue::SetValueNode(const std::string &value)
{
    if (value.empty()) {
        SetEmpty();
    }
    else {
        CreateValueNode();
        ValueNode()->SetValue(value);
    }
}

/**
 * @details If a value node exists, it will be deleted.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post No value node exists for the cell, in the underlying XML file.
 */
void XLCellValue::DeleteValueNode()
{
    if (HasValueNode()) ValueNode()->DeleteNode();
}

/**
 * @details If a value node doesn't exist, one will be created and returned. Otherwise, the existing one will be
 * returned.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post A value node exists for the cell in the underlying XML file.
 */
XMLNode *XLCellValue::CreateValueNode()
{
    if (!HasValueNode()) ParentCell()->CreateValueNode();
    return ValueNode();
}

/**
 * @details If the type string is empty, the type attribute will be deleted. Otherwise, the type attribute will be
 * created and set to the value of the type string.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post The type attribute is either deleted or has been set to the value of the input type string.
 */
void XLCellValue::SetTypeAttribute(const std::string &typeString)
{
    if (typeString.empty())
        DeleteTypeAttribute();
    else {
        CreateTypeAttribute();
        ParentCell()->CellNode()->Attribute("t")->SetValue(typeString);
    }
}

/**
 * @details If a type attribute exists, it will be deleted from the XML file.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post No type attribute for the cell exists in the XML file.
 */
void XLCellValue::DeleteTypeAttribute()
{
    if (HasTypeAttribute()) TypeAttribute()->DeleteAttribute();
}

/**
 * @details If a type attribute does not exist, a new one will be created, added to the XML document and returned.
 * Otherwise a pointer to the existing type attribute will be returned.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post An empty type attribute has been added to the XML file.
 * @note An empty type attribute may not be valid in Excel. Therefore, the value must be set immediately afterwards.
 */
XMLAttribute *XLCellValue::CreateTypeAttribute()
{
    if (!HasTypeAttribute())
        ParentCell()->CellNode()->AppendAttribute(ParentCell()->XmlDocument()->CreateAttribute("t", ""));

    return ParentCell()->CellNode()->Attribute("t");
}

/**
 * @details Obtain a pointer to the corresponding XMLNode object, requesting child node "v" from the cell node. If
 * no type node exists, a nullptr will be returned.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post The current object, and any associated objects, are unchanged.
 */
XMLNode *XLCellValue::ValueNode()
{
    return ParentCell()->CellNode()->ChildNode("v");
}

/**
 * @details Obtain a pointer to the corresponding XMLNode object, requesting child node "v" from the cell node. If
 * no type node exists, a nullptr will be returned.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post The current object, and any associated objects, are unchanged.
 */
const XMLNode *XLCellValue::ValueNode() const
{
    return ParentCell()->CellNode()->ChildNode("v");
}

/**
 * @details Determines the existance of the value node, by checking if a nullptr is returned.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post The current object, and any associated objects, are unchanged.
 */
bool XLCellValue::HasValueNode() const
{
    if (ParentCell()->CellNode()->ChildNode("v") == nullptr)
        return false;
    else
        return true;
}

/**
 * @details Obtain a pointer to the corresponding XMLAttribute object, requesting attribute "t" from the cell node. If
 * no type attribute exists, a nullptr will be returned.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post The current object, and any associated objects, are unchanged.
 */
XMLAttribute *XLCellValue::TypeAttribute()
{
    return ParentCell()->CellNode()->Attribute("t");
}

/**
 * @details Obtain a pointer to the corresponding XMLAttribute object, requesting attribute "t" from the cell node. If
 * no type attribute exists, a nullptr will be returned.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post The current object, and any associated objects, are unchanged.
 */
const XMLAttribute *XLCellValue::TypeAttribute() const
{
    return ParentCell()->CellNode()->Attribute("t");
}

/**
 * @details Determines the existance of the type attribute, by checking if a nullptr is returned.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post The current object, and any associated objects, are unchanged.
 */
bool XLCellValue::HasTypeAttribute() const
{
    if (ParentCell()->CellNode()->Attribute("t") == nullptr)
        return false;
    else
        return true;
}


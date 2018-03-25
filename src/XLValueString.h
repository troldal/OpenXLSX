//
// Created by KBA012 on 18-12-2017.
//

#ifndef OPENXL_XLSTRING_H
#define OPENXL_XLSTRING_H

#include <string>
#include "XLCellType.h"
#include "XLValue.h"

namespace RapidXLSX
{
    class XLCell;
    class XMLNode;

//======================================================================================================================
//========== XLValueString Class ============================================================================================
//======================================================================================================================

    /**
     * @brief This class represents the different types of string that are used in Excel. For the user, it behaves the
     * same, independent of string type.
     */
    class XLValueString: public XLValue
    {

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief Constructor
         * @param parent A reference to the parent XLCellValue object.
         */
        explicit XLValueString(XLCellValue &parent);

        /**
         * @brief Constructor
         * @param stringValue A std::string with the string value to initialize the object with.
         * @param parent A reference to the parent XLCellValue object.
         */
        explicit XLValueString(const std::string &stringValue, XLCellValue &parent);

        /**
         * @brief Constructor
         * @param stringIndex The index of the relevant shared string.
         * @param parent A reference to the parent XLCellValue object.
         */
        explicit XLValueString(unsigned long stringIndex, XLCellValue &parent);

        /**
         * @brief Copy constructor
         * @param other The object to be copied.
         * @todo A copy constructor does not really make sense, as an XLCellValueType object will belong to a specific
         * XLCell object. However, a copy constructor is required in order to use the std::variant data structure. A
         * solution should be found, that does not require a copy constructor.
         */
        XLValueString(const XLValueString &other) = default;

        /**
         * @brief Move constructor
         * @param other The object to be moved.
         * @note The move constructor has been explicitly deleted.
         */
        XLValueString(XLValueString &&other) = delete;

        /**
         * @brief Destructor
         * @note The default implementation has been used.
         */
        ~XLValueString() override = default;

        /**
         * @brief Copy assignment operator
         * @param other The object to be copied.
         * @return A reference to the object with the new value.
         */
        XLValueString &operator=(const XLValueString &other);

        /**
         * @brief Move assignment operator
         * @param other the object to be moved
         * @return A reference to the object with the new value.
         */
        XLValueString &operator=(XLValueString &&other) noexcept;

        /**
         * @brief Assignment operator
         * @param str A std::string to be assigned
         * @return A reference to the object with the new value.
         */
        XLValueString &operator=(const std::string &str);

        /**
         * @brief Creates a polymorphic clone of the object.
         * @param parent The parent XLCell object of the clone.
         * @return A unique_ptr with the clone.
         */
        std::unique_ptr<XLValue> Clone(XLCell &parent) override;

        /**
         * @brief Set the string value
         * @param stringValue A std::string with the new value
         * @param type The type of the string; the default is an ordinary string.
         */
        void Set(const std::string &stringValue, XLStringType type = XLStringType::String);

        /**
         * @brief Get the string value as a std::string.
         * @return A std::string with the string value.
         */
        const std::string &String() const;

        /**
         * @brief Get the value as a std::string, regardless of the type. For all string types, it is identical
         * to the String method, and defined only to comply with the XLCellValueType interface.
         * @return A std::string with the string value.
         */
        std::string AsString() const override;

        /**
         * @brief Get the string type of the string.
         * @return An XLStringType with the type of the current string.
         */
        XLStringType StringType() const;

        /**
         * @brief Get the value type of the object.
         * @return An XLValueType object with the correct value type for the object.
         */
        XLValueType ValueType() const override;

        /**
         * @brief Get the cell type of the object.
         * @return An XLCellType object with the correct cell type for the object.
         */
        XLCellType CellType() const override;

        /**
         * @brief Get the type string for the object.
         * @return A std::string with a string corresponding to the type attribute for the object.
         */
        std::string TypeString() const override;

        /**
         * @brief Get the index of the string, if it is of type SharedString.
         * @return A long int with the index of the shared string.
         * @exception Throws an XLException if the type is not SharedString.
         */
        long SharedStringIndex() const;

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Functions
//----------------------------------------------------------------------------------------------------------------------

    private:

        /**
         * @brief Private function used to initialize a string object.
         * @param stringValue A std::string to initialize the object with.
         */
        void Initialize(const std::string &stringValue);



//----------------------------------------------------------------------------------------------------------------------
//           Private Member Variables
//----------------------------------------------------------------------------------------------------------------------

    private:
        XLStringType m_type; /**< The type of the string, i.e. String (ordinary), SharedString or InlineString */

    };

}


#endif //OPENXLEXE_XLSTRING_H

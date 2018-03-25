//
// Created by Kenneth Balslev on 31/12/2017.
//

#ifndef OPENXLEXE_XLCELLVALUETYPE_H
#define OPENXLEXE_XLCELLVALUETYPE_H

#include <string>
#include <memory>
#include "XLCellType.h"

namespace RapidXLSX
{

    class XMLNode;
    class XLCell;
    class XLCellValue;

//======================================================================================================================
//========== XLValue Class =====================================================================================
//======================================================================================================================

    /**
     * @brief The XLCellValueType class is a pure virtual class, functioning as parent class to any class representing
     * an Excel value type.
     */
    class XLValue
    {

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief Constructor
         * @param parent A reference to the parent XLCellValue object.
         */
        explicit XLValue(XLCellValue &parent);

        /**
         * @brief Copy constructor
         * @param other The object to be copied.
         * @todo A copy constructor does not really make sense, as an XLCellValueType object will belong to a specific
         * XLCell object. However, a copy constructor is required in order to use the std::variant data structure. A
         * solution should be found, that does not require a copy constructor.
         */
        XLValue(const XLValue &other) = default;

        /**
         * @brief Move construcor
         * @param other The object to be move constructed.
         * @note This has been explicitly deleted.
         */
        XLValue(XLValue &&other) = delete;

        /**
         * @brief Destructor
         * @note Default implementation has been used.
         */
        virtual ~XLValue() = default;

        /**
         * @brief Copy assignment operator
         * @param other The object to be assigned from.
         * @return A reference to the current object with the new value-
         */
        XLValue &operator=(const XLValue &other);

        /**
         * @brief Move assignment operator.
         * @param other The object to be moved.
         * @return A reference to the object after the move.
         */
        XLValue &operator=(XLValue &&other) noexcept;

        /**
         * @brief Creates a polymorphic clone of the object
         * @param parent A reference to the parent XLCell object of the clone
         * @return A unique_ptr holding the clone.
         */
        virtual std::unique_ptr<XLValue> Clone(XLCell &parent) = 0;

        /**
         * @brief Get the value type of the object.
         * @return Returns an XLValueType object with the corresponding type.
         */
        virtual XLValueType ValueType() const = 0;

        /**
         * @brief Get the cell type of the object.
         * @return Return an XLCellType object with the corresponding type.
         */
        virtual XLCellType CellType() const = 0;

        /**
         * @brief Produce a string corrsponding to the type attribute in the underlying XML file.
         * @return A string with the type attribute to e used.
         */
        virtual std::string TypeString() const = 0;

        /**
         * @brief Produce a string with the current value, regardless of value type.
         * @return A string with the value as a string.
         */
        virtual std::string AsString() const = 0;

//----------------------------------------------------------------------------------------------------------------------
//           Protected Member Functions
//----------------------------------------------------------------------------------------------------------------------

    protected:

        /**
         * @brief Get the parent XLCellValue object.
         * @return A reference to the parent XLCellValue object.
         */
        XLCellValue &ParentCellValue();

        /**
         * @brief Get the parent XLCellValue object.
         * @return A const reference to the parent XLCellValue object.
         */
        const XLCellValue &ParentCellValue() const;

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Variables
//----------------------------------------------------------------------------------------------------------------------

    private:
        XLCellValue &m_parentCellValue; /**<  */

    };

}


#endif //OPENXLEXE_XLCELLVALUETYPE_H

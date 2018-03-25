//
// Created by Kenneth Balslev on 04/01/2018.
//

#ifndef OPENXLEXE_XLEMPTY_H
#define OPENXLEXE_XLEMPTY_H

#include "XLValue.h"
namespace RapidXLSX
{

//======================================================================================================================
//========== XLValueEmpty Class =============================================================================================
//======================================================================================================================

    /**
     * @brief This class encapsulates an 'empty' value, i.e. a cell without a value or type.
     */
    class XLValueEmpty: public XLValue
    {

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief Constructor
         * @param parent A reference to the parent XLCellValue object.
         */
        explicit XLValueEmpty(XLCellValue &parent);

        /**
         * @brief Copy constructor
         * @param other The XLEmpty object to be copied.
         * @note The default implementation has been used.
         */
        XLValueEmpty(const XLValueEmpty &other) = default;

        /**
         * @brief Move constructor
         * @param other the XLEmpty object to move.
         * @note The move constructor has been explicitly deleted.
         */
        XLValueEmpty(XLValueEmpty &&other) = delete;

        /**
         * @brief Destructor
         * @note The default implementation has been used.
         */
        ~XLValueEmpty() override = default;

        /**
         * @brief Copy assignment operator.
         * @param other The XLEmpty object to copy.
         * @return A reference to the object with the new value.
         * @note The default implementation has been used.
         */
        XLValueEmpty &operator=(const XLValueEmpty &other) = default;

        /**
         * @brief Move assignment operator.
         * @param other The XLEmpty object to move.
         * @return A reference to the object with the moved value.
         * @note The default implementation has been used.
         */
        XLValueEmpty &operator=(XLValueEmpty &&other) noexcept = default;

        /**
         * @brief Creates a polymorphic clone of the object.
         * @param parent The parent XLCell object of the clone.
         * @return A unique_ptr with the clone.
         */
        std::unique_ptr<XLValue> Clone(XLCell &parent) override;

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
         * @brief Get the value as a std::string, regardless of the type.
         * @return A std::string with the string value.
         */
        std::string AsString() const override;

    };

}


#endif //OPENXLEXE_XLEMPTY_H

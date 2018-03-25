//
// Created by Kenneth Balslev on 31/12/2017.
//

#ifndef OPENXLEXE_XLBOOLEAN_H
#define OPENXLEXE_XLBOOLEAN_H

#include "XLValue.h"
#include "XLCellType.h"

namespace RapidXLSX
{

//======================================================================================================================
//========== XLValueString Class ============================================================================================
//======================================================================================================================

    /**
     * @brief The XLBoolean class encapsulates the concept of a boolean value. The reason to use this rather than the
     * built-in type bool, is to enable polymorphic exchange of types and to avoid implicit conversion between e.g. int
     * and bool.
     */
    class XLValueBoolean: public XLValue
    {

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief Constructor
         * @param parent A reference to the parent XLCellValue object.
         */
        explicit XLValueBoolean(XLCellValue &parent);

        /**
         * @brief Constructor
         * @param boolValue An XLBool object with the value used to initialize the object.
         * @param parent A reference to the parent XLCellValue object.
         */
        XLValueBoolean(XLBool boolValue, XLCellValue &parent);

        /**
         * @brief Constructor
         * @param boolValue An int with the value used to initialize the object (0 for false; otherwise true).
         * @param parent A reference to the parent XLCellValue object.
         */
        XLValueBoolean(unsigned int boolValue, XLCellValue &parent);

        /**
         * @brief Constructor
         * @param boolValue A bool withe the value used to initialize the object.
         * @param parent A reference to the parent XLCellValue object.
         */
        XLValueBoolean(bool boolValue, XLCellValue &parent);

        /**
         * @brief Copy constructor
         * @param other The object to be copied.
         * @note This uses the default implementation.
         */
        XLValueBoolean(const XLValueBoolean &other) = default;

        /**
         * @brief Move constructor
         * @param other The object to be moved.
         * @note This uses the default implementation.
         */
        XLValueBoolean(XLValueBoolean &&other) = delete;

        /**
         * @brief Destructor
         * @note This uses the default implementation.
         */
        ~XLValueBoolean() override = default;

        /**
         * @brief Copy assignment operator.
         * @param other The object to be assigned from.
         * @return A reference to this object, with the new value.
         * @note This uses the default implementation.
         */
        XLValueBoolean &operator=(const XLValueBoolean &other) = default;

        /**
         * @brief Move assignment operator.
         * @param other The object to be moved.
         * @return A reference to this object, with the moved value.
         * @note This uses the default implementation.
         * @todo Is this the right way to do it?
         */
        XLValueBoolean &operator=(XLValueBoolean &&other) = default;

        /**
         * @brief Creates a polymorphic clone of the object.
         * @param parent The parent XLCell object of the clone.
         * @return A unique_ptr with the clone.
         */
        std::unique_ptr<XLValue> Clone(XLCell &parent) override;

        /**
         * @brief Set the bool value
         * @param boolValue The boolean value to the the object to.
         */
        void Set(XLBool boolValue);

        /**
         * @brief Get the value of the boolean.
         * @return An XLBool object with the value.
         */
        XLBool Boolean() const;

        /**
         * @brief Get the value as a string.
         * @return A std::string with "True" or "False", depending on the value.
         */
        std::string AsString() const override;

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

    };

}


#endif //OPENXLEXE_XLBOOLEAN_H

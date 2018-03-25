//
// Created by Kenneth Balslev on 04/01/2018.
//

#ifndef OPENXLEXE_XLERROR_H
#define OPENXLEXE_XLERROR_H

#include "XLValue.h"
namespace RapidXLSX
{

//======================================================================================================================
//========== XLValueError Class =============================================================================================
//======================================================================================================================

    /**
     * @brief This class encapsulates the idea of a cell with an error value. This happens if there is an error in
     * the formula.
     */
    class XLValueError: public XLValue
    {

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief
         * @param parent
         */
        explicit XLValueError(XLCellValue &parent);

        /**
         * @brief
         * @param other
         */
        XLValueError(const XLValueError &other) = default;

        /**
         * @brief
         * @param other
         */
        XLValueError(XLValueError &&other) = delete;

        /**
         * @brief
         */
        ~XLValueError() override = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLValueError &operator=(const XLValueError &other) = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLValueError &operator=(XLValueError &&other) noexcept = default;

        /**
         * @brief
         * @param parent
         * @return
         */
        std::unique_ptr<XLValue> Clone(XLCell &parent) override;

        /**
         * @brief
         * @return
         */
        XLValueType ValueType() const override;

        /**
         * @brief
         * @return
         */
        XLCellType CellType() const override;

        /**
         * @brief
         * @return
         */
        std::string TypeString() const override;

        /**
         * @brief
         * @return
         */
        std::string AsString() const override;

    };

}


#endif //OPENXLEXE_XLERROR_H

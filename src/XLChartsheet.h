//
// Created by Troldal on 04/10/16.
//

#ifndef OPENXL_XLCHARTSHEET_H
#define OPENXL_XLCHARTSHEET_H

#include "XLAbstractSheet.h"

namespace RapidXLSX
{

//======================================================================================================================
//========== XLChartsheet Class ========================================================================================
//======================================================================================================================

    /**
     * @brief Class representing the an Excel chartsheet.
     * @todo This class is largely unimplemented and works just as a placeholder.
     */
    class XLChartsheet: public XLAbstractSheet
    {

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief
         * @param name
         * @param parent
         * @param filePath
         */
        explicit XLChartsheet(XLWorkbook &parent,
                              const std::string &name,
                              const std::string &filePath);

        /**
         * @brief
         * @param other
         */
        XLChartsheet(const XLChartsheet &other) = default;

        /**
         * @brief
         */
        virtual ~XLChartsheet() = default;

        /**
         * @brief
         * @return
         */
        XLChartsheet &operator=(const XLChartsheet &) = default;

    };

}

#endif //OPENXL_XLCHARTSHEET_H

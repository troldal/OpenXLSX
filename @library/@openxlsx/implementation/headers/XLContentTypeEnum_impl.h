//
// Created by Troldal on 2019-03-01.
//

#ifndef OPENXLSX_XLCONTENTTYPEENUM_IMPL_H
#define OPENXLSX_XLCONTENTTYPEENUM_IMPL_H

namespace OpenXLSX::Impl {
    //======================================================================================================================
    //========== XLColumnVector Enum =======================================================================================
    //======================================================================================================================

    /**
     * @brief
     */
    enum class XLContentType {
        Workbook,
        WorkbookMacroEnabled,
        Worksheet,
        Chartsheet,
        ExternalLink,
        Theme,
        Styles,
        SharedStrings,
        Drawing,
        Chart,
        ChartStyle,
        ChartColorStyle,
        ControlProperties,
        CalculationChain,
        VBAProject,
        CoreProperties,
        ExtendedProperties,
        CustomProperties,
        Comments,
        Table,
        VMLDrawing,
        Unknown
    };
} // namespace OpenXLSX::Impl

#endif //OPENXLSX_XLCONTENTTYPEENUM_IMPL_H

//
// Created by Troldal on 2019-03-01.
//

#ifndef OPENXLSX_XLCONTENTTYPEENUM_IMPL_H
#define OPENXLSX_XLCONTENTTYPEENUM_IMPL_H

#include <XLDefinitions.h>

namespace OpenXLSX::Impl
{

    //======================================================================================================================
    //========== XLColumnVector Enum =======================================================================================
    //======================================================================================================================

    /**
     * @brief
     */
    enum class XLContentType
    {
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

    //======================================================================================================================
    //========== XLRelationshipType Enum ===================================================================================
    //======================================================================================================================

    /**
     * @brief An enum of the possible relationship (or XML document) types used in relationship (.rels) XML files.
     */
    enum class XLRelationshipType
    {
        CoreProperties,
        ExtendedProperties,
        CustomProperties,
        Workbook,
        Worksheet,
        ChartSheet,
        DialogSheet,
        MacroSheet,
        CalculationChain,
        ExternalLink,
        ExternalLinkPath,
        Theme,
        Styles,
        Chart,
        ChartStyle,
        ChartColorStyle,
        Image,
        Drawing,
        VMLDrawing,
        SharedStrings,
        PrinterSettings,
        VBAProject,
        ControlProperties,
        Unknown
    };

} // namespace OpenXLSX::Impl

#endif //OPENXLSX_XLCONTENTTYPEENUM_IMPL_H

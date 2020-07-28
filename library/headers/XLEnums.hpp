//
// Created by Troldal on 2019-03-01.
//

#ifndef OPENXLSX_XLCONTENTTYPEENUM_IMPL_H
#define OPENXLSX_XLCONTENTTYPEENUM_IMPL_H

// ===== OpenXLSX Includes ===== //
#include "XLDefinitions.hpp"

namespace OpenXLSX
{
    //======================================================================================================================
    //========== XLColumnVector Enum
    //=======================================================================================
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

    //======================================================================================================================
    //========== XLRelationshipType Enum
    //===================================================================================
    //======================================================================================================================

    /**
     * @brief An enum of the possible relationship (or XML document) types used in relationship (.rels) XML files.
     */
    enum class XLRelationshipType {
        CoreProperties,
        ExtendedProperties,
        CustomProperties,
        Workbook,
        Worksheet,
        Chartsheet,
        Dialogsheet,
        Macrosheet,
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

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLCONTENTTYPEENUM_IMPL_H

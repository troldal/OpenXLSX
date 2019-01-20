//
// Created by Troldal on 2019-01-11.
//

#ifndef OPENXLSX_XLPROPERTY_H
#define OPENXLSX_XLPROPERTY_H

namespace OpenXLSX
{

//======================================================================================================================
//========== XLDocumentProperties Enum =================================================================================
//======================================================================================================================

    /**
     * @brief The XLDocumentProperties class is an enumeration of the possible properties (metadata) that can be set
     * for a XLDocument object (and .xlsx file)
     */
    enum class XLProperty
    {
        Title,
        Subject,
        Creator,
        Keywords,
        Description,
        LastModifiedBy,
        LastPrinted,
        CreationDate,
        ModificationDate,
        Category,
        Application,
        DocSecurity,
        ScaleCrop,
        Manager,
        Company,
        LinksUpToDate,
        SharedDoc,
        HyperlinkBase,
        HyperlinksChanged,
        AppVersion
    };
    
//======================================================================================================================
//========== XLSheetType Enum ==========================================================================================
//======================================================================================================================

    /**
     * @brief The XLSheetType class is an enumeration of the available sheet types, e.g. Worksheet (ordinary
     * spreadsheets), and Chartsheet (sheets with only a chart).
     */
    enum class XLSheetType
    {
        WorkSheet, ChartSheet, DialogSheet, MacroSheet
    };


//======================================================================================================================
//========== XLSheetState Enum =========================================================================================
//======================================================================================================================

    /**
     * @brief The XLSheetState is an enumeration of the possible (visibility) states, e.g. Visible or Hidden.
     */
    enum class XLSheetState
    {
        Visible, Hidden, VeryHidden
    };
}

#endif //OPENXLSX_XLPROPERTY_H

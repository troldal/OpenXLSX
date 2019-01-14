//
// Created by Troldal on 2019-01-11.
//

#ifndef OPENXLSX_XLPROPERTY_H
#define OPENXLSX_XLPROPERTY_H

namespace OpenXLSX {

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
}

#endif //OPENXLSX_XLPROPERTY_H

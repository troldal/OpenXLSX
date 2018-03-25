//
// Created by Troldal on 02/09/16.
//

#ifndef OPENXL_XLABSTRACTSHEET_H
#define OPENXL_XLABSTRACTSHEET_H

#include "XLAbstractXMLFile.h"
#include "XLSpreadsheetElement.h"

namespace RapidXLSX
{
    class XLContentItem;
    class XLRelationshipItem;
    class XLWorkbook;


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


//======================================================================================================================
//========== XLAbstractSheet Class =====================================================================================
//======================================================================================================================

    /**
     * @brief The XLAbstractSheet is a generalized sheet class, which functions as superclass for specialized classes,
     * such as XLWorksheet. It implements functionality common to all sheet types. This is a pure abstract class,
     * so it cannot be instantiated.
     */
    class XLAbstractSheet: public XLAbstractXMLFile,
                           public XLSpreadsheetElement
    {
        friend class XLWorkbook;
        friend class XLCell;


//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief The constructor. There are no default constructor, so all parameters must be provided for
         * constructing an XLAbstractSheet object. Since this is a pure abstract class, instantiation is only
         * possible via one of the derived classes.
         * @param name The name of the new sheet.
         * @param parent A pointer to the parent XLDocument object.
         * @param filepath A std::string with the relative path to the sheet file in the .xlsx package.
         */
        explicit XLAbstractSheet(XLWorkbook &parent,
                                 const std::string &name,
                                 const std::string &filepath);

        /**
         * @brief The copy constructor.
         * @param other The object to be copied.
         * @note The default copy constructor is used, i.e. only shallow copying of pointer data members.
         * @todo Can this method be deleted?
         */
        XLAbstractSheet(const XLAbstractSheet &other) = default;

        /**
         * @brief The destructor
         * @note The default destructor is used, since cleanup of pointer data members is not required.
         */
        ~XLAbstractSheet() override = default;

        /**
         * @brief Assignment operator
         * @return A reference to the new object.
         * @note The default assignment operator is used, i.e. only shallow copying of pointer data members.
         * @todo Can this method be deleted?
         */
        XLAbstractSheet &operator=(const XLAbstractSheet &) = default;

        /**
         * @brief Method to retrieve the name of the sheet.
         * @return A std::string with the sheet name.
         */
        virtual const std::string &Name() const;

        /**
         * @brief Method for renaming the sheet.
         * @param name A std::string with the new name.
         */
        virtual void SetName(const std::string &name);

        /**
         * @brief Method for getting the current visibility state of the sheet.
         * @return An XLSheetState enum object, with the current sheet state.
         */
        virtual const XLSheetState &State() const;

        /**
         * @brief Method for setting the state of the sheet.
         * @param state An XLSheetState enum object with the new state.
         */
        virtual void SetState(XLSheetState state);

        /**
         * @brief Method to get the type of the sheet.
         * @return An XLSheetType enum object with the sheet type.
         */
        virtual const XLSheetType &Type() const;

        /**
         * @brief Method for cloning the sheet.
         * @param newName A std::string with the name of the clone
         * @return A pointer to the cloned object.
         * @note This is a pure abstract method. I.e. it is implemented in subclasses.
         */
        virtual XLAbstractSheet &Clone(const std::string &newName) = 0;

        /**
         * @brief Method for getting the index of the sheet.
         * @return An int with the index of the sheet.
         */
        virtual unsigned int Index() const;

        /**
         * @brief Method for setting the index of the sheet. This effectively moves the sheet to a different position.
         */
        virtual void SetIndex();


//----------------------------------------------------------------------------------------------------------------------
//           Protected Member Functions
//----------------------------------------------------------------------------------------------------------------------

    protected:
        /**
         * @brief Method for deleting the sheet from the workbook.
         * @todo Is this the best way to do this? May end up with an invalid object.
         */
        virtual void Delete();

        /**
         * @brief
         * @return
         */
        virtual const std::string &SheetName();


//----------------------------------------------------------------------------------------------------------------------
//           Private Member Variables
//----------------------------------------------------------------------------------------------------------------------

    private:
        std::string m_sheetName; /**< The sheet name given by the user */
        XLSheetType m_sheetType; /**< The sheet type, i.e. WorkSheet, ChartSheet, etc. */
        XLSheetState m_sheetState; /**< The state of the sheet, i.e. Visible, Hidden or VeryHidden */

        XMLNode *m_nodeInWorkbook; /**< A pointer to the relevant sheet node in workbook.xml */
        XMLNode *m_nodeInApp; /**< A pointer to the relevant TitleOfParts node in app.xml */
        XLContentItem *m_nodeInContentTypes; /**< A pointer to the relevant content type item in [Content_Types].xml */
        XLRelationshipItem *m_nodeInWorkbookRels; /**< A pointer to the relationship item in workbook.xml.rels */

    };
}

#endif //OPENXL_XLABSTRACTSHEET_H

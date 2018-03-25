//
// Created by Troldal on 24/07/16.
//

#ifndef OPENXL_XLDOCUMENT_H
#define OPENXL_XLDOCUMENT_H

#include <string>
#include <map>
#include <vector>
#include <iostream>
#include <fstream>
#include <boost/filesystem.hpp>
#include "XLContentTypes.h"
#include "XLDocumentAppProperties.h"
#include "XLDocumentCoreProperties.h"
#include "XLSharedStrings.h"
#include "XLStyles.h"
#include "XLWorkbook.h"
#include "XLAbstractSheet.h"
#include "XLRelationships.h"
#include "XLArchive.h"

namespace RapidXLSX
{

    class XLSharedStrings;
    class XLWorkbook;
    class XLWorksheet;

//======================================================================================================================
//========== XLDocumentProperties Enum =================================================================================
//======================================================================================================================

    /**
     * @brief The XLDocumentProperties class is an enumeration of the possible properties (metadata) that can be set
     * for a XLDocument object (and .xlsx file)
     */
    enum class XLDocumentProperties
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
//========== XLDocument Class ==========================================================================================
//======================================================================================================================

    /**
     * @brief This class encapsulates the concept of an excel file. It is different from the XLWorkbook, in that an
     * XLDocument holds an XLWorkbook together with its metadata, as well as methods for opening,
     * closing and saving the document.\n<b><em>The XLDocument is the entrypoint for clients using the RapidXLSX library.</em></b>
     */
    class XLDocument
    {
        friend class XLAbstractXMLFile;
        friend class XLWorkbook;
        friend class XLAbstractSheet;
        friend class XLSpreadsheetElement;

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief Constructor. The default constructor with no arguments.
         */
        explicit XLDocument();

        /**
         * @brief Constructor. An alternative constructor, taking the path to the .xlsx file as an argument.
         * @param docPath A std::string with the path to the .xlsx file.
         */
        explicit XLDocument(const std::string &docPath);

        /**
         * @brief Copy constructor
         * @param other The object to copy
         * @note Copy constructor explicitly deleted.
         */
        XLDocument(const XLDocument &other) = delete;

        /**
         * @brief Destructor
         */
        virtual ~XLDocument();

        /**
         * @brief Open the .xlsx file with the given path
         * @param fileName The path of the .xlsx file to open
         * @todo Consider opening the zipped files as streams, instead of unpacking to a temporaty folder
         */
        void OpenDocument(const std::string &fileName);

        /**
         * @brief Create a new .xlsx file with the given name.
         * @param fileName The path of the new .xlsx file.
         */
        void CreateDocument(const std::string &fileName);

        /**
         * @brief Close the current document
         */
        void CloseDocument() const;

        /**
         * @brief Save the current document using the current filename, overwriting the existing file.
         * @return true if successful; otherwise false.
         */
        bool SaveDocument();

        /**
         * @brief Save the document with a new name. If a file exists with that name, it will be overwritten.
         * @param fileName The path of the file
         * @return true if successful; otherwise false.
         */
        bool SaveDocumentAs(const std::string &fileName);

        /**
         * @brief Get the filename of the current document, e.g. "spreadsheet.xlsx".
         * @return A std::string with the filename.
         */
        std::string DocumentName() const;

        /**
         * @brief Get the full path of the current document, e.g. "drive/blah/spreadsheet.xlsx"
         * @return A std::string with the path.
         */
        std::string DocumentPath() const;

        /**
         * @brief Get the underlying workbook object.
         * @return A pointer to the XLWorkbook object
         */
        XLWorkbook &Workbook();

        /**
         * @brief
         * @return
         */
        const XLWorkbook &Workbook() const;

        /**
         * @brief Get the requested document property.
         * @param theProperty The name of the property to get.
         * @return The property as a string
         */
        std::string GetProperty(XLDocumentProperties theProperty) const;

        /**
         * @brief Set a property
         * @param theProperty The property to set.
         * @param value The value of the property, as a string
         */
        void SetProperty(XLDocumentProperties theProperty,
                         const std::string &value);

        /**
         * @brief Delete the property from the document
         * @param propertyName The property to delete from the document
         */
        void DeleteProperty(const std::string &propertyName);

        /**
         * @brief Get the root directory path of the temporary file structure created for the file.
         * @return A path object with the root directory.
         * @todo Consider making this a protected member.
         */
        const boost::filesystem::path &RootDirectory() const;

//----------------------------------------------------------------------------------------------------------------------
//           Protected Member Functions

//----------------------------------------------------------------------------------------------------------------------

    protected:

        /**
         * @brief Method for adding a ew (XML) file to the .xlsx package.
         * @param path The relative path of the new file.
         * @param content The contents (XML data) of the new file.
         */
        void AddFile(const std::string &path,
                     const std::string &content);

        /**
         * @brief Get the xml node in the app.xml file, for the sheet name.
         * @param sheetName A std::string with the name of the sheet.
         * @return A pointer to the XMLNode object.
         */
        XMLNode &SheetNameNode(const std::string &sheetName);

        /**
         * @brief Get the content item element in the contenttypes.xml file.
         * @param path A std::string with the relative path to the file in question.
         * @return A pointer to the XLContentItem.
         */
        XLContentItem &ContentItem(const std::string &path);

        /**
         * @brief
         * @param contentPath
         * @param contentType
         * @return
         */
        XLContentItem &AddContentItem(const std::string &contentPath, XLContentType contentType);

        /**
         * @brief Load all the associated files (app.xml, core.xml and workbook.xml)
         * @todo Include a CustomProperties object.
         */
        void LoadAllFiles();

        /**
         * @brief Getter method for the App Properties object.
         * @return A pointer to the XLDocAppProperties object.
         */
        XLDocAppProperties &AppProperties();

        /**
         * @brief Getter method for the App Properties object.
         * @return A pointer to the const XLDocAppProperties object.
         */
        const XLDocAppProperties &AppProperties() const;

        /**
         * @brief Getter method for the Core Properties object.
         * @return A pointer to the XLDocCoreProperties object.
         */
        XLDocCoreProperties &CoreProperties();

        /**
         * @brief Getter method for the Core Properties object.
         * @return A pointer to the const XLDocCoreProperties object.
         */
        const XLDocCoreProperties &CoreProperties() const;

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Variables
//----------------------------------------------------------------------------------------------------------------------

    private:

        boost::filesystem::path m_filePath; /**< The path to the original file*/
        boost::filesystem::path m_tempPath; /**< The path to the temporary folder with the package contents*/

        std::unique_ptr<XLRelationships> m_documentRelationships; /**< A pointer to the document relationships object*/
        std::unique_ptr<XLContentTypes> m_contentTypes; /**< A pointer to the content types object*/
        std::unique_ptr<XLDocAppProperties> m_docAppProperties; /**< A pointer to the App properties object */
        std::unique_ptr<XLDocCoreProperties> m_docCoreProperties; /**< A pointer to the Core properties object*/
        std::unique_ptr<XLWorkbook> m_workbook; /**< A pointer to the workbook object */

        std::map<std::string, XLAbstractXMLFile *> m_xmlFiles; /**< A std::map with all the associated XML files*/
        std::unique_ptr<XLArchive> m_archive; /**<A pointer the archive with the files.*/

    };
}

#endif //OPENXL_XLDOCUMENT_H

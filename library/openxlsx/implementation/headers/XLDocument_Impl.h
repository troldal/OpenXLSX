/*

   ____                               ____      ___ ____       ____  ____      ___
  6MMMMb                              `MM(      )M' `MM'      6MMMMb\`MM(      )M'
 8P    Y8                              `MM.     d'   MM      6M'    ` `MM.     d'
6M      Mb __ ____     ____  ___  __    `MM.   d'    MM      MM        `MM.   d'
MM      MM `M6MMMMb   6MMMMb `MM 6MMb    `MM. d'     MM      YM.        `MM. d'
MM      MM  MM'  `Mb 6M'  `Mb MMM9 `Mb    `MMd       MM       YMMMMb     `MMd
MM      MM  MM    MM MM    MM MM'   MM     dMM.      MM           `Mb     dMM.
MM      MM  MM    MM MMMMMMMM MM    MM    d'`MM.     MM            MM    d'`MM.
YM      M9  MM    MM MM       MM    MM   d'  `MM.    MM            MM   d'  `MM.
 8b    d8   MM.  ,M9 YM    d9 MM    MM  d'    `MM.   MM    / L    ,M9  d'    `MM.
  YMMMM9    MMYMMM9   YMMMM9 _MM_  _MM_M(_    _)MM_ _MMMMMMM MYMMMM9 _M(_    _)MM_
            MM
            MM
           _MM_

  Copyright (c) 2018, Kenneth Troldal Balslev

  All rights reserved.

  Redistribution and use in source and binary forms, with or without
  modification, are permitted provided that the following conditions are met:
  - Redistributions of source code must retain the above copyright
    notice, this list of conditions and the following disclaimer.
  - Redistributions in binary form must reproduce the above copyright
    notice, this list of conditions and the following disclaimer in the
    documentation and/or other materials provided with the distribution.
  - Neither the name of the author nor the
    names of any contributors may be used to endorse or promote products
    derived from this software without specific prior written permission.

  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
  ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
  WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
  DISCLAIMED. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY
  DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
  (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
  LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
  ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
  (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
  SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

 */

#ifndef OPENXLSX_IMPL_XLDOCUMENT_H
#define OPENXLSX_IMPL_XLDOCUMENT_H

// ===== Standard Library Includes ===== //
#include <string>
#include <map>
#include <vector>
#include <iostream>
#include <fstream>

// ===== OpenXLSX Includes ===== //
#include "XLAppProperties_Impl.h"
#include "XLCoreProperties_Impl.h"
#include "XLWorkbook_Impl.h"
#include "XLEnums_impl.h"
#include "XLRelationships_Impl.h"
#include "XLException.h"
#include "XLXml_Impl.h"
#include "XLDefinitions.h"
#include <Zippy.h>

namespace OpenXLSX::Impl
{

    class XLContentItem;

    class XLContentTypes;

    //======================================================================================================================
    //========== XLDocument Class ==========================================================================================
    //======================================================================================================================

    /**
     * @brief This class encapsulates the concept of an excel file. It is different from the XLWorkbook, in that an
     * XLDocument holds an XLWorkbook together with its metadata, as well as methods for opening,
     * closing and saving the document.\n<b><em>The XLDocument is the entrypoint for clients
     * using the RapidXLSX library.</em></b>
     */
    class XLDocument
    {

        //----------------------------------------------------------------------------------------------------------------------
        //           Friends
        //----------------------------------------------------------------------------------------------------------------------
        friend class XLAbstractXMLFile;

        friend class XLWorkbook;

        friend class XLSheet;

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
        explicit XLDocument(const std::string& docPath);

        /**
         * @brief Copy constructor
         * @param other The object to copy
         * @note Copy constructor explicitly deleted.
         * @todo Consider implementing this, as it may make sense to copy the entire document, although it may not make
         * sense to copy individual elements (e.g. the XLWorkbook object). Alternatively, implement a 'Clone' function.
         */
        XLDocument(const XLDocument& other) = delete;

        /**
         * @brief Destructor
         */
        virtual ~XLDocument();

        /**
         * @brief Open the .xlsx file with the given path
         * @param fileName The path of the .xlsx file to open
         * @todo Consider opening the zipped files as streams, instead of unpacking to a temporary folder
         */
        void OpenDocument(const std::string& fileName);

        /**
         * @brief Create a new .xlsx file with the given name.
         * @param fileName The path of the new .xlsx file.
         */
        void CreateDocument(const std::string& fileName);

        /**
         * @brief Close the current document
         */
        void CloseDocument();

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
        bool SaveDocumentAs(const std::string& fileName);

        /**
         * @brief Get the filename of the current document, e.g. "spreadsheet.xlsx".
         * @return A std::string with the filename.
         */
        const std::string& DocumentName() const;

        /**
         * @brief Get the full path of the current document, e.g. "drive/blah/spreadsheet.xlsx"
         * @return A std::string with the path.
         */
        const std::string& DocumentPath() const;

        /**
         * @brief Get the underlying workbook object.
         * @return A pointer to the XLWorkbook object
         */
        XLWorkbook* Workbook();

        /**
         * @brief Get the underlying workbook object, as a const object.
         * @return A const pointer to the XLWorkbook object.
         */
        const XLWorkbook* Workbook() const;

        /**
         * @brief Get the requested document property.
         * @param theProperty The name of the property to get.
         * @return The property as a string
         */
        std::string GetProperty(XLProperty theProperty) const;

        /**
         * @brief Set a property
         * @param theProperty The property to set.
         * @param value The value of the property, as a string
         */
        void SetProperty(XLProperty theProperty, const std::string& value);

        /**
         * @brief Delete the property from the document
         * @param propertyName The property to delete from the document
         */
        void DeleteProperty(XLProperty theProperty);

        //----------------------------------------------------------------------------------------------------------------------
        //           Protected Member Functions
        //----------------------------------------------------------------------------------------------------------------------

    protected:

        /**
         * @brief Method for adding a new (XML) file to the .xlsx package.
         * @param path The relative path of the new file.
         * @param content The contents (XML data) of the new file.
         */
        void AddOrReplaceXMLFile(const std::string& path, const std::string& content);

        /**
         * @brief Get an XML file from the .xlsx archive.
         * @param path The relative path of the file.
         * @return A std::string with the content of the file
         */
        std::string GetXMLFile(const std::string& path);

        /**
         * @brief Delete a file from the .xlsx archive.
         * @param path The path of the file to delete.
         */
        void DeleteXMLFile(const std::string& path);

        /**
         * @brief Get the xml node in the app.xml file, for the sheet name.
         * @param sheetName A std::string with the name of the sheet.
         * @return A pointer to the XMLNode object.
         */
        XMLNode SheetNameNode(const std::string& sheetName);

        /**
         * @brief Get the content item element in the contenttypes.xml file.
         * @param path A std::string with the relative path to the file in question.
         * @return A pointer to the XLContentItem.
         */
        XLContentItem ContentItem(const std::string& path);

        /**
         * @brief
         * @param contentPath
         * @param contentType
         * @return
         */
        XLContentItem AddContentItem(const std::string& contentPath, XLContentType contentType);

        /**
         * @brief
         * @param item
         */
        void DeleteContentItem(XLContentItem& item);

        /**
         * @brief Getter method for the App Properties object.
         * @return A pointer to the XLDocAppProperties object.
         */
        XLAppProperties* AppProperties();

        /**
         * @brief Getter method for the App Properties object.
         * @return A pointer to the const XLDocAppProperties object.
         */
        const XLAppProperties* AppProperties() const;

        /**
         * @brief Getter method for the Core Properties object.
         * @return A pointer to the XLDocCoreProperties object.
         */
        XLCoreProperties* CoreProperties();

        /**
         * @brief Getter method for the Core Properties object.
         * @return A pointer to the const XLDocCoreProperties object.
         */
        const XLCoreProperties* CoreProperties() const;

    private:

        template<typename T>
        typename std::enable_if_t<std::is_base_of<XLAbstractXMLFile, T>::value, std::unique_ptr<T>>
        CreateItem(const std::string& target) {

            if (!m_documentRelationships->TargetExists(target))
                throw XLException("Target does not exist!");
            return std::make_unique<T>(*this, m_documentRelationships->RelationshipByTarget(target).Target().value());
        }


        //----------------------------------------------------------------------------------------------------------------------
        //           Private Member Variables
        //----------------------------------------------------------------------------------------------------------------------

    private:

        std::string m_filePath; /**< The path to the original file*/

        std::unique_ptr<XLRelationships> m_documentRelationships; /**< A pointer to the document relationships object*/
        std::unique_ptr<XLContentTypes> m_contentTypes; /**< A pointer to the content types object*/
        std::unique_ptr<XLAppProperties> m_docAppProperties; /**< A pointer to the App properties object */
        std::unique_ptr<XLCoreProperties> m_docCoreProperties; /**< A pointer to the Core properties object*/
        std::unique_ptr<XLWorkbook> m_workbook; /**< A pointer to the workbook object */
        std::unique_ptr<Zippy::ZipArchive> m_archive; /**<  */

    };

} // namespace OpenXLSX::Impl

#endif //OPENXLSX_IMPL_XLDOCUMENT_H

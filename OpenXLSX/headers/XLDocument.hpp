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

#ifndef OPENXLSX_XLDOCUMENT_HPP
#define OPENXLSX_XLDOCUMENT_HPP

#pragma warning(push)
#pragma warning(disable : 4251)
#pragma warning(disable : 4275)

// ===== External Includes ===== //
#include <algorithm>
#include <fstream>
#include <iostream>
#include <list>
#include <string>

// ===== OpenXLSX Includes ===== //
#include "IZipArchive.hpp"
#include "OpenXLSX-Exports.hpp"
#include "XLCommandQuery.hpp"
#include "XLContentTypes.hpp"
#include "XLException.hpp"
#include "XLProperties.hpp"
#include "XLRelationships.hpp"
#include "XLSharedStrings.hpp"
#include "XLStyles.hpp"
#include "XLWorkbook.hpp"
#include "XLXmlData.hpp"
#include "XLZipArchive.hpp"

namespace OpenXLSX
{
    /**
     * @brief The XLDocumentProperties class is an enumeration of the possible properties (metadata) that can be set
     * for a XLDocument object (and .xlsx file)
     */
    enum class XLProperty {
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

    /**
     * @brief This class encapsulates the concept of an excel file. It is different from the XLWorkbook, in that an
     * XLDocument holds an XLWorkbook together with its metadata, as well as methods for opening,
     * closing and saving the document.\n<b><em>The XLDocument is the entrypoint for clients
     * using the RapidXLSX library.</em></b>
     */
    class OPENXLSX_EXPORT XLDocument final
    {
        //----- Friends
        friend class XLXmlFile;
        friend class XLWorkbook;
        friend class XLSheet;
        friend class XLStyles;
        friend class XLXmlData;

        //---------- Public Member Functions
    public:
        /**
         * @brief Constructor. The default constructor with no arguments.
         */
        explicit XLDocument(const IZipArchive& zipArchive = XLZipArchive());

        /**
         * @brief Constructor. An alternative constructor, taking the path to the .xlsx file as an argument.
         * @param docPath A std::string with the path to the .xlsx file.
         */
        explicit XLDocument(const std::string& docPath, const IZipArchive& zipArchive = XLZipArchive());

        /**
         * @brief Copy constructor
         * @param other The object to copy
         * @note Copy constructor explicitly deleted.
         */
        XLDocument(const XLDocument& other) = delete;

        /**
         * @brief
         * @param other
         */
        XLDocument(XLDocument&& other) noexcept = default;

        /**
         * @brief Destructor
         */
        ~XLDocument();

        /**
         * @brief
         * @param other
         * @return
         */
        XLDocument& operator=(const XLDocument& other) = delete;

        /**
         * @brief
         * @param other
         * @return
         */
        XLDocument& operator=(XLDocument&& other) noexcept = default;

        /**
         * @brief Open the .xlsx file with the given path
         * @param fileName The path of the .xlsx file to open
         */
        void open(const std::string& fileName);

        /**
         * @brief Create a new .xlsx file with the given name.
         * @param fileName The path of the new .xlsx file.
         */
        void create(const std::string& fileName);

        /**
         * @brief Close the current document
         */
        void close();

        /**
         * @brief Save the current document using the current filename, overwriting the existing file.
         * @return true if successful; otherwise false.
         */
        void save();

        /**
         * @brief Save the document with a new name. If a file exists with that name, it will be overwritten.
         * @param fileName The path of the file
         * @return true if successful; otherwise false.
         */
        void saveAs(const std::string& fileName);

        /**
         * @brief Get the filename of the current document, e.g. "spreadsheet.xlsx".
         * @return A std::string with the filename.
         */
        const std::string& name() const;

        /**
         * @brief Get the full path of the current document, e.g. "drive/blah/spreadsheet.xlsx"
         * @return A std::string with the path.
         */
        const std::string& path() const;

        /**
         * @brief Get the underlying workbook object, as a const object.
         * @return A const pointer to the XLWorkbook object.
         */
        XLWorkbook workbook() const;

        /**
         * @brief Get the requested document property.
         * @param prop The name of the property to get.
         * @return The property as a string
         */
        std::string property(XLProperty prop) const;

        /**
         * @brief Set a property
         * @param prop The property to set.
         * @param value The getValue of the property, as a string
         */
        void setProperty(XLProperty prop, const std::string& value);

        /**
         * @brief
         * @return
         */
        explicit operator bool() const;

        /**
         * @brief
         * @return
         */
        bool isOpen() const;

        /**
         * @brief Delete the property from the document
         * @param theProperty The property to delete from the document
         */
        void deleteProperty(XLProperty theProperty);

        /**
         * @brief
         * @param command
         */
        void execCommand(const XLCommand& command);

        /**
         * @brief
         * @param query
         * @return
         */
        XLQuery execQuery(const XLQuery& query) const;

        /**
         * @brief
         * @param query
         * @return
         */
        XLQuery execQuery(const XLQuery& query);

        /**
         * @brief
         * @return
         */
        const XLStyles& styles() const;


        //----------------------------------------------------------------------------------------------------------------------
        //           Protected Member Functions
        //----------------------------------------------------------------------------------------------------------------------

    protected:
        /**
         * @brief Get an XML file from the .xlsx archive.
         * @param path The relative path of the file.
         * @return A std::string with the content of the file
         */
        std::string extractXmlFromArchive(const std::string& path);

        /**
         * @brief
         * @param path
         * @return
         */
        XLXmlData* getXmlDataByPath(const std::string& path);

        /**
         * @brief
         * @param path
         * @return
         */
        const XLXmlData* getXmlDataByPath(const std::string& path) const;

        /**
         * @brief
         * @param path
         * @return
         */
        bool hasXmlData(const std::string& path) const;
        
        /**
         * @brief Explore the tree to find the sheet element
         * @param sheetName to be found
         * @return a pointer to XLXmlData or nullptr if sheetName doesn not exist
         */
        XLXmlData* getXmlDataByName(const std::string& name) const;

        /**
         * @brief
         * @param name
         * @return the path to xl/worksheets/_rels/sheet{0}.xml.rels whether it exist or not.
         * Return a empty string if sheetName does not exists
         */
        std::string getSheetRelsPath(const std::string& sheetName) const;

        /**
         * @brief determine the available id disponible for filename
         * @param type
         * @return Return the available id.
         */
        uint16_t availableFileID(XLContentType type);

        /**
         * @brief determine the available id disponible for Id field
         * @param type
         * @return Return the available id.
         * @note fill the gap if any
         */
        uint16_t availableSheetID();

        /**
         * @brief create a new table in the doc
         * @param sheetName
         * @param tableName
         * @param reference
         */
        void createTable(const std::string& sheetName, const std::string& tableName, const std::string& reference);

        /**
         * @brief delete the corresponding table
         * @param tableName the table to be deleted
         */
        void deleteTable(const std::string& tableName);
        
        //----------------------------------------------------------------------------------------------------------------------
        //           Private Member Variables
        //----------------------------------------------------------------------------------------------------------------------

    private:
        std::string m_filePath {}; /**< The path to the original file*/
        std::string m_realPath {}; /**<  */

        mutable std::list<XLXmlData>    m_data {};     /**<  */
        mutable XLSharedStrings  m_sharedStrings;     /**< The sharedstrings object (one for each doc)*/

        XLRelationships m_docRelationships {}; /**< A pointer to the document relationships object*/
        XLRelationships m_wbkRelationships {}; /**< A pointer to the document relationships object*/
        XLContentTypes  m_contentTypes {};     /**< A pointer to the content types object*/
        XLAppProperties m_appProperties {};    /**< A pointer to the App properties object */
        XLProperties    m_coreProperties {};   /**< A pointer to the Core properties object*/
        XLXmlData*      m_XmlWorkbook {};
        XLWorkbook      m_workbook {};         /**< A pointer to the workbook object */
        IZipArchive     m_archive {};          /**<  */
        XLStyles        m_styles {};           /**< Styles object >*/
    };

}    // namespace OpenXLSX

#pragma warning(pop)
#endif    // OPENXLSX_XLDOCUMENT_HPP

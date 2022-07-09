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
#include <filesystem>
#include <fstream>
#include <iostream>
#include <list>
#include <map>
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
#include "XLWorkbook.hpp"
#include "XLXmlData.hpp"
#include "XLZipArchive.hpp"

namespace OpenXLSX
{
    namespace fs = std::filesystem;

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
        friend class XLXmlData;

        //---------- Public Member Functions
    public:
        /**
         * @brief Constructor. This is the default constructor. It has one argument, which has a default value.
         * @note The zipArchive parameter uses an XLZipArchive object (based on the miniz library) as the default value.
         * This is recommended in most cases. However, in rare cases a different library may be used. The only requirement
         * is that the object passed as an argument subscribes to the interface defined by IZipArchive.
         * @param zipArchive The zip archive object, used to read and modify the .zip archive holding the spreadsheet data.
         */
        explicit XLDocument(const IZipArchive& zipArchive = XLZipArchive());

        /**
         * @brief Constructor. An alternative constructor, taking the path to the .xlsx file as an argument.
         * @param docPath A std::string with the path to the .xlsx file.
         * @param zipArchive
         */
        explicit XLDocument(const fs::path& docPath, const IZipArchive& zipArchive = XLZipArchive());

        /**
         * @brief
         * @param docPath
         * @param tempDir
         * @param zipArchive
         */
        explicit XLDocument(const fs::path& docPath, const fs::path& tempDir, const IZipArchive& zipArchive = XLZipArchive());

        /**
         * @brief Copy constructor
         * @param other The object to copy
         * @note Copy constructor explicitly deleted, as copying would cause a copy of all spreadsheet data.
         * @todo Consider "copying" using reference semantics, by using the pimpl idiom.
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
        void open(const fs::path& fileName);

        /**
         * @brief Create a new .xlsx file with the given name.
         * @param fileName The path of the new .xlsx file.
         */
        void create(const fs::path& fileName);

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
        void saveAs(const fs::path& fileName);

        /**
         * @brief Get the filename of the current document, e.g. "spreadsheet.xlsx".
         * @return A std::string with the filename.
         */
        fs::path name() const;

        /**
         * @brief Get the full path of the current document, e.g. "drive/blah/spreadsheet.xlsx"
         * @return A std::string with the path.
         */
        const fs::path& path() const;

        /**
         * @brief Set the path of the directory where the temporary file will be placed while opening/editing.
         * @details Set the path of the directory where the temporary file will be placed while opening/editing.
         * By default, and if a blank path is passed as the argument, the temporary file will be placed in the
         * same directory as the original file.
         * Typically, it is not necessary to set the temporary directory. But in some cases, the application
         * may not have write access to the path where the target spreadsheet file is located. In that case, it
         * can be useful to use a separate directory for the temporary file. This could be the path provided by
         * the std::filesystem::temp_directory_path() function, which will return the path to a directory suitable
         * for temporary files, and which is guaranteed to exist on all platforms.
         * @note The temporary directory must be set before opening or creating an Excel spreadsheet.
         * @param path Path to the temporary directory to be used.
         */
        void setTemporaryDir(const fs::path& path = "") const;

        /**
         * @brief Get the path to the temporary directory currently used.
         * @return an fs::path object with the path to the temporary directory.
         */
        const fs::path& temporaryDir() const;

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
        XLXmlData* getXmlData(const std::string& path);

        /**
         * @brief
         * @param path
         * @return
         */
        const XLXmlData* getXmlData(const std::string& path) const;

        /**
         * @brief
         * @param path
         * @return
         */
        bool hasXmlData(const std::string& path) const;

        //----------------------------------------------------------------------------------------------------------------------
        //           Private Member Variables
        //----------------------------------------------------------------------------------------------------------------------

    private:
        fs::path m_filePath {}; /**< The path to the original spreadsheet file */
        mutable fs::path m_tempPath {}; /**< The path to the temporary spreadsheet file */
        mutable fs::path m_tempDir {}; /**< The path to the parent directory of the temporary spreadsheet file */

        mutable std::list<XLXmlData>    m_data {};              /**<  */
        mutable std::deque<std::string> m_sharedStringCache {}; /**<  */
        mutable XLSharedStrings         m_sharedStrings {};     /**<  */

        XLRelationships m_docRelationships {}; /**< A document relationships object*/
        XLRelationships m_wbkRelationships {}; /**< A document relationships object*/
        XLContentTypes  m_contentTypes {};     /**< A content types object*/
        XLAppProperties m_appProperties {};    /**< An app properties object */
        XLProperties    m_coreProperties {};   /**< A core properties object*/
        XLWorkbook      m_workbook {};         /**< A workbook object */
        IZipArchive     m_archive {};          /**< A zip archive interface object */
    };

}    // namespace OpenXLSX

#pragma warning(pop)
#endif    // OPENXLSX_XLDOCUMENT_HPP

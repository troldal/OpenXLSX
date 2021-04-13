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
#include <map>
#include <string>

// ===== OpenXLSX Includes ===== //
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
         * @brief Constructor. The default constructor with no arguments.
         */
        XLDocument() = default;

        /**
         * @brief Constructor. An alternative constructor, taking the path to the .xlsx file as an argument.
         * @param docPath A std::string with the path to the .xlsx file.
         */
        explicit XLDocument(const std::string& docPath);

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

        void resetCalcChain();

        /**
         * @brief Get the requested document property.
         * @param prop The name of the property to get.
         * @return The property as a string
         */
        std::string property(XLProperty prop) const;

        /**
         * @brief Set a property
         * @param prop The property to set.
         * @param value The value of the property, as a string
         */
        void setProperty(XLProperty prop, const std::string& value);

        /**
         * @brief Delete the property from the document
         * @param theProperty The property to delete from the document
         */
        void deleteProperty(XLProperty theProperty);

        /**
         * @brief
         * @param command
         */
        template<typename Command>
        void executeCommand(Command command)
        {
            if constexpr (std::is_same_v<Command, XLCommandSetSheetName>) {
                m_appProperties.setSheetName(command.sheetName(), command.newName());
                m_workbook.setSheetName(command.sheetID(), command.sheetName());
            }

            else if constexpr (std::is_same_v<Command, XLCommandSetSheetVisibility>) {    // NOLINT
            }

            else if constexpr (std::is_same_v<Command, XLCommandSetSheetColor>) {
            }

            else if constexpr (std::is_same_v<Command, XLCommandSetSheetIndex>) {
                auto sheetName = executeQuery(XLQuerySheetName(command.sheetID())).sheetName();
                m_workbook.setSheetIndex(sheetName, command.sheetIndex());
            }

            else if constexpr (std::is_same_v<Command, XLCommandResetCalcChain>) {
                m_archive.deleteEntry("xl/calcChain.xml");
                m_data.erase(std::find_if(m_data.begin(), m_data.end(), [&](const XLXmlData& item) {
                       return item.getXmlPath() == "xl/calcChain.xml";
                }));
            }

            else if constexpr (std::is_same_v<Command, XLCommandAddSharedStrings>) {
                std::string sharedStrings {
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
                    "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\">\n"
                    "  <si>\n"
                    "    <t/>\n"
                    "  </si>\n"
                    "</sst>"
                };
                m_contentTypes.addOverride("/xl/sharedStrings.xml", XLContentType::SharedStrings);
                m_wbkRelationships.addRelationship(XLRelationshipType::SharedStrings, "sharedStrings.xml");
                m_archive.addEntry("xl/sharedStrings.xml", sharedStrings);
            }

            else if constexpr (std::is_same_v<Command, XLCommandAddWorksheet>) {
                std::string emptyWorksheet {
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
                    " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
                    " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\""
                    " xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">"
                    "<dimension ref=\"A1\"/>"
                    "<sheetViews>"
                    "<sheetView workbookViewId=\"0\"/>"
                    "</sheetViews>"
                    "<sheetFormatPr baseColWidth=\"10\" defaultRowHeight=\"16\" x14ac:dyDescent=\"0.2\"/>"
                    "<sheetData/>"
                    "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
                    "</worksheet>"
                };
                m_contentTypes.addOverride(command.sheetPath(), XLContentType::Worksheet);
                m_wbkRelationships.addRelationship(XLRelationshipType::Worksheet, command.sheetPath().substr(4));
                m_appProperties.appendSheetName(command.sheetName());
                m_archive.addEntry(command.sheetPath().substr(1), emptyWorksheet);
                m_data.emplace_back(
                    /* parentDoc */ this,
                    /* xmlPath   */ command.sheetPath().substr(1),
                    /* xmlID     */ m_wbkRelationships.relationshipByTarget(command.sheetPath().substr(4)).id(),
                    /* xmlType   */ XLContentType::Worksheet);
            }

            else if constexpr (std::is_same_v<Command, XLCommandAddChartsheet>) {
            }

            else if constexpr (std::is_same_v<Command, XLCommandDeleteSheet>) {
                m_appProperties.deleteSheetName(command.sheetName());
                auto sheetPath = "/xl/" + m_wbkRelationships.relationshipById(command.sheetID()).target();
                m_archive.deleteEntry(sheetPath.substr(1));
                m_contentTypes.deleteOverride(sheetPath);
                m_wbkRelationships.deleteRelationship(command.sheetID());
                m_data.erase(std::find_if(m_data.begin(), m_data.end(), [&](const XLXmlData& item) {
                    return item.getXmlPath() == sheetPath.substr(1);
                }));
            }

            else if constexpr (std::is_same_v<Command, XLCommandCloneSheet>) {
                auto internalID = m_workbook.createInternalSheetID();
                auto sheetPath  = "/xl/worksheets/sheet" + std::to_string(internalID) + ".xml";
                if (m_workbook.sheetExists(command.cloneName()))
                    throw XLException("Sheet named \"" + command.cloneName() + "\" already exists.");

                if (m_wbkRelationships.relationshipById(command.sheetID()).type() == XLRelationshipType::Worksheet) {
                    m_contentTypes.addOverride(sheetPath, XLContentType::Worksheet);
                    m_wbkRelationships.addRelationship(XLRelationshipType::Worksheet, sheetPath.substr(4));
                    m_appProperties.appendSheetName(command.cloneName());
                    m_archive.addEntry(sheetPath.substr(1), std::find_if(m_data.begin(), m_data.end(), [&](const XLXmlData& data) {
                                                                return data.getXmlPath().substr(3) ==
                                                                       m_wbkRelationships.relationshipById(command.sheetID()).target();
                                                            })->getRawData());
                    m_data.emplace_back(
                        /* parentDoc */ this,
                        /* xmlPath   */ sheetPath.substr(1),
                        /* xmlID     */ m_wbkRelationships.relationshipByTarget(sheetPath.substr(4)).id(),
                        /* xmlType   */ XLContentType::Worksheet);
                }
                else {
                    m_contentTypes.addOverride(sheetPath, XLContentType::Chartsheet);
                    m_wbkRelationships.addRelationship(XLRelationshipType::Chartsheet, sheetPath.substr(4));
                    m_appProperties.appendSheetName(command.cloneName());
                    m_archive.addEntry(sheetPath.substr(1), std::find_if(m_data.begin(), m_data.end(), [&](const XLXmlData& data) {
                                                                return data.getXmlPath().substr(3) ==
                                                                       m_wbkRelationships.relationshipById(command.sheetID()).target();
                                                            })->getRawData());
                    m_data.emplace_back(
                        /* parentDoc */ this,
                        /* xmlPath   */ sheetPath.substr(1),
                        /* xmlID     */ m_wbkRelationships.relationshipByTarget(sheetPath.substr(4)).id(),
                        /* xmlType   */ XLContentType::Chartsheet);
                }

                m_workbook.prepareSheetMetadata(command.cloneName(), internalID);
            }

            else {
                throw XLException("Invalid command object.");
            }
        }

        /**
         * @brief
         * @tparam Query
         * @param query
         * @return
         */
        template<typename Query>
        Query executeQuery(Query query) const
        {
            if constexpr (std::is_same_v<Query, XLQuerySheetName>) {    // NOLINT
                query.setSheetName(m_workbook.sheetName(query.sheetID()));
                return query;
            }

            else if constexpr (std::is_same_v<Query, XLQuerySheetIndex>) {
                return query;
            }

            else if constexpr (std::is_same_v<Query, XLQuerySheetVisibility>) {
                query.setSheetVisibility(m_workbook.sheetVisibility(query.sheetID()));
                return query;
            }

            else if constexpr (std::is_same_v<Query, XLQuerySheetType>) {
                if (m_wbkRelationships.relationshipById(query.sheetID()).type() == XLRelationshipType::Worksheet)
                    query.setSheetType(XLContentType::Worksheet);
                else
                    query.setSheetType(XLContentType::Chartsheet);

                return query;
            }

            else if constexpr (std::is_same_v<Query, XLQuerySheetID>) {
                query.setSheetVisibility(m_workbook.sheetVisibility(query.sheetName()));
                return query;
            }

            else if constexpr (std::is_same_v<Query, XLQuerySheetRelsID>) {
                query.setSheetID(m_wbkRelationships.relationshipByTarget(query.sheetPath().substr(4)).id());
                return query;
            }

            else if constexpr (std::is_same_v<Query, XLQuerySheetRelsTarget>) {
                query.setSheetTarget(m_wbkRelationships.relationshipById(query.sheetID()).target());
                return query;
            }

            else {
                throw XLException("Invalid query object.");
            }
        }

        /**
         * @brief
         * @tparam Query
         * @param query
         * @return
         */
        template<typename Query>
        Query executeQuery(Query query)
        {
            if constexpr (std::is_same_v<Query, XLQuerySheetName>) {    // NOLINT
                return static_cast<const XLDocument&>(*this).executeQuery(query);
            }

            else if constexpr (std::is_same_v<Query, XLQuerySheetIndex>) {
                return static_cast<const XLDocument&>(*this).executeQuery(query);
            }

            else if constexpr (std::is_same_v<Query, XLQuerySheetVisibility>) {
                return static_cast<const XLDocument&>(*this).executeQuery(query);
            }

            else if constexpr (std::is_same_v<Query, XLQuerySheetType>) {
                return static_cast<const XLDocument&>(*this).executeQuery(query);
            }

            else if constexpr (std::is_same_v<Query, XLQuerySheetID>) {
                return static_cast<const XLDocument&>(*this).executeQuery(query);
            }

            else if constexpr (std::is_same_v<Query, XLQuerySheetRelsID>) {
                return static_cast<const XLDocument&>(*this).executeQuery(query);
            }

            else if constexpr (std::is_same_v<Query, XLQuerySheetRelsTarget>) {
                return static_cast<const XLDocument&>(*this).executeQuery(query);
            }

            else if constexpr (std::is_same_v<Query, XLQuerySharedStrings>) {
                query.setSharedStrings(&m_sharedStrings);
                return query;
            }

            else if constexpr (std::is_same_v<Query, XLQueryXmlData>) {
                auto result =
                    std::find_if(m_data.begin(), m_data.end(), [&](const XLXmlData& item) { return item.getXmlPath() == query.xmlPath(); });
                if (result == m_data.end()) throw XLException("Path does not exist in zip archive.");
                query.setXmlData(&*result);
                return query;
            }

            else {
                throw XLException("Invalid query object.");
            }
        }

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

        //----------------------------------------------------------------------------------------------------------------------
        //           Private Member Variables
        //----------------------------------------------------------------------------------------------------------------------

    private:
        std::string                  m_filePath {}; /**< The path to the original file*/
        std::string                  m_realPath {}; /**<  */
        mutable std::list<XLXmlData> m_data {};     /**<  */

        XLRelationships m_docRelationships; /**< A pointer to the document relationships object*/
        XLRelationships m_wbkRelationships; /**< A pointer to the document relationships object*/
        XLContentTypes  m_contentTypes;     /**< A pointer to the content types object*/
        XLAppProperties m_appProperties;    /**< A pointer to the App properties object */
        XLProperties    m_coreProperties;   /**< A pointer to the Core properties object*/
        XLSharedStrings m_sharedStrings;    /**<  */
        XLWorkbook      m_workbook;         /**< A pointer to the workbook object */
        XLZipArchive    m_archive;          /**<  */
    };

}    // namespace OpenXLSX

#pragma warning(pop)
#endif    // OPENXLSX_XLDOCUMENT_HPP

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

// ===== External Includes ===== //
#include <nowide/fstream.hpp>
#include <pugixml.hpp>
#if defined(_WIN32)
#    include <random>
#endif

#include <vector>
#include <algorithm>

// ===== OpenXLSX Includes ===== //
#include "XLContentTypes.hpp"
#include "XLConstants.hpp"
#include "XLDocument.hpp"
#include "XLSheet.hpp"
#include "XLCellRange.hpp"
#include "XLTemplates.hpp"

using namespace OpenXLSX;

XLDocument::XLDocument(const IZipArchive& zipArchive) : m_archive(zipArchive) {}

/**
 * @details An alternative constructor, taking a std::string with the path to the .xlsx package as an argument.
 */
XLDocument::XLDocument(const std::string& docPath, const IZipArchive& zipArchive) : m_archive(zipArchive)
{
    open(docPath);
}

/**
 * @details The destructor calls the closeDocument method before the object is destroyed.
 */
XLDocument::~XLDocument()
{
    close();
}

/**
 * @details The openDocument method opens the .xlsx package in the following manner:
 * - Check if a document is already open. If yes, close it.
 * - Create a temporary folder for the contents of the .xlsx package
 * - Unzip the contents of the package to the temporary folder.
 * - load the contents into the data structure for manipulation.
 */
void XLDocument::open(const std::string& fileName)
{
    // Check if a document is already open. If yes, close it.
    // TODO: Consider throwing if a file is already open.
    if (m_archive.isOpen()) close();
    m_filePath = fileName;
    m_archive.open(m_filePath);

    // ===== Add and open the Relationships and [Content_Types] files for the document level.
    m_data.emplace_back(this, "[Content_Types].xml");
    m_data.emplace_back(this, "_rels/.rels");
    m_data.emplace_back(this, "xl/_rels/workbook.xml.rels");

    m_contentTypes     = XLContentTypes(getXmlDataByPath("[Content_Types].xml"));
    m_docRelationships = XLRelationships(getXmlDataByPath("_rels/.rels"));
    m_wbkRelationships = XLRelationships(getXmlDataByPath("xl/_rels/workbook.xml.rels"));

    if (!m_archive.hasEntry("xl/sharedStrings.xml"))
        execCommand(XLCommand(XLCommandType::AddSharedStrings));

    // ===== Add remaining spreadsheet elements to the vector of XLXmlData objects.
    for (auto& item : m_contentTypes.getContentItems()) {
        if (item.path().substr(0, 4) == "/xl/" && !(item.path() == "/xl/workbook.xml"))
            m_data.emplace_back(/* parentDoc */ this,
                                /* xmlPath   */ item.path().substr(1),
                                /* xmlID     */ m_wbkRelationships.relationshipByTarget(item.path().substr(4)).id(),
                                /* name      */ std::string(),                 
                                /* xmlType   */ item.type());
        else
            m_data.emplace_back(/* parentDoc */ this,
                                /* xmlPath   */ item.path().substr(1),
                                /* xmlID     */ m_docRelationships.relationshipByTarget(item.path().substr(1)).id(),
                                /* name      */ std::string(),  
                                /* xmlType   */ item.type());
    }

    m_XmlWorkbook    = getXmlDataByPath("xl/workbook.xml");
    m_XmlWorkbook->setName("workbook");
    m_workbook       = XLWorkbook(m_XmlWorkbook);

    // load worksheets/_rels/sheet{0}.xml.rels if any and connect nodes
    for (auto& wsItem : m_data){
        if (wsItem.getXmlType() == XLContentType::Worksheet){
            
            // Connect worksheets to workbook
            m_XmlWorkbook->addChildNode(&wsItem);
            wsItem.setParentNode(m_XmlWorkbook);
           
            // Retrieve sheet name and set it
            wsItem.setName(m_XmlWorkbook->getXmlDocument()->document_element().child("sheets")
                    .find_child_by_attribute("r:id", wsItem.getXmlID().c_str())
                    .attribute("name").value());
            
            std::string sheetRelsPath = getSheetRelsPath(wsItem.getName());
            if (m_archive.hasEntry(sheetRelsPath)){ // there is a xl/worksheets/_rels/sheet{0}.xml.rels

                // Adding the sheet{0}.xml.rels
                XLXmlData* pRels = &(m_data.emplace_back(/* parentDoc */ this,
                                                        /* xmlPath   */ sheetRelsPath,
                                                        /* xmlID     */ std::string(),
                                                        /* name      */ std::string(),
                                                        /* xmlType   */ XLContentType::WorksheetRelations,
                                                        /*parentNode */ &wsItem ));
                        
                wsItem.addChildNode(pRels);

                auto sheetRels = XLRelationships(pRels);
                // Loop throught the sheet relations to connect childrens
                for (auto& childItem : sheetRels.relationships()){
                    std::string targetPath = childItem.target();
                    // Some may be relative path
                    if (targetPath.substr(0,2) == "..") 
                        targetPath = "xl" + targetPath.substr(2);

                    XLXmlData* pDocChild = getXmlDataByPath(targetPath); // Pointer to table elmnt

                    pDocChild->setName(pDocChild->getXmlDocument()->child("table")
                                                .attribute("name").value());
                    wsItem.addChildNode(pDocChild);
                    pDocChild->setParentNode(&wsItem);
                    pDocChild->setXmlID(sheetRels.relationshipByTarget(childItem.target()).id());
                            //auto type = childItem.type();
                }
                        
            }
                
        } // if XLContentType::Worksheet
    } // for m_data

    // ===== Open the workbook and document property items
    // TODO: If property data doesn't exist, consider creating them, instead of ignoring it.
    m_coreProperties = (hasXmlData("docProps/core.xml") ? XLProperties(getXmlDataByPath("docProps/core.xml")) : XLProperties());
    m_appProperties  = (hasXmlData("docProps/app.xml") ? XLAppProperties(getXmlDataByPath("docProps/app.xml")) : XLAppProperties());
    m_sharedStrings  = XLSharedStrings(getXmlDataByPath("xl/sharedStrings.xml"));
}


/**
 * @details Create a new document. This is done by saving the data in XLTemplate.h in binary format.
 */
void XLDocument::create(const std::string& fileName)
{
    // ===== Create a temporary output file stream.
    nowide::ofstream outfile(fileName, std::ios::binary);

    // ===== Stream the binary data for an empty workbook to the output file.
    // ===== Casting, in particular reinterpret_cast, is discouraged, but in this case it is unfortunately unavoidable.
    outfile.write(reinterpret_cast<const char*>(XLTemplate::templateData), XLTemplate::templateSize);    // NOLINT
    outfile.close();

    open(fileName);
}

/**
 * @details The document is closed by deleting the temporary folder structure.
 */
void XLDocument::close()
{
    if (m_archive) m_archive.close();
    m_filePath.clear();
    m_data.clear();

    m_wbkRelationships = XLRelationships();
    m_docRelationships = XLRelationships();
    m_contentTypes     = XLContentTypes();
    m_appProperties    = XLAppProperties();
    m_coreProperties   = XLProperties();
    m_workbook         = XLWorkbook();
}

/**
 * @details Save the document with the same name. The existing file will be overwritten.
 */
void XLDocument::save()
{
    saveAs(m_filePath);
}

/**
 * @details Save the document with a new name. If present, the 'calcChain.xml file will be ignored. The reason for this
 * is that changes to the document may invalidate the calcChain.xml file. Deleting will force Excel to re-create the
 * file. This will happen automatically, without the user noticing.
 */
void XLDocument::saveAs(const std::string& fileName)
{
    m_filePath = fileName;

    // ===== Delete the calcChain.xml file in order to force re-calculation of the sheet
    // TODO: Is this the best way to do it? Maybe there is a flag that can be set, that forces re-calculalion.
    execCommand(XLCommand(XLCommandType::ResetCalcChain));

    // ===== Add all xml items to archive and save the archive.
    for (auto& item : m_data)
        m_archive.addEntry(item.getXmlPath(), item.getRawData());
    m_archive.save(m_filePath);
}

/**
 * @details
 * @todo Currently, this method returns the full path, which is not the intention.
 */
const std::string& XLDocument::name() const
{
    return m_filePath;
}

/**
 * @details
 */
const std::string& XLDocument::path() const
{
    return m_filePath;
}

/**
 * @details Get a const pointer to the underlying XLWorkbook object.
 */
XLWorkbook XLDocument::workbook() const
{
    return m_workbook;
}

/**
 * @details Get the value for a property.
 */
std::string XLDocument::property(XLProperty prop) const
{
    switch (prop) {
        case XLProperty::Application:
            return m_appProperties.property("Application");
        case XLProperty::AppVersion:
            return m_appProperties.property("AppVersion");
        case XLProperty::Category:
            return m_coreProperties.property("cp:category");
        case XLProperty::Company:
            return m_appProperties.property("Company");
        case XLProperty::CreationDate:
            return m_coreProperties.property("dcterms:created");
        case XLProperty::Creator:
            return m_coreProperties.property("dc:creator");
        case XLProperty::Description:
            return m_coreProperties.property("dc:description");
        case XLProperty::DocSecurity:
            return m_appProperties.property("DocSecurity");
        case XLProperty::HyperlinkBase:
            return m_appProperties.property("HyperlinkBase");
        case XLProperty::HyperlinksChanged:
            return m_appProperties.property("HyperlinksChanged");
        case XLProperty::Keywords:
            return m_coreProperties.property("cp:keywords");
        case XLProperty::LastModifiedBy:
            return m_coreProperties.property("cp:lastModifiedBy");
        case XLProperty::LastPrinted:
            return m_coreProperties.property("cp:lastPrinted");
        case XLProperty::LinksUpToDate:
            return m_appProperties.property("LinksUpToDate");
        case XLProperty::Manager:
            return m_appProperties.property("Manager");
        case XLProperty::ModificationDate:
            return m_coreProperties.property("dcterms:modified");
        case XLProperty::ScaleCrop:
            return m_appProperties.property("ScaleCrop");
        case XLProperty::SharedDoc:
            return m_appProperties.property("SharedDoc");
        case XLProperty::Subject:
            return m_coreProperties.property("dc:subject");
        case XLProperty::Title:
            return m_coreProperties.property("dc:title");
        default:
            return "";    // To silence compiler warning.
    }
}

/**
 * @details Set the value for a property.
 *
 * If the property is a datetime, it must be in the W3CDTF format, i.e. YYYY-MM-DDTHH:MM:SSZ. Also, the time should
 * be GMT. Creating a time point in this format can be done as follows:
 * ```cpp
 * #include <iostream>
 * #include <iomanip>
 * #include <ctime>
 * #include <sstream>
 *
 * // other stuff here
 *
 * std::time_t t = std::time(nullptr);
 * std::tm tm = *std::gmtime(&t);
 *
 * std::stringstream ss;
 * ss << std::put_time(&tm, "%FT%TZ");
 * auto datetime = ss.str();
 *
 * // datetime now is a string with the datetime in the right format.
 *
 * ```
 */
void XLDocument::setProperty(XLProperty prop, const std::string& value) // NOLINT
{
    switch (prop) {
        case XLProperty::Application:
            m_appProperties.setProperty("Application", value);
            break;
        case XLProperty::AppVersion:    // ===== TODO: Clean up this section
            try {
                std::stof(value);
            }
            catch (...) {
                throw XLPropertyError("Invalid property value");
            }

            if (value.find('.') != std::string::npos) {
                if (!value.substr(value.find('.') + 1).empty() && value.substr(value.find('.') + 1).size() <= 5) { // NOLINT
                    if (!value.substr(0, value.find('.')).empty() && value.substr(0, value.find('.')).size() <= 2) {
                        m_appProperties.setProperty("AppVersion", value);
                    }
                    else
                        throw XLPropertyError("Invalid property value");
                }
                else
                    throw XLPropertyError("Invalid property value");
            }
            else
                throw XLPropertyError("Invalid property value");

            break;

        case XLProperty::Category:
            m_coreProperties.setProperty("cp:category", value);
            break;
        case XLProperty::Company:
            m_appProperties.setProperty("Company", value);
            break;
        case XLProperty::CreationDate:
            m_coreProperties.setProperty("dcterms:created", value);
            break;
        case XLProperty::Creator:
            m_coreProperties.setProperty("dc:creator", value);
            break;
        case XLProperty::Description:
            m_coreProperties.setProperty("dc:description", value);
            break;
        case XLProperty::DocSecurity:
            if (value == "0" || value == "1" || value == "2" || value == "4" || value == "8")
                m_appProperties.setProperty("DocSecurity", value);
            else
                throw XLPropertyError("Invalid property value");
            break;

        case XLProperty::HyperlinkBase:
            m_appProperties.setProperty("HyperlinkBase", value);
            break;
        case XLProperty::HyperlinksChanged:
            if (value == "true" || value == "false")
                m_appProperties.setProperty("HyperlinksChanged", value);
            else
                throw XLPropertyError("Invalid property value");

            break;

        case XLProperty::Keywords:
            m_coreProperties.setProperty("cp:keywords", value);
            break;
        case XLProperty::LastModifiedBy:
            m_coreProperties.setProperty("cp:lastModifiedBy", value);
            break;
        case XLProperty::LastPrinted:
            m_coreProperties.setProperty("cp:lastPrinted", value);
            break;
        case XLProperty::LinksUpToDate:
            if (value == "true" || value == "false")
                m_appProperties.setProperty("LinksUpToDate", value);
            else
                throw XLPropertyError("Invalid property value");
            break;

        case XLProperty::Manager:
            m_appProperties.setProperty("Manager", value);
            break;
        case XLProperty::ModificationDate:
            m_coreProperties.setProperty("dcterms:modified", value);
            break;
        case XLProperty::ScaleCrop:
            if (value == "true" || value == "false")
                m_appProperties.setProperty("ScaleCrop", value);
            else
                throw XLPropertyError("Invalid property value");
            break;

        case XLProperty::SharedDoc:
            if (value == "true" || value == "false")
                m_appProperties.setProperty("SharedDoc", value);
            else
                throw XLPropertyError("Invalid property value");
            break;

        case XLProperty::Subject:
            m_coreProperties.setProperty("dc:subject", value);
            break;
        case XLProperty::Title:
            m_coreProperties.setProperty("dc:title", value);
            break;
    }
}

/**
 * @details Delete a property
 */
void XLDocument::deleteProperty(XLProperty theProperty)
{
    setProperty(theProperty, "");
}

/**
 * @details
 */
void XLDocument::execCommand(const XLCommand& command) {

    switch (command.type()) {
        case XLCommandType::SetSheetName:
            m_appProperties.setSheetName(command.getParam<std::string>("sheetName"), command.getParam<std::string>("newName"));
            m_workbook.setSheetName(command.getParam<std::string>("sheetID"), command.getParam<std::string>("newName"));
            break;

        case XLCommandType::SetSheetColor:
            // TODO: To be implemented
            break;

        case XLCommandType::SetSheetVisibility:
            m_workbook.setSheetVisibility(command.getParam<std::string>("sheetID"), command.getParam<std::string>("sheetVisibility"));
            break;

        case XLCommandType::SetSheetIndex:
            {
                XLQuery qry(XLQueryType::QuerySheetName);
                auto sheetName = execQuery(qry.setParam("sheetID", command.getParam<std::string>("sheetID"))).result<std::string>();
                m_workbook.setSheetIndex(sheetName, command.getParam<uint16_t>("sheetIndex"));
            }
            break;

        case XLCommandType::SetSheetActive:
            m_workbook.setSheetActive(command.getParam<std::string>("sheetID"));
            break;

        case XLCommandType::ResetCalcChain:
            {
                m_archive.deleteEntry("xl/calcChain.xml");
                auto item = std::find_if(m_data.begin(), m_data.end(), [&](const XLXmlData& item) {
                    return item.getXmlPath() == "xl/calcChain.xml";
                });

                if (item != m_data.end()) m_data.erase(item);
            }
            break;
        case XLCommandType::AddSharedStrings:
            {
                m_contentTypes.addOverride("/xl/sharedStrings.xml", XLContentType::SharedStrings);
                m_wbkRelationships.addRelationship(XLRelationshipType::SharedStrings, "sharedStrings.xml");
                m_archive.addEntry("xl/sharedStrings.xml", XLTemplate::sharedStrings);
            }
            break;
        case XLCommandType::AddWorksheet:
            {
                auto internalID = availableFileID(XLContentType::Worksheet);
                std::string sheetPath = "xl/worksheets/sheet" + std::to_string(internalID) + ".xml";
                m_contentTypes.addOverride("/" + sheetPath, XLContentType::Worksheet);
                m_wbkRelationships.addRelationship(XLRelationshipType::Worksheet, sheetPath.substr(3));
                m_appProperties.appendSheetName(command.getParam<std::string>("sheetName"));
                m_archive.addEntry(sheetPath, XLTemplate::emptyWorksheet);
                m_data.emplace_back(
                    /* parentDoc */ this,
                    /* xmlPath   */ sheetPath,
                    /* xmlID     */ m_wbkRelationships.relationshipByTarget(sheetPath.substr(3)).id(),
                    /* xmlID     */ command.getParam<std::string>("sheetName"),               
                    /* xmlType   */ XLContentType::Worksheet);

                // TODO move elsewhere ?
                m_workbook.prepareSheetMetadata(command.getParam<std::string>("sheetName"), internalID);
            }
            break;
        case XLCommandType::AddChartsheet:
            // TODO: To be implemented
            break;
        case XLCommandType::AddTable:
            {
                //TODO worksheet !
                createTable(command.getParam<std::string>("worksheet"),
                            command.getParam<std::string>("tableName"),
                            command.getParam<std::string>("reference"));
                //get nexta available filename and next available id
                // check that the name dont already exist
            }
            // TODO: To be implemented
            break;
        case XLCommandType::DeleteSheet:
        {
            m_appProperties.deleteSheetName(command.getParam<std::string>("sheetName"));
            auto sheetPath = "/xl/" + m_wbkRelationships.relationshipById(command.getParam<std::string>("sheetID")).target();
            m_archive.deleteEntry(sheetPath.substr(1));
            m_contentTypes.deleteOverride(sheetPath);
            m_wbkRelationships.deleteRelationship(command.getParam<std::string>("sheetID"));
            m_data.erase(std::find_if(m_data.begin(), m_data.end(), [&](const XLXmlData& item) {
                return item.getXmlPath() == sheetPath.substr(1);
            }));
        }
            break;
        case XLCommandType::CloneSheet:
        {
            auto internalID = availableFileID(XLContentType::Worksheet);
            auto sheetPath  = "/xl/worksheets/sheet" + std::to_string(internalID) + ".xml";
            if (m_workbook.sheetExists(command.getParam<std::string>("cloneName")))
                throw XLInternalError("Sheet named \"" + command.getParam<std::string>("cloneName") + "\" already exists.");

            if (m_wbkRelationships.relationshipById(command.getParam<std::string>("sheetID")).type() == XLRelationshipType::Worksheet) {
                m_contentTypes.addOverride(sheetPath, XLContentType::Worksheet);
                m_wbkRelationships.addRelationship(XLRelationshipType::Worksheet, sheetPath.substr(4));
                m_appProperties.appendSheetName(command.getParam<std::string>("cloneName"));
                m_archive.addEntry(sheetPath.substr(1), std::find_if(m_data.begin(), m_data.end(), [&](const XLXmlData& data) {
                                                            return data.getXmlPath().substr(3) ==
                                                                   m_wbkRelationships.relationshipById(command.getParam<std::string>
                                                                       ("sheetID")).target();
                                                        })->getRawData());
                m_data.emplace_back(
                    /* parentDoc */ this,
                    /* xmlPath   */ sheetPath.substr(1),
                    /* xmlID     */ m_wbkRelationships.relationshipByTarget(sheetPath.substr(4)).id(),
                    /* name      */ command.getParam<std::string>("cloneName"),
                    /* xmlType   */ XLContentType::Worksheet);
            }
            else {
                m_contentTypes.addOverride(sheetPath, XLContentType::Chartsheet);
                m_wbkRelationships.addRelationship(XLRelationshipType::Chartsheet, sheetPath.substr(4));
                m_appProperties.appendSheetName(command.getParam<std::string>("cloneName"));
                m_archive.addEntry(sheetPath.substr(1), std::find_if(m_data.begin(), m_data.end(), [&](const XLXmlData& data) {
                                                            return data.getXmlPath().substr(3) ==
                                                                   m_wbkRelationships.relationshipById(command.getParam<std::string>
                                                                       ("sheetID")).target();
                                                        })->getRawData());
                m_data.emplace_back(
                    /* parentDoc */ this,
                    /* xmlPath   */ sheetPath.substr(1),
                    /* xmlID     */ m_wbkRelationships.relationshipByTarget(sheetPath.substr(4)).id(),
                    /* name      */ command.getParam<std::string>("cloneName"),
                    /* xmlType   */ XLContentType::Chartsheet);
            }

            m_workbook.prepareSheetMetadata(command.getParam<std::string>("cloneName"), internalID);
        }
            break;
    }
}

/**
 * @details
 */
XLQuery XLDocument::execQuery(const XLQuery& query) const
{

    switch (query.type()) {
        case XLQueryType::QuerySheetName:
            return XLQuery(query).setResult(m_workbook.sheetName(query.getParam<std::string>("sheetID")));

        case XLQueryType::QuerySheetIndex:
            return query;

        case XLQueryType::QuerySheetVisibility:
            return XLQuery(query).setResult(m_workbook.sheetVisibility(query.getParam<std::string>("sheetID")));

        case XLQueryType::QuerySheetType:
            if (m_wbkRelationships.relationshipById(query.getParam<std::string>("sheetID")).type() == XLRelationshipType::Worksheet)
                return XLQuery(query).setResult(XLContentType::Worksheet);
            else
                return XLQuery(query).setResult(XLContentType::Chartsheet);

        case XLQueryType::QuerySheetIsActive:
            return XLQuery(query).setResult(m_workbook.sheetIsActive(query.getParam<std::string>("sheetID")));

        case XLQueryType::QuerySheetID:
            return XLQuery(query).setResult(m_workbook.sheetVisibility(query.getParam<std::string>("sheetID")));

        case XLQueryType::QuerySheetRelsID:
            return XLQuery(query).setResult(m_wbkRelationships.relationshipByTarget(query.getParam<std::string>("sheetPath").substr(4)).id());

        case XLQueryType::QuerySheetRelsTarget:
            return XLQuery(query).setResult(m_wbkRelationships.relationshipById(query.getParam<std::string>("sheetID")).target());

        case XLQueryType::QuerySharedStrings:
            return XLQuery(query).setResult(&m_sharedStrings);
            //return XLQuery(query).setResult(1);
        
        case XLQueryType::QuerySheetFromName:
            return XLQuery(query).setResult(getXmlDataByName(query.getParam<std::string>("sheetName")));
                
        case XLQueryType::QueryTableFromName:
            {
                for (auto& item : m_data){
                    if(item.getXmlType() == XLContentType::Table){
                        if (item.getXmlDocument()->child("table").attribute("name").value() == query.getParam<std::string>("tableName"))
                            return XLQuery(query).setResult(&item);
                    }
                }
                throw XLInternalError("Table does not exist in zip archive (" + query.getParam<std::string>("tableName") + ")");

                return query; // Table not found
            }
        case XLQueryType::QueryXmlData:
            {
                auto result =
                    std::find_if(m_data.begin(), m_data.end(), [&](const XLXmlData& item) { return item.getXmlPath() == query.getParam<std::string>("xmlPath"); });
                if (result == m_data.end()) throw XLInternalError("Path does not exist in zip archive (" + query.getParam<std::string>("xmlPath") + ")");
                return XLQuery(query).setResult(&*result);
            }
    }

    return query; // Needed in order to suppress compiler warning
}

/**
 * @details
 */
XLQuery XLDocument::execQuery(const XLQuery& query)
{
    return static_cast<const XLDocument&>(*this).execQuery(query);
}

/**
 * @details
 */
XLDocument::operator bool() const
{
    return !!m_archive; // NOLINT
}

/**
 * @details
 */
bool XLDocument::isOpen() const {
    return this->operator bool();
}

/**
 * @details
 */
XLXmlData* XLDocument::getXmlDataByPath(const std::string& path)
{
    if (!hasXmlData(path)) throw XLInternalError("Path does not exist in zip archive.");
    return &*std::find_if(m_data.begin(), m_data.end(), [&](const XLXmlData& item) { return item.getXmlPath() == path; });
//    auto result = std::find_if(m_data.begin(), m_data.end(), [&](const XLXmlData& item) { return item.getXmlPath() == path; });
//    if (result == m_data.end()) throw XLInternalError("Path does not exist in zip archive.");
//    return &*result;
}

/**
 * @details
 */
const XLXmlData* XLDocument::getXmlDataByPath(const std::string& path) const
{
    if (!hasXmlData(path)) throw XLInternalError("Path does not exist in zip archive.");
    return &*std::find_if(m_data.begin(), m_data.end(), [&](const XLXmlData& item) { return item.getXmlPath() == path; });
//    if (result == m_data.end()) throw XLInternalError("Path does not exist in zip archive.");
//    return &*result;
}

/**
 * @details
 */
bool XLDocument::hasXmlData(const std::string& path) const
{
    return std::find_if(m_data.begin(), m_data.end(), [&](const XLXmlData& item) { return item.getXmlPath() == path; }) != m_data.end();
}

/**
 * @details
 */
std::string XLDocument::extractXmlFromArchive(const std::string& path)
{
    return (m_archive.hasEntry(path) ? m_archive.getEntry(path) : "");
}

XLXmlData* XLDocument::getXmlDataByName(const std::string& name) const
{
    // for(auto objXmlData : m_XmlWorkbook->getChildNodes())
    for(auto& objXmlData : m_data)
        if(objXmlData.getName() == name)
            return &objXmlData;

    return nullptr;
}

XLSharedStrings& XLDocument::getSharedString() const
{
    return m_sharedStrings;
}

std::string XLDocument::getSheetRelsPath(const std::string& sheetName) const
{
    XLXmlData* wsItem = getXmlDataByName(sheetName);
    if (wsItem == nullptr)  // sheetName does not exist
        return std::string();
    
    std::string sheetPath = wsItem->getXmlPath();
    std::string::size_type n = sheetPath.find_last_of('/');

    if(n>sheetPath.size()) // path does not contains '/'
        return std::string();
    
    // Some spreadsheets use absolute rather than relative paths in relationship items.
    if (sheetPath.substr(0,4) == "/xl/") 
        sheetPath = "xl/" + sheetPath.substr(4);

    const std::string basePath = sheetPath.substr(0, n);
    const std::string fileName = sheetPath.substr(n);

    if (fileName.substr(0, 6) != "/sheet") // This is not a sheet path
        return std::string();
    
    return basePath + "/_rels" + fileName + ".rels";
}

uint16_t XLDocument::availableFileID(XLContentType type)
{
    std::vector<uint32_t> fileIndices;

    for (auto& wsItem : m_data)
        if (wsItem.getXmlType() == type){
            std::string path = wsItem.getXmlPath();
            std::string::size_type n = path.find_last_of("/");
            // TODO this works because "sheet", "table" and "chart" have same length
            fileIndices.push_back(std::stoi(path.substr(n+6,path.size()-n-10))); // removing "table and .xml"
        }

    // Get the first available indice for file and id (filling missings)
    std::sort (fileIndices.begin(), fileIndices.end());
    uint32_t nFile = 1;
    for(uint32_t i : fileIndices)
        if (i == nFile)
            nFile +=1;
    
    return nFile;
}

/**
 * @todo check if the reference overlap an existing table
 */
void XLDocument::createTable(const std::string& sheetName, const std::string& tableName, const std::string& reference)
{  
    XLXmlData* wks = getXmlDataByName(sheetName);
   
    std::vector<uint32_t> idIndices;

    std::string basePath = "xl/tables/";
    std::string name = tableName;
    // Loop through existing tables to find the available:
    // - filename
    // - index
    // - check name
    for (auto& wsItem : m_data)
        if (wsItem.getXmlType() == XLContentType::Table){
            std::string path = wsItem.getXmlPath();
            std::string::size_type n = path.find_last_of("/");

            XMLNode tableNode = wsItem.getXmlDocument()->child("table");
            idIndices.push_back(std::stoi(tableNode.attribute("id").value()));
            if (( name == tableNode.attribute("name").value()) ||
                ( name == tableNode.attribute("displayName").value()))
                name += INCREMENT_STRING;
        }

    std::sort (idIndices.begin(), idIndices.end());
    uint32_t nId = 1;
    for(uint32_t i : idIndices)
        if (i == nId)
            nId +=1;
    
    std::string fileTable = "table" 
                    + std::to_string(availableFileID(XLContentType::Table)) 
                    + ".xml";
    std::string filePath = basePath + fileTable;  

    std::string sheetRelsPath = getSheetRelsPath(sheetName);
    if (sheetRelsPath.empty()) // sheetname does not exist or error in the path
        return;
    

    if (!m_archive.hasEntry(sheetRelsPath)){ // if the file does not exists xl/worksheets/_rels/sheet{0}.xml.rels
        m_archive.addEntry(sheetRelsPath, XLTemplate::emptySheetRels); // add the table{0}.xml file
        m_data.emplace_back(
            /* parentDoc */ this,
            /* xmlPath   */ sheetRelsPath,
            /* xmlID     */ std::string(),
            /* name      */ std::string(),               
            /* xmlType   */ XLContentType::WorksheetRelations,
            /* parentNode*/ wks);
        wks->addChildNode(getXmlDataByPath(sheetRelsPath));
    }

    // adding Rels
    auto sheetRels = XLRelationships(getXmlDataByPath(sheetRelsPath));
    auto relItem = sheetRels.addRelationship(XLRelationshipType::Table, "../tables/" + fileTable);

    // adding the file table in content, create it and register it
    m_contentTypes.addOverride("/" + filePath, XLContentType::Table); // add to contentTypes
    m_archive.addEntry(filePath, XLTemplate::emptyTable); // add the table{0}.xml file
    
    m_data.emplace_back(
            /* parentDoc */ this,
            /* xmlPath   */ filePath,
            /* xmlID     */ relItem.id(),
            /* name      */ name,               
            /* xmlType   */ XLContentType::Table,
            /* parentNode*/ wks);
    XLXmlData* tableXml = getXmlDataByName(name);
    wks->addChildNode(tableXml);
    
    // add tablePart in sheet{0}.xml
    XMLNode tableParts = wks->getXmlDocument()->document_element().child("tableParts");
    if(!tableParts){
        tableParts = wks->getXmlDocument()->child("worksheet").append_child("tableParts"); 
        tableParts.append_attribute("count").set_value("0");
    }

    tableParts.attribute("count")
                .set_value(std::to_string(std::stoi(
                tableParts.attribute("count").value()) + 1 ).c_str());
    XMLNode newPart = tableParts.append_child("tablePart");
    newPart.append_attribute("r:id").set_value(relItem.id().c_str());

    // Prepare table{0}.xml
    XMLDocument* tableDoc = tableXml->getXmlDocument();
    XMLNode tableNode = tableDoc->child("table");

    // Get safe ref
    auto pair = XLCellRange::topLeftBottomRight(reference);
    auto topLeft =  XLCellReference(pair.first);
    auto bottomRight = XLCellReference(pair.second);
    std::string ref = topLeft.address(false) + ":" + bottomRight.address(false);


    tableNode.attribute("id").set_value(std::to_string(nId).c_str());
    tableNode.attribute("name").set_value(name.c_str());
    tableNode.attribute("displayName").set_value(name.c_str());
    tableNode.attribute("ref").set_value(ref.c_str());

    // Set up the columns !!!
    const XLWorksheet* pWks = new XLWorksheet(wks);
    auto topRight = XLCellReference(topLeft.row(),bottomRight.column());
    auto headerRange = XLCellRange(wks->getXmlDocument()->first_child().child("sheetData"),
                                    topLeft,
                                    topRight,
                                    pWks);
    
    // Setup all the columns name
    XMLNode columnsNode = tableNode.child("tableColumns");
    uint16_t colId = 1;
    std::vector<std::string> colNames;

    for(auto& cell : headerRange){
        std::string colName = cell.value();
        if (colName.empty())
            colName = "Column";
        bool notValid = true;
        while (notValid){
            if (std::find(colNames.begin(), colNames.end(), colName) != colNames.end())
                colName += INCREMENT_STRING;
            else
                notValid = false;
        }
        colNames.push_back(colName);
        cell.value() = colName;
        // Fill shared string
        // Fill cell value !
        auto newNode = columnsNode.append_child("tableColumn");
        newNode.append_attribute("id").set_value(std::to_string(colId).c_str());
        newNode.append_attribute("name").set_value(colName.c_str());
        colId +=1;
    }
    // adding the table columns count
    columnsNode.attribute("count").set_value(std::to_string(colId-1).c_str());

    // TODEL
    delete pWks;
    
}

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
#include <algorithm>
#include <iterator>
#include <pugixml.hpp>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"
#include "XLSheet.hpp"
#include "XLWorkbook.hpp"

using namespace OpenXLSX;

namespace
{
    /**
     * @brief
     * @param doc
     * @return
     */
    XMLNode sheetsNode(const XMLDocument& doc)
    {
        return doc.document_element().child("sheets");
    }
}    // namespace

/**
 * @details The constructor initializes the member variables and calls the loadXMLData from the
 * XLAbstractXMLFile base class.
 */
XLWorkbook::XLWorkbook(XLXmlData* xmlData) : XLXmlFile(xmlData) {}

/**
 * @details
 */
XLWorkbook::~XLWorkbook() = default;

/**
 * @details
 */
XLSheet XLWorkbook::sheet(const std::string& sheetName)
{
    // ===== First determine if the sheet exists.
    if (xmlDocument().document_element().child("sheets").find_child_by_attribute("name", sheetName.c_str()) == nullptr)
        throw XLException("Sheet \"" + sheetName + "\" does not exist");

    // ===== Find the sheet data corresponding to the sheet with the requested name
    std::string xmlID =
        xmlDocument().document_element().child("sheets").find_child_by_attribute("name", sheetName.c_str()).attribute("r:id").value();
    std::string xmlPath = parentDoc().executeQuery(XLQuerySheetRelsTarget(xmlID)).sheetTarget();
    return XLSheet(parentDoc().executeQuery(XLQueryXmlData("xl/" + xmlPath)).xmlData());
}

/**
 * @details Create a vector with sheet nodes, retrieve the node at the requested index, get sheet name and return the
 * corresponding sheet object.
 */
XLSheet XLWorkbook::sheet(uint16_t index)
{
    if (index < 1 || index > sheetCount()) throw XLException("Sheet index is out of bounds");
    return sheet(
        std::vector<XMLNode>(sheetsNode(xmlDocument()).begin(), sheetsNode(xmlDocument()).end())[index - 1].attribute("name").as_string());
}

/**
 * @details
 */
XLWorksheet XLWorkbook::worksheet(const std::string& sheetName)
{
    return sheet(sheetName).get<XLWorksheet>();
}

/**
 * @details
 */
XLChartsheet XLWorkbook::chartsheet(const std::string& sheetName)
{
    return sheet(sheetName).get<XLChartsheet>();
}

/**
 * @details
 */
bool XLWorkbook::hasSharedStrings() const
{
    return parentDoc().executeQuery(XLQuerySharedStrings()).sharedStrings() != nullptr;
}

/**
 * @details
 */
XLSharedStrings* XLWorkbook::sharedStrings()
{
    return parentDoc().executeQuery(XLQuerySharedStrings()).sharedStrings();
}

/**
 * @details
 */
void XLWorkbook::deleteNamedRanges()
{
    for (auto& child : xmlDocument().document_element().child("definedNames").children()) child.parent().remove_child(child);
}

/**
 * @details
 */
void XLWorkbook::deleteSheet(const std::string& sheetName)
{
    // ===== Determine ID and type of sheet, as well as current worksheet count.
    auto sheetID   = sheetsNode(xmlDocument()).find_child_by_attribute("name", sheetName.c_str()).attribute("r:id").value();    // NOLINT
    auto sheetType = parentDoc().executeQuery(XLQuerySheetType(relationshipID())).sheetType();
    auto worksheetCount =
        std::count_if(sheetsNode(xmlDocument()).children().begin(), sheetsNode(xmlDocument()).children().end(), [&](const XMLNode& item) {
            return parentDoc().executeQuery(XLQuerySheetType(item.attribute("r:id").value())).sheetType() == XLContentType::Worksheet;
        });

    // ===== If this is the last worksheet in the workbook, throw an exception.
    if (worksheetCount == 1 && sheetType == XLContentType::Worksheet)
        throw XLException("Invalid operation. There must be at least one worksheet in the workbook.");

    // ===== Delete the sheet data as well as the sheet node from Workbook.xml
    parentDoc().executeCommand(XLCommandDeleteSheet(sheetID, sheetName));
    sheetsNode(xmlDocument()).remove_child(sheetsNode(xmlDocument()).find_child_by_attribute("name", sheetName.c_str()));

    // TODO: The 'activeSheet' property may need to be updated.
}

/**
 * @details
 */
void XLWorkbook::addWorksheet(const std::string& sheetName)
{
    // ===== If a sheet with the given name already exists, throw an exception.
    if (xmlDocument().document_element().child("sheets").find_child_by_attribute("name", sheetName.c_str()))
        throw XLException("Sheet named \"" + sheetName + "\" already exists.");

    // ===== Create new internal (workbook) ID for the sheet
    auto internalID = createInternalSheetID();

    // ===== Create xml file for new worksheet and add metadata to the workbook file.
    parentDoc().executeCommand(XLCommandAddWorksheet(sheetName, "/xl/worksheets/sheet" + std::to_string(internalID) + ".xml"));
    prepareSheetMetadata(sheetName, internalID);
}

/**
 * @details
 * @todo If the original sheet's tabSelected attribute is set, ensure it is un-set in the clone.
 */
void XLWorkbook::cloneSheet(const std::string& existingName, const std::string& newName)
{
    parentDoc().executeCommand(XLCommandCloneSheet(sheetID(existingName), newName));
}

/**
 * @details
 */
uint16_t XLWorkbook::createInternalSheetID()
{
    return static_cast<uint16_t>(std::max_element(xmlDocument().document_element().child("sheets").children().begin(),
                                                  xmlDocument().document_element().child("sheets").children().end(),
                                                  [](const XMLNode& a, const XMLNode& b) {
                                                      return a.attribute("sheetId").as_uint() < b.attribute("sheetId").as_uint();
                                                  })
                                     ->attribute("sheetId")
                                     .as_uint() +
                                 1);
}

/**
 * @details
 */
std::string XLWorkbook::sheetID(const std::string& sheetName)
{
    return xmlDocument().document_element().child("sheets").find_child_by_attribute("name", sheetName.c_str()).attribute("r:id").value();
}

/**
 * @details
 */
std::string XLWorkbook::sheetName(const std::string& sheetID) const
{
    return xmlDocument().document_element().child("sheets").find_child_by_attribute("r:id", sheetID.c_str()).attribute("name").value();
}

/**
 * @details
 */
std::string XLWorkbook::sheetVisibility(const std::string& sheetID) const
{
    return xmlDocument().document_element().child("sheets").find_child_by_attribute("r:id", sheetID.c_str()).attribute("state").value();
}

/**
 * @details
 */
void XLWorkbook::prepareSheetMetadata(const std::string& sheetName, uint16_t internalID)
{
    // ===== Add new child node to the "sheets" node.
    auto node = sheetsNode(xmlDocument()).append_child("sheet");

    // ===== append the required attributes to the newly created sheet node.
    std::string sheetPath            = "/xl/worksheets/sheet" + std::to_string(internalID) + ".xml";
    node.append_attribute("name")    = sheetName.c_str();
    node.append_attribute("sheetId") = std::to_string(internalID).c_str();
    node.append_attribute("r:id")    = parentDoc().executeQuery(XLQuerySheetRelsID(sheetPath)).sheetID().c_str();
}

/**
 * @details
 */
void XLWorkbook::setSheetName(const std::string& sheetRID, const std::string& newName)
{
    auto sheetName = xmlDocument().document_element().child("sheets").find_child_by_attribute("r:id", sheetRID.c_str()).attribute("name");

    updateSheetReferences(sheetName.value(), newName);
    sheetName.set_value(newName.c_str());
}

/**
 * @details
 */
void XLWorkbook::setSheetVisibility(const std::string& sheetRID, const std::string& state)
{
    // ===== Retrieve or create the visibility ("state") attribute for the sheet.
    auto stateAttribute =
        xmlDocument().document_element().child("sheets").find_child_by_attribute("r:id", sheetRID.c_str()).attribute("state");
    if (!stateAttribute) {
        stateAttribute =
            xmlDocument().document_element().child("sheets").find_child_by_attribute("r:id", sheetRID.c_str()).append_attribute("state");
    }

    // ===== Set the visibility attribute.
    stateAttribute.set_value(state.c_str());
}

/**
 * @details
 */
void XLWorkbook::setSheetIndex(const std::string& sheetName, unsigned int index)
{
    // ===== Check that the input is valid
    //    if (index < 1 || index > std::distance(xmlDocument().document_element().child("sheets").children().begin(),
    //                                           xmlDocument().document_element().child("sheets").children().end()))
    //        throw XLException("Invalid sheet index");

    // ===== If the new index is equal to the current, don't do anything
    if (index == std::distance(xmlDocument().document_element().child("sheets").children().begin(),
                               std::find_if(xmlDocument().document_element().child("sheets").children().begin(),
                                            xmlDocument().document_element().child("sheets").children().end(),
                                            [&](const XMLNode& item) { return sheetName == item.attribute("name").value(); })))
        return;

    // ===== Modify the node in the XML file
    if (index <= 1)
        sheetsNode(xmlDocument()).prepend_move(sheetsNode(xmlDocument()).find_child_by_attribute("name", sheetName.c_str()));
    else if (index >= sheetCount())
        sheetsNode(xmlDocument()).append_move(sheetsNode(xmlDocument()).find_child_by_attribute("name", sheetName.c_str()));
    else {
        auto vec           = std::vector<XMLNode>(sheetsNode(xmlDocument()).children().begin(), sheetsNode(xmlDocument()).children().end());
        auto existingSheet = vec[index - 1];
        if (indexOfSheet(sheetName) < index) {
            sheetsNode(xmlDocument())
                .insert_move_after(sheetsNode(xmlDocument()).find_child_by_attribute("name", sheetName.c_str()), existingSheet);
        }
        else if (indexOfSheet(sheetName) > index) {
            sheetsNode(xmlDocument())
                .insert_move_before(sheetsNode(xmlDocument()).find_child_by_attribute("name", sheetName.c_str()), existingSheet);
        }
    }

    // ===== Updated defined names with worksheet scopes.
    for (auto& definedName : xmlDocument().document_element().child("definedNames").children()) {
        definedName.attribute("localSheetId").set_value(indexOfSheet(sheetName) - 1);
    }

    // ===== Update the activeTab attribute.
    //        unsigned int index = 0;
    //        for (auto& item : getSheetsNode().children()) {
    //            if (m_activeSheet == item) {
    //                XmlDocument().first_child().child("bookViews").first_child().attribute("activeTab").set_value(index);
    //                break;
    //            }
    //            index++;
    //        }
}

/**
 * @details
 */
unsigned int XLWorkbook::indexOfSheet(const std::string& sheetName) const
{
    // ===== Iterate through sheet nodes. When a match is found, return the index;
    unsigned int index = 1;
    for (auto& sheet : sheetsNode(xmlDocument()).children()) {
        if (sheetName == sheet.attribute("name").value()) return index;
        index++;
    }

    // ===== If a match is not found, throw an exception.
    throw XLException("Sheet does not exist");
}

/**
 * @details
 */
XLSheetType XLWorkbook::typeOfSheet(const std::string& sheetName) const
{
    if (!sheetExists(sheetName)) throw XLException("Sheet with name \"" + sheetName + "\" doesn't exist.");

    if (worksheetExists(sheetName))
        return XLSheetType::Worksheet;
    else
        return XLSheetType::Chartsheet;
}

/**
 * @details
 */
XLSheetType XLWorkbook::typeOfSheet(unsigned int index) const
{
    std::string name =
        std::vector<XMLNode>(sheetsNode(xmlDocument()).begin(), sheetsNode(xmlDocument()).end())[index - 1].attribute("name").as_string();
    return typeOfSheet(name);
}

/**
 * @details
 */
unsigned int XLWorkbook::sheetCount() const
{
    return static_cast<unsigned int>(
        std::distance(sheetsNode(xmlDocument()).children().begin(), sheetsNode(xmlDocument()).children().end()));
}

/**
 * @details
 */
unsigned int XLWorkbook::worksheetCount() const
{
    return static_cast<unsigned int>(worksheetNames().size());
}

/**
 * @details
 */
unsigned int XLWorkbook::chartsheetCount() const
{
    return static_cast<unsigned int>(chartsheetNames().size());
}

/**
 * @details
 */
std::vector<std::string> XLWorkbook::sheetNames() const
{
    std::vector<std::string> results;

    for (const auto& item : sheetsNode(xmlDocument()).children()) results.emplace_back(item.attribute("name").value());

    return results;
}

/**
 * @details
 */
std::vector<std::string> XLWorkbook::worksheetNames() const
{
    std::vector<std::string> results;

    for (const auto& item : sheetsNode(xmlDocument()).children()) {
        if (parentDoc().executeQuery(XLQuerySheetType(item.attribute("r:id").value())).sheetType() == XLContentType::Worksheet)
            results.emplace_back(item.attribute("name").value());
    }

    return results;
}

/**
 * @details
 */
std::vector<std::string> XLWorkbook::chartsheetNames() const
{
    std::vector<std::string> results;

    for (const auto& item : sheetsNode(xmlDocument()).children()) {
        if (parentDoc().executeQuery(XLQuerySheetType(item.attribute("r:id").value())).sheetType() == XLContentType::Chartsheet)
            results.emplace_back(item.attribute("name").value());
    }

    return results;
}

/**
 * @details
 */
bool XLWorkbook::sheetExists(const std::string& sheetName) const
{
    return chartsheetExists(sheetName) || worksheetExists(sheetName);
}

/**
 * @details
 */
bool XLWorkbook::worksheetExists(const std::string& sheetName) const
{
    auto wksNames = worksheetNames();
    return std::find(wksNames.begin(), wksNames.end(), sheetName) != wksNames.end();
}

/**
 * @details
 */
bool XLWorkbook::chartsheetExists(const std::string& sheetName) const
{
    auto chsNames = chartsheetNames();
    return std::find(chsNames.begin(), chsNames.end(), sheetName) != chsNames.end();
}

/**
 * @details The UpdateSheetName member function searches theroug the usages of the old name and replaces with the
 * new sheet name.
 * @todo Currently, this function only searches through defined names. Consider using this function to update the
 * actual sheet name as well.
 */
void XLWorkbook::updateSheetReferences(const std::string& oldName, const std::string& newName)
{
    //        for (auto& sheet : m_sheets) {
    //            if (sheet.sheetType == XLSheetType::WorkSheet)
    //                Worksheet(sheet.sheetNode.attribute("name").value())->UpdateSheetName(oldName, newName);
    //        }

    // ===== Set up temporary variables
    std::string oldNameTemp = oldName;
    std::string newNameTemp = newName;
    std::string formula;

    // ===== If the sheet name contains spaces, it should be enclosed in single quotes (')
    if (oldName.find(' ') != std::string::npos) oldNameTemp = "\'" + oldName + "\'";
    if (newName.find(' ') != std::string::npos) newNameTemp = "\'" + newName + "\'";

    // ===== Ensure only sheet names are replaced (references to sheets always ends with a '!')
    oldNameTemp += '!';
    newNameTemp += '!';

    // ===== Iterate through all defined names
    for (auto& definedName : xmlDocument().document_element().child("definedNames").children()) {
        formula = definedName.text().get();

        // ===== Skip if formula contains a '[' and ']' (means that the defined refers to external workbook)
        if (formula.find('[') == std::string::npos && formula.find(']') == std::string::npos) {
            // ===== For all instances of the old sheet name in the formula, replace with the new name.
            while (formula.find(oldNameTemp) != std::string::npos) {
                formula.replace(formula.find(oldNameTemp), oldNameTemp.length(), newNameTemp);
            }
            definedName.text().set(formula.c_str());
        }
    }
}

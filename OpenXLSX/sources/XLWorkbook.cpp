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
#include "utilities/XLUtilities.hpp"

using namespace OpenXLSX;

namespace
{
    /**
     * @brief
     * @param doc
     * @return
     */
    XMLNode sheetsNode(const XMLDocument& doc) { return doc.document_element().child("sheets"); }
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
        throw XLInputError("Sheet \"" + sheetName + "\" does not exist");

    // ===== Find the sheet data corresponding to the sheet with the requested name
    const std::string xmlID =
        xmlDocument().document_element().child("sheets").find_child_by_attribute("name", sheetName.c_str()).attribute("r:id").value();

    XLQuery pathQuery(XLQueryType::QuerySheetRelsTarget);
    pathQuery.setParam("sheetID", xmlID);
    auto xmlPath = parentDoc().execQuery(pathQuery).result<std::string>();

    // Some spreadsheets use absolute rather than relative paths in relationship items.
    if (xmlPath.substr(0, 4) == "/xl/") xmlPath = xmlPath.substr(4);

    XLQuery xmlQuery(XLQueryType::QueryXmlData);
    xmlQuery.setParam("xmlPath", "xl/" + xmlPath);
    return XLSheet(parentDoc().execQuery(xmlQuery).result<XLXmlData*>());
}

/**
 * @details iterate over sheetsNode and count element nodes until index, get sheet name and return the corresponding sheet object
 *
 * @old-details Create a vector with sheet nodes, retrieve the node at the requested index, get sheet name and return the
 * corresponding sheet object.
 */
XLSheet XLWorkbook::sheet(uint16_t index)    // 2024-04-30: whitespace support
{
    if (index < 1 || index > sheetCount()) throw XLInputError("Sheet index is out of bounds");

    // ===== Find the n-th node_element that corresponds to index
    uint16_t curIndex = 0;
    for (XMLNode node = sheetsNode(xmlDocument()).first_child_of_type(pugi::node_element);
         not node.empty();
         node         = node.next_sibling_of_type(pugi::node_element))
    {
        if (++curIndex == index) return sheet(node.attribute("name").as_string());
    }

    // ===== If execution gets here, there are less element nodes than index in sheetsNode, this should never happen
    using namespace std::literals::string_literals;
    throw XLInternalError("Sheet index "s + std::to_string(index) + " is out of bounds"s);
}

/**
 * @details
 */
XLWorksheet XLWorkbook::worksheet(const std::string& sheetName) { return sheet(sheetName).get<XLWorksheet>(); }

/**
 * @details
 */
XLWorksheet XLWorkbook::worksheet(uint16_t index) { return sheet(index).get<XLWorksheet>(); }

/**
 * @details
 */
XLChartsheet XLWorkbook::chartsheet(const std::string& sheetName) { return sheet(sheetName).get<XLChartsheet>(); }

/**
 * @details
 */
XLChartsheet XLWorkbook::chartsheet(uint16_t index) { return sheet(index).get<XLChartsheet>(); }

/**
 * @details
 */
bool XLWorkbook::hasSharedStrings() const
{
    return true;    // parentDoc().executeQuery(XLQuerySharedStrings()).sharedStrings() != nullptr;
}

/**
 * @details
 */
XLSharedStrings XLWorkbook::sharedStrings()
{
    const XLQuery query(XLQueryType::QuerySharedStrings);
    return parentDoc().execQuery(query).result<XLSharedStrings>();
}

/**
 * @details
 */
void XLWorkbook::deleteNamedRanges()
{
    xmlDocument()
        .document_element()
        .child("definedNames")
        .remove_children();    // 2024-05-02: why loop if pugixml has a function for "delete all children"?
    // for (auto& child : xmlDocument().document_element().child("definedNames").children()) child.parent().remove_child(child);
}

/**
 * @details
 */
void XLWorkbook::deleteSheet(const std::string& sheetName)    // 2024-05-02: whitespace support
                                                              // CAUTION: execCommand on underlying XML with whitespaces not verified
{
    // ===== Determine ID and type of sheet, as well as current worksheet count.
    auto    sheetID = sheetsNode(xmlDocument()).find_child_by_attribute("name", sheetName.c_str()).attribute("r:id").value();    // NOLINT
    XLQuery sheetTypeQuery(XLQueryType::QuerySheetType);
    sheetTypeQuery.setParam("sheetID", std::string(sheetID));    // BUGFIX 2024-05-02: was using relationshipID() instead of sheetID,
                                                                 // leading to a bad sheetType & a failed check to not delete last worksheet
    const auto sheetType = parentDoc().execQuery(sheetTypeQuery).result<XLContentType>();

    // 2024-04-30: this should be whitespace-safe due to lambda expression checking for an attribute that non element nodes can not have
    const auto worksheetCount =
        std::count_if(sheetsNode(xmlDocument()).children().begin(), sheetsNode(xmlDocument()).children().end(), [&](const XMLNode& item) {
            if (item.type() != pugi::node_element) return false;    // 2024-05-02 this avoids the unnecessary query below
            XLQuery query(XLQueryType::QuerySheetType);
            query.setParam("sheetID", std::string(item.attribute("r:id").value()));
            return parentDoc().execQuery(query).result<XLContentType>() == XLContentType::Worksheet;
        });

    // ===== If this is the last worksheet in the workbook, throw an exception.
    if (worksheetCount == 1 && sheetType == XLContentType::Worksheet)
        throw XLInputError("Invalid operation. There must be at least one worksheet in the workbook.");

    // ===== Delete the sheet data as well as the sheet node from Workbook.xml
    parentDoc().execCommand(
        XLCommand(XLCommandType::DeleteSheet).setParam("sheetID", std::string(sheetID)).setParam("sheetName", sheetName));
    XMLNode sheet = sheetsNode(xmlDocument()).find_child_by_attribute("name", sheetName.c_str());
    if (not sheet.empty()) {
        // ===== Delete all non element nodes (comments, whitespaces) following the sheet being deleted from workbook.xml <sheets> node
        XMLNode nonElementNode = sheet.next_sibling();
        while (not nonElementNode.empty() && nonElementNode.type() != pugi::node_element) {
            sheetsNode(xmlDocument()).remove_child(nonElementNode);
            nonElementNode = nonElementNode.next_sibling();
        }
        sheetsNode(xmlDocument()).remove_child(sheet);    // delete the actual sheet entry
    }

    if (sheetIsActive(sheetID))
        xmlDocument().document_element().child("bookViews").first_child_of_type(pugi::node_element).remove_attribute("activeTab");
}

/**
 * @details
 */
void XLWorkbook::addWorksheet(const std::string& sheetName)
{
    // ===== If a sheet with the given name already exists, throw an exception.
    if (xmlDocument().document_element().child("sheets").find_child_by_attribute("name", sheetName.c_str()))
        throw XLInputError("Sheet named \"" + sheetName + "\" already exists.");

    // ===== Create new internal (workbook) ID for the sheet
    auto internalID = createInternalSheetID();

    // ===== Create xml file for new worksheet and add metadata to the workbook file.
    parentDoc().execCommand(XLCommand(XLCommandType::AddWorksheet)
                                .setParam("sheetName", sheetName)
                                .setParam("sheetPath", "/xl/worksheets/sheet" + std::to_string(internalID) + ".xml"));
    prepareSheetMetadata(sheetName, internalID);
}

/**
 * @details
 * @todo If the original sheet's tabSelected attribute is set, ensure it is un-set in the clone.
 *       TBD: See comment in XLWorkbook::setSheetActive - should the tabSelected actually be un-set? It's not the same as the active tab,
 *        which does not need to be selected
 */
void XLWorkbook::cloneSheet(const std::string& existingName, const std::string& newName)
{
    parentDoc().execCommand(XLCommand(XLCommandType::CloneSheet).setParam("sheetID", sheetID(existingName)).setParam("cloneName", newName));
}

/**
 * @details
 */
uint16_t XLWorkbook::createInternalSheetID()    // 2024-04-30: whitespace support
{
    XMLNode  sheet           = xmlDocument().document_element().child("sheets").first_child_of_type(pugi::node_element);
    uint32_t maxSheetIdFound = 0;
    while (not sheet.empty()) {
        uint32_t thisSheetId = sheet.attribute("sheetId").as_uint();
        if (thisSheetId > maxSheetIdFound) maxSheetIdFound = thisSheetId;
        sheet = sheet.next_sibling_of_type(pugi::node_element);
    }
    return maxSheetIdFound + 1;
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

    XLQuery query(XLQueryType::QuerySheetRelsID);
    query.setParam("sheetPath", sheetPath);
    node.append_attribute("r:id") = parentDoc().execQuery(query).result<std::string>().c_str();
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
void XLWorkbook::setSheetVisibility(const std::string& sheetRID, const std::string& state)    // 2024-04-30: whitespace support
{
    // ===== First, determine if there are other sheets that are visible
    int visibleSheets = 0;
    // for (const auto& item : xmlDocument().document_element().child("sheets").children()) {
    XMLNode item = xmlDocument().document_element().child("sheets").first_child_of_type(pugi::node_element);
    while (not item.empty()) {
        if (std::string(item.attribute("r:id").value()) != sheetRID) {
            if (isVisible(item)) ++visibleSheets;
        }
        item = item.next_sibling_of_type(pugi::node_element);
    }

    bool hideSheet = not isVisibleState(state);    // 2024-04-30: save for later use on activating another sheet if needed

    // ===== If there are no other visible sheets, and the current sheet is to be made hidden, throw an exception.
    if (hideSheet && visibleSheets == 0) throw XLSheetError("At least one sheet must be visible.");

    // ===== Then, retrieve or create the visibility ("state") attribute for the sheet, and set it to the "state" value
    auto stateAttribute =
        xmlDocument().document_element().child("sheets").find_child_by_attribute("r:id", sheetRID.c_str()).attribute("state");
    if (stateAttribute.empty()) {
        stateAttribute =
            xmlDocument().document_element().child("sheets").find_child_by_attribute("r:id", sheetRID.c_str()).prepend_attribute("state");
    }
    stateAttribute.set_value(state.c_str());

    // Next, find the index of the sheet...
    std::string name =
        xmlDocument().document_element().child("sheets").find_child_by_attribute("r:id", sheetRID.c_str()).attribute("name").value();
    auto index = indexOfSheet(name) - 1;    // 2024-05-01: activeTab property stores an array index, NOT the value of r_ID?

    // ...and determine the index of the active sheet
    XMLNode workbookView       = xmlDocument().document_element().child("bookViews").first_child_of_type(pugi::node_element);
    auto    activeTabAttribute = workbookView.attribute("activeTab");
    if (activeTabAttribute.empty()) {
        activeTabAttribute = workbookView.append_attribute("activeTab");
        activeTabAttribute.set_value(0);
    }
    const auto activeTabIndex = activeTabAttribute.as_uint();

    // Finally, if the current sheet is the active one, set the "activeTab" attribute to the first visible sheet in the workbook
    if (hideSheet && activeTabIndex == index) {    // BUGFIX 2024-04-30: previously, the active tab was re-set even if the current sheet was
                                                   // being set to "visible" (when already being visible)
        XMLNode item = xmlDocument().document_element().child("sheets").first_child_of_type(pugi::node_element);
        while (not item.empty()) {
            if (isVisible(item))
            {    // BUGFIX 2024-05-01: old check was testing state != "hidden" || != "veryHidden", which was always true
                activeTabAttribute.set_value(indexOfSheet(item.attribute("name").value()) - 1);
                break;
            }
            item = item.next_sibling_of_type(pugi::node_element);
        }
    }
}

/**
 * @details
 * @done In some cases (eg. if a sheet is moved to the position before the selected sheet), multiple sheets are selected when opened
 * in Excel.
 */
void XLWorkbook::setSheetIndex(const std::string& sheetName, unsigned int index)    // 2024-05-01: whitespace support
{
    // ===== Determine the index of the active tab, if any
    const XMLNode  workbookView     = xmlDocument().document_element().child("bookViews").first_child_of_type(pugi::node_element);
    const uint32_t activeSheetIndex = 1 + workbookView.attribute("activeTab").as_uint(0);    // use index in the 1.. range

    // ===== Attempt to locate the sheet with sheetName, and look out for the sheet @index, and the sheet @activeIndex while iterating over
    // the sheets
    XMLNode      sheetToMove      {};    // determine the sheet matching sheetName, if any
    unsigned int sheetToMoveIndex = 0;
    XMLNode      existingSheet    {};    // determine the sheet at index, if any
    std::string  activeSheet_rId  {};    // determine the r:id of the sheet at activeIndex, if any

    unsigned int sheetIndex   = 1;
    XMLNode      curSheet     = sheetsNode(xmlDocument()).first_child_of_type(pugi::node_element);
    int          thingsToFind = (activeSheetIndex > 0) ? 3 : 2;    // if there is no active tab configured, no need to search for its name

    while (not curSheet.empty() && thingsToFind > 0) {    // permit early loop exit when all sheets are located
        if (sheetToMove.empty() && (curSheet.attribute("name").value() == sheetName)) {
            sheetToMoveIndex = sheetIndex;
            sheetToMove      = curSheet;
            --thingsToFind;
        }
        if (existingSheet.empty() && (sheetIndex == index)) {
            existingSheet = curSheet;
            --thingsToFind;
        }
        if (activeSheet_rId.empty() && (sheetIndex == activeSheetIndex)) {    // if no active sheet: activeSheetIndex 0 never matches
            activeSheet_rId = curSheet.attribute("r:id").value();
            --thingsToFind;
        }
        ++sheetIndex;
        curSheet = curSheet.next_sibling_of_type(pugi::node_element);
    }

    // ===== If a sheet with sheetName was not found
    if (sheetToMove.empty()) throw XLInputError(std::string(__func__) + std::string(": no worksheet exists with name ") + sheetName);

    // ==== name was matched

    // ===== If the new index is equal to the current, don't do anything
    if (index == sheetToMoveIndex) return;

    // ===== Modify the node in the XML file
    if (existingSheet.empty())    // new index is beyond last sheet
        sheetsNode(xmlDocument()).append_move(sheetToMove);
    else {    // existingSheet was found
        if (sheetToMoveIndex < index)
            sheetsNode(xmlDocument()).insert_move_after(sheetToMove, existingSheet);
        else    // sheetToMoveIndex > index, because if equal, function never gets here
            sheetsNode(xmlDocument()).insert_move_before(sheetToMove, existingSheet);
    }

    // ===== Updated defined names with worksheet scopes. TBD what this does
    XMLNode definedName = xmlDocument().document_element().child("definedNames").first_child_of_type(pugi::node_element);
    while (not definedName.empty()) {
        // TBD: is the current definedName actually associated with the sheet that was moved?
        definedName.attribute("localSheetId").set_value(sheetToMoveIndex - 1);
        definedName = definedName.next_sibling_of_type(pugi::node_element);
    }

    // ===== Update the activeTab attribute.
    if ((activeSheetIndex < std::min(index, sheetToMoveIndex)) ||
        (activeSheetIndex > std::max(index, sheetToMoveIndex)))    // if active sheet was not within the set of sheets affected by the move
        return;                                                        // nothing to do

    if (activeSheet_rId.length() > 0) setSheetActive(activeSheet_rId);
}

/**
 * @details
 */
unsigned int XLWorkbook::indexOfSheet(const std::string& sheetName) const    // 2024-05-01: whitespace support
{
    // ===== Iterate through sheet nodes. When a match is found, return the index;
    unsigned int index = 1;
    for (XMLNode sheet = sheetsNode(xmlDocument()).first_child_of_type(pugi::node_element);
         not sheet.empty();
         sheet         = sheet.next_sibling_of_type(pugi::node_element))
    {
        if (sheetName == sheet.attribute("name").value()) return index;
        index++;
    }

    // ===== If a match is not found, throw an exception.
    throw XLInputError("Sheet does not exist");
}

/**
 * @details
 */
XLSheetType XLWorkbook::typeOfSheet(const std::string& sheetName) const
{
    if (!sheetExists(sheetName)) throw XLInputError("Sheet with name \"" + sheetName + "\" doesn't exist.");

    if (worksheetExists(sheetName)) return XLSheetType::Worksheet;
    return XLSheetType::Chartsheet;
}

/**
 * @details
 */
XLSheetType XLWorkbook::typeOfSheet(unsigned int index) const    // 2024-05-01: whitespace support
{
    unsigned int thisIndex = 1;
    for (XMLNode sheet = sheetsNode(xmlDocument()).first_child_of_type(pugi::node_element);
         not sheet.empty();
         sheet         = sheet.next_sibling_of_type(pugi::node_element))
    {
        if (thisIndex == index) return typeOfSheet(sheet.attribute("name").as_string());
        ++thisIndex;
    }

    using namespace std::literals::string_literals;
    throw XLInputError(std::string(__func__) + ": index "s + std::to_string(index) + " is out of range"s);
}

/**
 * @details
 */
unsigned int XLWorkbook::sheetCount() const    // 2024-04-30: whitespace support
{
    unsigned int count = 0;
    for (XMLNode node = sheetsNode(xmlDocument()).first_child_of_type(pugi::node_element);
         not node.empty();
         node         = node.next_sibling_of_type(pugi::node_element))
        ++count;
    return count;
}

/**
 * @details
 */
unsigned int XLWorkbook::worksheetCount() const { return static_cast<unsigned int>(worksheetNames().size()); }

/**
 * @details
 */
unsigned int XLWorkbook::chartsheetCount() const { return static_cast<unsigned int>(chartsheetNames().size()); }

/**
 * @details
 */
std::vector<std::string> XLWorkbook::sheetNames() const    // 2024-05-01: whitespace support
{
    std::vector<std::string> results;

    for (XMLNode item = sheetsNode(xmlDocument()).first_child_of_type(pugi::node_element);
         not item.empty();
         item         = item.next_sibling_of_type(pugi::node_element))
        results.emplace_back(item.attribute("name").value());

    return results;
}

/**
 * @details
 */
std::vector<std::string> XLWorkbook::worksheetNames() const    // 2024-05-01: whitespace support
{
    std::vector<std::string> results;

    for (XMLNode item = sheetsNode(xmlDocument()).first_child_of_type(pugi::node_element);
         not item.empty();
         item         = item.next_sibling_of_type(pugi::node_element))
    {
        XLQuery query(XLQueryType::QuerySheetType);
        query.setParam("sheetID", std::string(item.attribute("r:id").value()));
        if (parentDoc().execQuery(query).result<XLContentType>() == XLContentType::Worksheet)
            results.emplace_back(item.attribute("name").value());
    }

    return results;
}

/**
 * @details
 */
std::vector<std::string> XLWorkbook::chartsheetNames() const    // 2024-05-01: whitespace support
{
    std::vector<std::string> results;

    for (XMLNode item = sheetsNode(xmlDocument()).first_child_of_type(pugi::node_element);
         not item.empty();
         item         = item.next_sibling_of_type(pugi::node_element))
    {
        XLQuery query(XLQueryType::QuerySheetType);
        query.setParam("sheetID", std::string(item.attribute("r:id").value()));
        if (parentDoc().execQuery(query).result<XLContentType>() == XLContentType::Chartsheet)
            results.emplace_back(item.attribute("name").value());
    }

    return results;
}

/**
 * @details
 */
bool XLWorkbook::sheetExists(const std::string& sheetName) const { return chartsheetExists(sheetName) || worksheetExists(sheetName); }

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
 * @details The UpdateSheetName member function searches throug the usages of the old name and replaces with the
 * new sheet name.
 * @todo Currently, this function only searches through defined names. Consider using this function to update the
 * actual sheet name as well.
 */
void XLWorkbook::updateSheetReferences(
    const std::string& oldName,
    const std::string& newName)    // 2024-05-01: whitespace support with TBD to verify definedNames logic
{
    //        for (auto& sheet : m_sheets) {
    //            if (sheet.sheetType == XLSheetType::WorkSheet)
    //                Worksheet(sheet.sheetNode.attribute("name").getValue())->UpdateSheetName(oldName, newName);
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

    // ===== Iterate through all defined names // TODO 2024-05-01: verify definedNames logic
    XMLNode definedName = xmlDocument().document_element().child("definedNames").first_child_of_type(pugi::node_element);
    for (; not definedName.empty(); definedName = definedName.next_sibling_of_type(pugi::node_element)) {
        formula = definedName.text().get();

        // ===== Skip if formula contains a '[' and ']' (means that the defined refers to external workbook)
        if (formula.find('[') == std::string::npos && formula.find(']') == std::string::npos) {
            // ===== For all instances of the old sheet name in the formula, replace with the new name.
            while (formula.find(oldNameTemp) != std::string::npos) {    // NOLINT
                formula.replace(formula.find(oldNameTemp), oldNameTemp.length(), newNameTemp);
            }
            definedName.text().set(formula.c_str());
        }
    }
}

/**
 * @details
 */
void XLWorkbook::setFullCalculationOnLoad()
{
    XMLNode calcPr = xmlDocument().document_element().child("calcPr");

    auto getOrCreateAttribute = [&calcPr](const char* attributeName) {
        XMLAttribute attr = calcPr.attribute(attributeName);
        if (attr.empty()) attr = calcPr.append_attribute(attributeName);
        return attr;
    };

    getOrCreateAttribute("forceFullCalc").set_value(true);
    getOrCreateAttribute("fullCalcOnLoad").set_value(true);
}

/**
 * @details
 */
void XLWorkbook::print(std::basic_ostream<char>& ostr) const { xmlDocument().document_element().print (ostr); }

/**
 * @details
 */
bool XLWorkbook::sheetIsActive(const std::string& sheetRID) const    // 2024-04-30: whitespace support
{
    const XMLNode      workbookView       = xmlDocument().document_element().child("bookViews").first_child_of_type(pugi::node_element);
    const XMLAttribute activeTabAttribute = workbookView.attribute("activeTab");
    const int32_t      activeTabIndex     = (not activeTabAttribute.empty() ? activeTabAttribute.as_int() : -1); // 2024-05-29 BUGFIX: activeTabAttribute was being read as_uint
    if (activeTabIndex == -1) return false;    // 2024-05-29 early exit: no need to try and match sheetRID if there *is* no active tab

    int32_t index = 0;    // 2024-06-04 BUGFIX: index should support -1 as 2024-05-29 change below sets it to -1 for preventing a match with activeTabIndex
    XMLNode      item  = sheetsNode(xmlDocument()).first_child_of_type(pugi::node_element);
    while (not item.empty()) {
        if (std::string(item.attribute("r:id").value()) == sheetRID) break;
        ++index;
        item = item.next_sibling_of_type(pugi::node_element);
    }
    if (item.empty()) index = -1;    // 2024-05-29: prevent a match if activeTabIndex invalidly points to a non-existing sheet

    return index == activeTabIndex;
}

/**
 * @details
 * @done: no exception if setSheetActive fails, instead return false
 * @done: fail by returning false if sheetRID is either not found or belongs to a sheet that is not visible
 * @note: this makes some bug fixes from 2024-05-29 obsolete
 * @note: changed behavior: attempting to setSheetActive on a non-existing or non-visible sheet will no longer unselect the active sheet
 */
bool XLWorkbook::setSheetActive(const std::string& sheetRID)    // 2024-04-30: whitespace support
{
    XMLNode            workbookView       = xmlDocument().document_element().child("bookViews").first_child_of_type(pugi::node_element);
    const XMLAttribute activeTabAttribute = workbookView.attribute("activeTab");
    int32_t            activeTabIndex     = -1;    // negative == no active tab identified
    if (not activeTabAttribute.empty()) activeTabIndex = activeTabAttribute.as_int();

    int32_t index = 0;    // index should have the same data type as activeTabIndex for comparisons
    XMLNode      item  = sheetsNode(xmlDocument()).first_child_of_type(pugi::node_element);
    while (not item.empty() && (std::string(item.attribute("r:id").value()) != sheetRID)) {
        ++index;
        item = item.next_sibling_of_type(pugi::node_element);
    }
    // ===== 2024-06-19: Fail without action if sheet is not found or sheet is not visible
    if (item.empty() || !isVisible(item)) return false;

    // NOTE: XLSheet XLWorkbook::sheet(uint16_t index) is using a 1-based index, while the workbookView attribute activeTab is using a
    // 0-based index

    // ===== If an active sheet was found, but sheetRID is not the same sheet: attempt to unselect the old active sheet.
    if ((activeTabIndex != -1) && (index != activeTabIndex)) sheet(activeTabIndex + 1).setSelected(false);    // see NOTE above

    // ===== Set the activeTab property for the workbook.xml sheets node
    if (workbookView.attribute("activeTab").empty()) workbookView.append_attribute("activeTab");
    workbookView.attribute("activeTab").set_value(index);
    // sheet(index + 1).setSelected(true);  // it appears that an active sheet does not have to be selected
    return true;    // success
}

/**
 * @details evaluate a sheet node state attribute where "hidden" or "veryHidden" means not visible
 * @note 2024-05-01 BUGFIX: veryHidden was not checked (in setSheetActive)
 */
bool XLWorkbook::isVisibleState(std::string const& state) const { return (state != "hidden" && state != "veryHidden"); }

/**
 * @details function only returns meaningful information when used with a sheet node
 *          (or nodes with state attribute allowing values visible, hidden, veryHidden)
 */
bool XLWorkbook::isVisible(XMLNode const& sheetNode) const
{
    if (sheetNode.empty()) return false;                      // empty nodes can't be visible
    if (sheetNode.attribute("state").empty()) return true;    // no state attribute means no hidden-ness can be flagged
    // ===== Attribute exists and value must be checked
    return isVisibleState(sheetNode.attribute("state").value());
}

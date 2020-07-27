//
// Created by Troldal on 02/09/16.
//

#include "XLWorkbook.hpp"

#include "XLChartsheet.hpp"
#include "XLDocument.hpp"
#include "XLSheet.hpp"
#include "XLWorksheet.hpp"

#include <algorithm>
#include <iterator>
#include <utility>
#include <vector>

using namespace OpenXLSX;

namespace
{
    /**
     * @brief
     * @param sheetsNode
     * @param sheetRID
     * @return
     */
    inline XMLNode getSheetNodeByRID(XMLNode sheetsNode, const std::string& sheetRID)
    {
        return sheetsNode.find_child_by_attribute("r:id", sheetRID.c_str());
    }

    /**
     * @brief
     * @param sheetsNode
     * @param sheetID
     * @return
     */
    inline XMLNode getSheetNodeByID(XMLNode sheetsNode, uint16_t sheetID)
    {
        return sheetsNode.find_child_by_attribute("sheetId", std::to_string(sheetID).c_str());
    }

    /**
     * @brief
     * @param sheetsNode
     * @param sheetName
     * @return
     */
    inline XMLNode getSheetNodeByName(XMLNode sheetsNode, const std::string& sheetName)
    {
        return sheetsNode.find_child_by_attribute("name", sheetName.c_str());
    }

    /**
     * @brief
     * @param sheetsNode
     * @param sheetRID
     * @return
     */
    inline uint16_t getSheetIndexByRID(XMLNode sheetsNode, const std::string& sheetRID)
    {
        auto sheetNode = getSheetNodeByRID(sheetsNode, sheetRID);
        if (!sheetsNode) throw XLException("Sheet does not exist");

        for (auto iter = sheetsNode.children().begin(); iter != sheetsNode.children().end(); ++iter)
            if (*iter == sheetNode) return std::distance(sheetNode.children().begin(), iter);
        throw XLException("Sheet with rID \"" + sheetRID + "\" does not exist.");
    }

    /**
     * @brief
     * @param sheetsNode
     * @param sheetID
     * @return
     */
    inline uint16_t getSheetIndexByID(XMLNode sheetsNode, uint16_t sheetID)
    {
        auto sheetNode = getSheetNodeByID(sheetsNode, sheetID);
        if (!sheetsNode) throw XLException("Sheet does not exist");

        for (auto iter = sheetsNode.children().begin(); iter != sheetsNode.children().end(); ++iter)
            if (*iter == sheetNode) return std::distance(sheetNode.children().begin(), iter);
        throw XLException("Sheet with ID \"" + std::to_string(sheetID) + "\" does not exist.");
    }

    /**
     * @brief
     * @param sheetsNode
     * @param sheetName
     * @return
     */
    inline uint16_t getSheetIndexByName(XMLNode sheetsNode, const std::string& sheetName)
    {
        auto sheetNode = getSheetNodeByName(sheetsNode, sheetName);
        if (!sheetsNode) throw XLException("Sheet does not exist");

        for (auto iter = sheetsNode.children().begin(); iter != sheetsNode.children().end(); ++iter)
            if (*iter == sheetNode) return std::distance(sheetNode.children().begin(), iter);
        throw XLException("Sheet with name \"" + sheetName + "\" does not exist.");
    }

    /**
     * @brief
     * @param sheetsNode
     * @param sheetRID
     * @return
     */
    inline std::string getSheetNameByRID(XMLNode sheetsNode, const std::string& sheetRID)
    {
        return getSheetNodeByRID(sheetsNode, sheetRID).attribute("name").value();
    }

    /**
     * @brief
     * @param sheetsNode
     * @param sheetName
     * @return
     */
    inline std::string getSheetRIDByName(XMLNode sheetsNode, const std::string& sheetName)
    {
        return getSheetNodeByName(sheetsNode, sheetName).attribute("r:id").value();
    }

    /**
     * @brief
     * @param sheetsNode
     * @return
     */
    inline uint16_t calculateNewSheetNumber(XMLNode sheetsNode)
    {
        return std::max_element(
                   sheetsNode.children().begin(),
                   sheetsNode.children().end(),
                   [](const XMLNode& a, const XMLNode& b) { return a.attribute("sheetId").as_uint() < b.attribute("sheetId").as_uint(); })
                   ->attribute("sheetId")
                   .as_uint() +
               1;
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
    return sheet(std::vector<XMLNode>(getSheetsNode().begin(), getSheetsNode().end())[index - 1].attribute("name").as_string());
}

/**
 * @details
 */
XLWorksheet XLWorkbook::worksheet(const std::string& sheetName)
{
    return sheet(sheetName).Get<XLWorksheet>();
}

/**
 * @details
 */
XLChartsheet XLWorkbook::chartsheet(const std::string& sheetName)
{
    return sheet(sheetName).Get<XLChartsheet>();
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
    auto sheetID        = getSheetsNode().find_child_by_attribute("name", sheetName.c_str()).attribute("r:id").value();    // NOLINT
    auto sheetType      = parentDoc().executeQuery(XLQuerySheetType(getRID())).sheetType();
    auto worksheetCount = std::count_if(getSheetsNode().children().begin(), getSheetsNode().children().end(), [&](const XMLNode& item) {
        return parentDoc().executeQuery(XLQuerySheetType(item.attribute("r:id").value())).sheetType() == XLContentType::Worksheet;
    });

    // ===== If this is the last worksheet in the workbook, throw an exception.
    if (worksheetCount == 1 && sheetType == XLContentType::Worksheet)
        throw XLException("Invalid operation. There must be at least one worksheet in the workbook.");

    // ===== Delete the sheet data as well as the sheet node from Workbook.xml
    parentDoc().executeCommand(XLCommandDeleteSheet(sheetID, sheetName));
    getSheetsNode().remove_child(getSheetsNode().find_child_by_attribute("name", sheetName.c_str()));

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
    auto internalID = calculateNewSheetNumber(xmlDocument().document_element().child("sheets"));

    // ===== Create xml file for new worksheet and add metadata to the workbook file.
    parentDoc().executeCommand(XLCommandAddWorksheet(sheetName, "/xl/worksheets/sheet" + std::to_string(internalID) + ".xml"));
    prepareSheetMetadata(sheetName, internalID);
}

/**
 * @details
 * @todo If the original sheet's tabSelected attribute is set, ensure it is un-set in the clone.
 */
void XLWorkbook::cloneSheet(const std::string& extName, const std::string& newName)
{
    // ===== If a sheet with the given name already exists, throw an exception.
    if (xmlDocument().document_element().child("sheets").find_child_by_attribute("name", newName.c_str()))
        throw XLException("Sheet named \"" + newName + "\" already exists.");

    // ===== Create new internal (workbook) ID for the sheet and retrieve the sheet ID for the sheet to clone.
    auto        internalID = calculateNewSheetNumber(xmlDocument().document_element().child("sheets"));
    std::string sheetID    = getSheetsNode().find_child_by_attribute("name", extName.c_str()).attribute("r:id").value();

    // ===== Create xml file for new worksheet and add metadata to the workbook file.
    parentDoc().executeCommand(XLCommandCloneSheet(sheetID, newName, "/xl/worksheets/sheet" + std::to_string(internalID) + ".xml"));
    prepareSheetMetadata(newName, internalID);
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
    auto node = getSheetsNode().append_child("sheet");

    // ===== append the required attributes to the newly created sheet node.
    std::string sheetPath            = "/xl/worksheets/sheet" + std::to_string(internalID) + ".xml";
    node.append_attribute("name")    = sheetName.c_str();
    node.append_attribute("sheetId") = std::to_string(internalID).c_str();
    node.append_attribute("r:id")    = parentDoc().executeQuery(XLQuerySheetRelsID(sheetPath)).sheetID().c_str();
}

/**
 * @details
 * @todo Consider putting this in an anonymous namespace.
 */
XMLNode XLWorkbook::getSheetsNode() const
{
    return xmlDocument().first_child().child("sheets");
}

/**
 * @details
 */
void XLWorkbook::setSheetName(const std::string& sheetRID, const std::string& newName)
{
    auto sheetName = xmlDocument().document_element().child("sheets").find_child_by_attribute("r:id", sheetRID.c_str()).attribute("name");

    updateSheetName(sheetName.value(), newName);
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
    if (index < 1 || index > std::distance(xmlDocument().document_element().child("sheets").children().begin(),
                                           xmlDocument().document_element().child("sheets").children().end()))
        throw XLException("Invalid sheet index");

    // ===== If the new index is equal to the current, don't do anything
    if (index == std::distance(xmlDocument().document_element().child("sheets").children().begin(),
                               std::find_if(xmlDocument().document_element().child("sheets").children().begin(),
                                            xmlDocument().document_element().child("sheets").children().end(),
                                            [&](const XMLNode& item) { return sheetName == item.attribute("name").value(); })))
        return;

    // ===== Modify the node in the XML file
    if (index <= 1)
        getSheetsNode().prepend_move(getSheetsNode().find_child_by_attribute("name", sheetName.c_str()));
    else if (index >= sheetCount())
        getSheetsNode().append_move(getSheetsNode().find_child_by_attribute("name", sheetName.c_str()));
    else {
        auto currentSheet = getSheetsNode().first_child();
        for (auto i = 1; i < index; ++i) currentSheet = currentSheet.next_sibling();
        getSheetsNode().insert_move_before(getSheetsNode().find_child_by_attribute("name", sheetName.c_str()), currentSheet);
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
    for (auto& sheet : getSheetsNode().children()) {
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
    std::string name = std::vector<XMLNode>(getSheetsNode().begin(), getSheetsNode().end())[index - 1].attribute("name").as_string();
    return typeOfSheet(name);
}

/**
 * @details
 */
unsigned int XLWorkbook::sheetCount() const
{
    return std::distance(getSheetsNode().children().begin(), getSheetsNode().children().end());
}

/**
 * @details
 */
unsigned int XLWorkbook::worksheetCount() const
{
    return worksheetNames().size();
}

/**
 * @details
 */
unsigned int XLWorkbook::chartsheetCount() const
{
    return chartsheetNames().size();
}

/**
 * @details
 */
std::vector<std::string> XLWorkbook::sheetNames() const
{
    std::vector<std::string> results;

    for (const auto& item : getSheetsNode().children()) results.emplace_back(item.attribute("name").value());

    return results;
}

/**
 * @details
 */
std::vector<std::string> XLWorkbook::worksheetNames() const
{
    std::vector<std::string> results;

    for (const auto& item : getSheetsNode().children()) {
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

    for (const auto& item : getSheetsNode().children()) {
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
void XLWorkbook::updateSheetName(const std::string& oldName, const std::string& newName)
{
    //    for (auto& sheet : m_sheets) {
    //        if (sheet.sheetType == XLSheetType::WorkSheet)
    //            Worksheet(sheet.sheetNode.attribute("name").value())->UpdateSheetName(oldName, newName);
    //    }
    //
    //    // ===== Set up temporary variables
    //    std::string oldNameTemp = oldName;
    //    std::string newNameTemp = newName;
    //    std::string formula;
    //
    //    // ===== If the sheet name contains spaces, it should be enclosed in single quotes (')
    //    if (oldName.find(' ') != std::string::npos) oldNameTemp = "\'" + oldName + "\'";
    //    if (newName.find(' ') != std::string::npos) newNameTemp = "\'" + newName + "\'";
    //
    //    // ===== Ensure only sheet names are replaced (references to sheets always ends with a '!')
    //    oldNameTemp += '!';
    //    newNameTemp += '!';
    //
    //    // ===== Iterate through all defined names
    //    for (auto& definedName : XmlDocument()->document_element().child("definedNames").children()) {
    //        formula = definedName.text().get();
    //
    //        // ===== Skip if formula contains a '[' and ']' (means that the defined refers to external workbook)
    //        if (formula.find('[') == string::npos && formula.find(']') == string::npos) {
    //            // ===== For all instances of the old sheet name in the formula, replace with the new name.
    //            while (formula.find(oldNameTemp) != string::npos) {
    //                formula.replace(formula.find(oldNameTemp), oldNameTemp.length(), newNameTemp);
    //            }
    //            definedName.text().set(formula.c_str());
    //        }
    //    }
}

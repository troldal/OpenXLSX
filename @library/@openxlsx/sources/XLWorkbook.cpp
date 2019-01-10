//
// Created by Troldal on 02/09/16.
//

#include <cstring>
#include <pugixml.hpp>

#include "XLWorkbook.h"
#include "XLWorksheet.h"
#include "XLDocument.h"

using namespace std;
using namespace OpenXLSX::Impl;


/**
 * @details The constructor initializes the member variables and calls the loadXMLData from the
 * XLAbstractXMLFile base class.
 */
XLWorkbook::XLWorkbook(XLDocument &parent,
                       const std::string &filePath)

    : XLAbstractXMLFile(parent, filePath),
      XLSpreadsheetElement(parent),
      m_sheetsNode(std::make_unique<XMLNode>()),
      m_sheetNodes(),
      m_sheets(),
      m_sheetPaths(),
      m_sheetId(0),
      m_sheetCount(0),
      m_worksheetCount(0),
      m_chartsheetCount(0),
      m_definedNames(std::make_unique<XMLNode>()),
      m_relationships(nullptr),
      m_sharedStrings(nullptr),
      m_styles(nullptr)
{
    SetModified();
    ParseXMLData();
}

/**
 * @details
 */
bool XLWorkbook::ParseXMLData()
{

    // Set up the Workbook Relationships.
    m_relationships.reset(new XLRelationships(*ParentDocument(), "xl/_rels/workbook.xml.rels"));

    // Find the "sheets" section in the Workbook.xml file
    for (auto &node : XmlDocument()->first_child().children()) {
        if (string(node.name()) == "sheets") *m_sheetsNode = node;
        if (string(node.name()) == "definedNames") *m_definedNames = node;
    }

    for (auto &sheetNode : m_sheetsNode->children())
        m_sheetNodes[sheetNode.attribute("name").value()] = sheetNode;

    for (auto const &item : *m_relationships->Relationships()) {
        string path = item.second->Target();

        switch (item.second->Type()) {

            // Set up SharedStrings file
            case XLRelationshipType::SharedStrings :
                CreateSharedStrings(*item.second);
                break;

                // Set up Styles file
            case XLRelationshipType::Styles :
                CreateStyles(*item.second);
                break;

                // Set up Worksheet files
            case XLRelationshipType::Worksheet :
                CreateWorksheet(*item.second);
                break;

                // Set up Chartsheet files
            case XLRelationshipType::ChartSheet :
                CreateChartsheet(*item.second);
                break;

            default:
                break;
        }
    }

    return true;
}

/**
 * @details
 * @todo Currently, an XLWorksheet object cannot be created from the const method. This should be fixed.
 */
XLWorksheet * XLWorkbook::Worksheet(const std::string &sheetName)
{
    if (m_sheets.find(sheetName) == m_sheets.end()) throw std::range_error("Sheet \"" + sheetName + "\" does not exist");

    if (m_sheets.at(sheetName) == nullptr) {
        m_sheets.at(sheetName) = make_unique<XLWorksheet>(*this, sheetName, m_sheetPaths.at(sheetName));
        m_childXmlDocuments.at(m_sheetPaths.at(sheetName)) = m_sheets.at(sheetName).get();
    }

    XLSheet *sheet = m_sheets.at(sheetName).get();
    if (sheet->Type() == XLSheetType::WorkSheet)
        return dynamic_cast<XLWorksheet *>(sheet);
    else
        throw std::range_error("Sheet does not exist");
}

/**
 * @details
 */
const XLWorksheet * XLWorkbook::Worksheet(const std::string &sheetName) const
{
    if (m_sheets.find(sheetName) == m_sheets.end()) throw std::range_error("Sheet does not exist");

    XLSheet *sheet = m_sheets.at(sheetName).get();
    if (sheet->Type() == XLSheetType::WorkSheet)
        return dynamic_cast<XLWorksheet *>(sheet);
    else
        throw std::range_error("Sheet does not exist");
}

/**
 * @details
 */
XLSheet *XLWorkbook::Sheet(unsigned int index)
{
    return nullptr;
}

/**
 * @details
 */
XLSheet *XLWorkbook::Sheet(const std::string &sheetName)
{
    return nullptr;
}

/**
 * @details
 */
XLChartsheet *XLWorkbook::Chartsheet(const std::string &sheetName)
{
    return nullptr;
}

/**
 * @details
 */
const XLChartsheet *XLWorkbook::Chartsheet(const std::string &sheetName) const
{
    return nullptr;
}

/**
 * @details
 */
bool XLWorkbook::HasSharedStrings() const
{
    if (m_sharedStrings) return true;
    return false;
}

/**
 * @details
 */
XLSharedStrings *XLWorkbook::SharedStrings() const
{
    return m_sharedStrings.get();
}

/**
 * @details
 */
void XLWorkbook::DeleteNamedRanges()
{
    for (auto &child : m_definedNames->children()) child.parent().remove_child(child);
    m_isModified = true;
}

/**
 * @details
 */
void XLWorkbook::DeleteSheet(const std::string &sheetName)
{

    // Erase file Path from internal datastructure.
    m_childXmlDocuments.erase(Worksheet(sheetName)->FilePath());

    // Clear Worksheet and set to safe State
    Worksheet(sheetName)->Delete();

    // Delete the pointer to the object
    m_sheets.erase(sheetName);

    // Delete the node from Workbook.xml
    m_sheetNodes.at(sheetName).parent().remove_child(m_sheetNodes.at(sheetName));
    m_sheetNodes.erase(sheetName);

    m_sheetCount--;
    m_worksheetCount--;
}

/**
 * @details
 */
void XLWorkbook::AddWorksheet(const std::string &sheetName,
                              unsigned int index)
{
    CreateWorksheet(*InitiateWorksheet(sheetName, index), XLWorksheet::NewSheetXmlData());
}

/**
 * @details
 * @todo If the original sheet's tebSelected attribute is set, ensure it is un-set in the clone.
 */
void XLWorkbook::CloneWorksheet(const std::string &extName,
                                const std::string &newName,
                                unsigned int index)
{
    CreateWorksheet(*InitiateWorksheet(newName, index), Worksheet(extName)->GetXmlData());
}

/**
 * @details
 */
XLRelationshipItem *XLWorkbook::InitiateWorksheet(const std::string &sheetName, unsigned int index)
{
    std::string worksheetPath = "/xl/worksheets/sheet" + to_string(m_sheetId + 1) + ".xml";

    // Add content item to document
    ParentDocument()->AddContentItem(worksheetPath, XLContentType::Worksheet);

    // Add relationship item
    XLRelationshipItem &item = *m_relationships->AddRelationship(XLRelationshipType::Worksheet,
                                                                "worksheets/sheet" + to_string(m_sheetId + 1) + ".xml");

    // insert Sheet node at the given Index
    if (index == 0) {
        auto node = m_sheetsNode->append_child("sheet");
        node.append_attribute("name") = sheetName.c_str();
        node.append_attribute("sheetId") = to_string(m_sheetId + 1).c_str();
        node.append_attribute("r:id") = item.Id().c_str();
        m_sheetNodes[sheetName] = node;
    }
    else {
        auto curNode = m_sheetsNode->first_child();
        if (!curNode) {
            auto node = m_sheetsNode->append_child("sheet");
            node.append_attribute("name") = sheetName.c_str();
            node.append_attribute("sheetId") = to_string(m_sheetId + 1).c_str();
            node.append_attribute("r:id") = item.Id().c_str();
            m_sheetNodes[sheetName] = node;
        }
        else {
            unsigned int idx = 1;
            while (idx < index) {
                curNode = curNode.next_sibling();
                ++idx;
                if (!curNode) {
                    auto node = m_sheetsNode->append_child("sheet");
                    node.append_attribute("name") = sheetName.c_str();
                    node.append_attribute("sheetId") = to_string(m_sheetId + 1).c_str();
                    node.append_attribute("r:id") = item.Id().c_str();
                    m_sheetNodes[sheetName] = node;
                }
            }
            auto node = m_sheetsNode->insert_child_before("sheet", curNode);
            node.append_attribute("name") = sheetName.c_str();
            node.append_attribute("sheetId") = to_string(m_sheetId + 1).c_str();
            node.append_attribute("r:id") = item.Id().c_str();
            m_sheetNodes[sheetName] = node;
        }
    }

    // Add entry to the App Properties
    if (index == 0) ParentDocument()->AppProperties()->InsertSheetName(sheetName, WorksheetCount() + 1);
    else ParentDocument()->AppProperties()->InsertSheetName(sheetName, index);

    ParentDocument()->AppProperties()->SetHeadingPair("Worksheets", WorksheetCount() + 1);

    return &item;
}

/**
 * @details
 * @todo To be implemented.
 */
void XLWorkbook::AddChartsheet(const std::string &sheetName,
                               unsigned int index)
{

}

/**
 * @details
 * @todo To be implemented.
 */
void XLWorkbook::MoveSheet(unsigned int index)
{

}

/**
 * @details
 */
unsigned int XLWorkbook::IndexOfSheet(const std::string &sheetName)
{

    auto node = m_sheetsNode->first_child();
    if (node.type() == pugi::node_null) throw runtime_error("Sheet does not exist.");

    unsigned int index = 1;
    while (node) {
        if (strcmp(node.attribute("name").value(), sheetName.c_str()) == 0) break;
        node = node.next_sibling();
        if (!node) throw runtime_error("Sheet does not exist.");
        ++index;
    }

    return index;
}

/**
 * @details
 */
void XLWorkbook::SetIndexOfSheet(const std::string &sheetName,
                                 unsigned int index)
{

}

/**
 * @details
 */
unsigned int XLWorkbook::SheetCount() const
{
    return m_sheetCount;
}

/**
 * @details
 */
unsigned int XLWorkbook::WorksheetCount() const
{
    return m_worksheetCount;
}

/**
 * @details
 */
unsigned int XLWorkbook::ChartsheetCount() const
{
    return m_chartsheetCount;
}

/**
 * @details
 */
XLStyles *XLWorkbook::Styles()
{
    return m_styles.get();
}

/**
 * @details
 */
bool XLWorkbook::SheetExists(const std::string &sheetName) const
{

    bool result = false;
    for (const auto &iter : m_sheets) {
        if (iter.second->Name() == sheetName) {
            result = true;
            break;
        }
    }

    return result;
}

/**
 * @details
 */
bool XLWorkbook::WorksheetExists(const std::string &sheetName) const
{

    bool result = false;
    for (const auto &iter : m_sheets) {
        if (iter.second->Name() == sheetName && iter.second->Type() == XLSheetType::WorkSheet) {
            result = true;
            break;
        }
    }

    return result;
}

/**
 * @details
 */
bool XLWorkbook::ChartsheetExists(const std::string &sheetName) const
{

    bool result = false;
    for (const auto &iter : m_sheets) {
        if (iter.second->Name() == sheetName && iter.second->Type() == XLSheetType::ChartSheet) {
            result = true;
            break;
        }
    }

    return result;
}

/**
 * @details
 */
XLRelationships *XLWorkbook::Relationships()
{
    return m_relationships.get();
}

/**
 * @details
 */
const XLRelationships *XLWorkbook::Relationships() const
{
    return m_relationships.get();
}

/**
 * @details
 */
XMLNode XLWorkbook::SheetNode(const string &sheetName)
{
    return m_sheetNodes.at(sheetName);
}

/**
 * @details
 */
void XLWorkbook::CreateSharedStrings(const XLRelationshipItem &item)
{
    m_sharedStrings.reset(new XLSharedStrings(*ParentDocument(), "xl/" + item.Target()));
    m_childXmlDocuments[m_sharedStrings->FilePath()] = m_sharedStrings.get();
}

/**
 * @details
 */
void XLWorkbook::CreateStyles(const XLRelationshipItem &item)
{
    m_styles.reset(new XLStyles(*ParentDocument(), "xl/" + item.Target()));
    m_childXmlDocuments[m_styles->FilePath()] = m_styles.get();
}

/**
 * @details
 */
void XLWorkbook::CreateWorksheet(const XLRelationshipItem &item, const std::string& xmlData)
{

    // Find the appropriate sheet node in the Workbook .xml file; get the name and id of the worksheet.
    string name;
    for (auto &node : m_sheetsNode->children()) {
        if (string(node.attribute("r:id").value()) == item.Id()) {
            name = node.attribute("name").value();
            break;
        }
    }

    // If xmlData is empty, set the m_sheets and m_childXmlDocuments elements to nullptr. The worksheet will then be
    // lazy-instantiated when the worksheet is requested using the 'Worksheet" function.
    // If xmlData is provided (i.e. a new sheet is created or an existing is cloned), create the new worksheets accordingly.
    m_sheetPaths[name] = "xl/" + item.Target();
    if (xmlData.empty()) {
        m_sheets[name] = nullptr;
        m_childXmlDocuments["xl/" + item.Target()] = nullptr;
    } else {
        m_sheets[name] = make_unique<XLWorksheet>(*this, name, "xl/" + item.Target(), xmlData);
        m_childXmlDocuments["xl/" + item.Target()] = m_sheets.at(name).get();
    }

    // Update internal counters.
    m_sheetCount++;
    m_worksheetCount++;
    m_sheetId++;
}

/**
 * @details
 */
void XLWorkbook::CreateChartsheet(const XLRelationshipItem &item)
{

    string name;
    auto node = m_sheetsNode->first_child();
    while (node) {
        if (strcmp(node.attribute("r:id").value(), item.Id().c_str()) == 0) {
            name = node.attribute("name").value();
            break;
        }
        node = node.next_sibling();
    }

    //TODO: Create Chartsheet object here.


    m_sheetCount++;
    m_chartsheetCount++;
    m_sheetId++;
}

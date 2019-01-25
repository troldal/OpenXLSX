//
// Created by Troldal on 02/09/16.
//

#include <cstring>
#include <pugixml.hpp>
#include <XLWorkbook_Impl.h>

#include "XLWorksheet_Impl.h"

using namespace std;
using namespace OpenXLSX;

/**
 * @details The constructor initializes the member variables and calls the loadXMLData from the
 * XLAbstractXMLFile base class.
 */
Impl::XLWorkbook::XLWorkbook(XLDocument& parent,
                             const std::string& filePath)

        : XLAbstractXMLFile(parent,
                            filePath),
          XLSpreadsheetElement(parent),
          m_sheetsNode(),
          m_sheets(),
          m_sheetPaths(),
          m_sheetId(0),
          m_sheetCount(0),
          m_worksheetCount(0),
          m_chartsheetCount(0),
          m_definedNames(),
          m_relationships(nullptr),
          m_sharedStrings(nullptr),
          m_styles(nullptr) {

    ParseXMLData();
}

/**
 * @details
 */
bool Impl::XLWorkbook::ParseXMLData() {

    // Set up the Workbook Relationships.
    m_relationships.reset(new XLRelationships(*ParentDocument(), "xl/_rels/workbook.xml.rels"));

    // Find the "sheets" section in the Workbook.xml file
    for (auto& node : XmlDocument()->first_child().children()) {
        if (string(node.name()) == "sheets")
            m_sheetsNode   = node;
        if (string(node.name()) == "definedNames")
            m_definedNames = node;
    }

    for (auto const& item : *m_relationships->Relationships()) {
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
Impl::XLWorksheet* Impl::XLWorkbook::Worksheet(const std::string& sheetName) {

    if (m_sheets.find(sheetName) == m_sheets.end())
        throw std::range_error("Sheet \"" + sheetName + "\" does not exist");

    if (m_sheets.at(sheetName) == nullptr) {
        m_sheets.at(sheetName)                             = make_unique<XLWorksheet>(*this,
                                                                                      sheetName,
                                                                                      m_sheetPaths.at(sheetName));
        m_childXmlDocuments.at(m_sheetPaths.at(sheetName)) = m_sheets.at(sheetName).get();
    }

    XLSheet* sheet = m_sheets.at(sheetName).get();
    if (sheet->Type() == XLSheetType::WorkSheet)
        return dynamic_cast<XLWorksheet*>(sheet);
    else
        throw std::range_error("Sheet does not exist");
}

/**
 * @details
 */
const Impl::XLWorksheet* Impl::XLWorkbook::Worksheet(const std::string& sheetName) const {

    if (m_sheets.find(sheetName) == m_sheets.end())
        throw std::range_error("Sheet does not exist");

    XLSheet* sheet = m_sheets.at(sheetName).get();
    if (sheet->Type() == XLSheetType::WorkSheet)
        return dynamic_cast<XLWorksheet*>(sheet);
    else
        throw std::range_error("Sheet does not exist");
}

/**
 * @details
 * @todo Fix returning of Chartsheet objects.
 */
Impl::XLSheet* Impl::XLWorkbook::Sheet(unsigned int index) {

    auto node = m_sheetsNode.first_child();
    if (index > 1) {
        for (int i = 1; i < index; ++i) {
            node = node.next_sibling();
        }
    }

    if (m_worksheets.find(node.attribute("name").as_string()) != m_worksheets.end())
        return Worksheet(node.attribute("name").as_string());
    //else
    //return Chartsheet(node.attribute("name").as_string());
}

/**
 * @details
 */
Impl::XLSheet* Impl::XLWorkbook::Sheet(const std::string& sheetName) {

    if (m_worksheets.find(sheetName) != m_worksheets.end())
        return Worksheet(sheetName);
}

/**
 * @details
 */
Impl::XLChartsheet* Impl::XLWorkbook::Chartsheet(const std::string& sheetName) {

    return nullptr;
}

/**
 * @details
 */
const Impl::XLChartsheet* Impl::XLWorkbook::Chartsheet(const std::string& sheetName) const {

    return nullptr;
}

/**
 * @details
 */
bool Impl::XLWorkbook::HasSharedStrings() const {

    if (m_sharedStrings)
        return true;
    return false;
}

/**
 * @details
 */
Impl::XLSharedStrings* Impl::XLWorkbook::SharedStrings() const {

    return m_sharedStrings.get();
}

/**
 * @details
 */
void Impl::XLWorkbook::DeleteNamedRanges() {

    for (auto& child : m_definedNames.children())
        child.parent().remove_child(child);
}

/**
 * @details
 * @todo Throw exception if there is only one worksheet in the workbook.
 */
void Impl::XLWorkbook::DeleteSheet(const std::string& sheetName) {

    // Erase file Path from internal datastructure.
    m_childXmlDocuments.erase(Worksheet(sheetName)->FilePath());

    // Clear Worksheet and set to safe State
    Worksheet(sheetName)->Delete();

    // Delete the pointer to the object
    m_sheets.erase(sheetName);

    // Delete the node from Workbook.xml
    for (auto& sheet : m_sheetsNode.children()) {
        if (sheet.child_value() == sheetName)
            m_sheetsNode.remove_child(sheet);
    }

    m_sheetCount--;
    m_worksheetCount--;
}

/**
 * @details
 */
void Impl::XLWorkbook::AddWorksheet(const std::string& sheetName,
                                    unsigned int index) {

    CreateWorksheet(*InitiateWorksheet(sheetName,
                                       index),
                    XLWorksheet::NewSheetXmlData());
}

/**
 * @details
 * @todo If the original sheet's tebSelected attribute is set, ensure it is un-set in the clone.
 */
void Impl::XLWorkbook::CloneWorksheet(const std::string& extName,
                                      const std::string& newName,
                                      unsigned int index) {

    CreateWorksheet(*InitiateWorksheet(newName,
                                       index),
                    Worksheet(extName)->GetXmlData());
}

/**
 * @details
 */
Impl::XLRelationshipItem* Impl::XLWorkbook::InitiateWorksheet(const std::string& sheetName,
                                                              unsigned int index) {

    std::string worksheetPath = "/xl/worksheets/sheet" + to_string(m_sheetId + 1) + ".xml";

    // Add content item to document
    ParentDocument()->AddContentItem(worksheetPath,
                                     XLContentType::Worksheet);

    // Add relationship item
    XLRelationshipItem& item = *m_relationships->AddRelationship(XLRelationshipType::Worksheet,
                                                                 "worksheets/sheet" + to_string(m_sheetId + 1)
                                                                 + ".xml");

    // insert Sheet node at the given Index
    if (index == 0) {
        auto node = m_sheetsNode.append_child("sheet");
        node.append_attribute("name")    = sheetName.c_str();
        node.append_attribute("sheetId") = to_string(m_sheetId + 1).c_str();
        node.append_attribute("r:id")    = item.Id().c_str();
    }
    else {
        auto curNode = m_sheetsNode.first_child();
        if (!curNode) {
            auto node = m_sheetsNode.append_child("sheet");
            node.append_attribute("name")    = sheetName.c_str();
            node.append_attribute("sheetId") = to_string(m_sheetId + 1).c_str();
            node.append_attribute("r:id")    = item.Id().c_str();
        }
        else {
            unsigned int idx  = 1;
            while (idx < index) {
                curNode = curNode.next_sibling();
                ++idx;
                if (!curNode) {
                    auto node = m_sheetsNode.append_child("sheet");
                    node.append_attribute("name")    = sheetName.c_str();
                    node.append_attribute("sheetId") = to_string(m_sheetId + 1).c_str();
                    node.append_attribute("r:id")    = item.Id().c_str();
                }
            }
            auto         node = m_sheetsNode.insert_child_before("sheet",
                                                                  curNode);
            node.append_attribute("name")    = sheetName.c_str();
            node.append_attribute("sheetId") = to_string(m_sheetId + 1).c_str();
            node.append_attribute("r:id")    = item.Id().c_str();
        }
    }

    // Add entry to the App Properties
    if (index == 0)
        ParentDocument()->AppProperties()->InsertSheetName(sheetName,
                                                           WorksheetCount() + 1);
    else
        ParentDocument()->AppProperties()->InsertSheetName(sheetName,
                                                           index);

    ParentDocument()->AppProperties()->SetHeadingPair("Worksheets",
                                                      WorksheetCount() + 1);

    return &item;
}

void Impl::XLWorkbook::UpdateSheetNames() {

    for (auto& sheet : m_sheets) {

        if (sheet.second != nullptr && sheet.first != sheet.second->Name()) {

            auto oldName = sheet.first;
            auto newName = sheet.second->Name();

            if (sheet.second->Type() == XLSheetType::WorkSheet) {
                m_worksheets[newName] = m_worksheets.at(oldName);
                m_worksheets.erase(oldName);
            }
            else {
                m_chartsheets[newName] = m_chartsheets.at(oldName);
                m_chartsheets.erase(oldName);
            }

            m_sheets[newName] = std::move(m_sheets.at(oldName));
            m_sheets.erase(oldName);

            break;

        }
    }
}

/**
 * @details
 * @todo To be implemented.
 */
void Impl::XLWorkbook::AddChartsheet(const std::string& sheetName,
                                     unsigned int index) {

}

/**
 * @details
 * @todo To be implemented.
 */
void Impl::XLWorkbook::MoveSheet(const std::string& sheetName, unsigned int index) {

}

/**
 * @details
 */
unsigned int Impl::XLWorkbook::IndexOfSheet(const std::string& sheetName) {

    auto node = m_sheetsNode.first_child();
    if (node.type() == pugi::node_null)
        throw runtime_error("Sheet does not exist.");

    unsigned int index = 1;
    while (node) {
        if (strcmp(node.attribute("name").value(),
                   sheetName.c_str()) == 0)
            break;
        node = node.next_sibling();
        if (!node)
            throw runtime_error("Sheet does not exist.");
        ++index;
    }

    return index;
}

/**
 * @details
 */
unsigned int Impl::XLWorkbook::SheetCount() const {

    return m_sheetCount;
}

/**
 * @details
 */
unsigned int Impl::XLWorkbook::WorksheetCount() const {

    return m_worksheetCount;
}

/**
 * @details
 */
unsigned int Impl::XLWorkbook::ChartsheetCount() const {

    return m_chartsheetCount;
}

/**
 * @details
 */
Impl::XLStyles* Impl::XLWorkbook::Styles() {

    return m_styles.get();
}

/**
 * @details
 */
bool Impl::XLWorkbook::SheetExists(const std::string& sheetName) const {

    if (m_sheets.find(sheetName) == m_sheets.end())
        return false;
    else
        return true;
}

/**
 * @details
 */
bool Impl::XLWorkbook::WorksheetExists(const std::string& sheetName) const {

    if (m_worksheets.find(sheetName) == m_worksheets.end())
        return false;
    else
        return true;

/*    if (m_sheets.find(sheetName) == m_sheets.end())
        return false;
    else {

        bool result = false;
        for (const auto& iter : m_sheets) {
            if (iter.second->Name() == sheetName && iter.second->Type() == XLSheetType::WorkSheet) {
                result = true;
                break;
            }
        }

        return result;
    }*/
}

/**
 * @details
 */
bool Impl::XLWorkbook::ChartsheetExists(const std::string& sheetName) const {

    if (m_chartsheets.find(sheetName) == m_chartsheets.end())
        return false;
    else
        return true;

/*    bool result = false;
    for (const auto& iter : m_sheets) {
        if (iter.second->Name() == sheetName && iter.second->Type() == XLSheetType::ChartSheet) {
            result = true;
            break;
        }
    }

    return result;*/
}

/**
 * @details
 */
Impl::XLRelationships* Impl::XLWorkbook::Relationships() {

    return m_relationships.get();
}

/**
 * @details
 */
const Impl::XLRelationships* Impl::XLWorkbook::Relationships() const {

    return m_relationships.get();
}

/**
 * @details
 */
XMLNode Impl::XLWorkbook::SheetNode(const string& sheetName) {

    auto sheet = m_sheetsNode.find_child_by_attribute("name", sheetName.c_str());
    if (!sheet)
        throw XLException("Sheet named " + sheetName + " does not exist.");
    return sheet;
}

/**
 * @details
 */
void Impl::XLWorkbook::CreateSharedStrings(const XLRelationshipItem& item) {

    m_sharedStrings.reset(new XLSharedStrings(*ParentDocument(), "xl/" + item.Target()));
    m_childXmlDocuments[m_sharedStrings->FilePath()] = m_sharedStrings.get();
}

/**
 * @details
 */
void Impl::XLWorkbook::CreateStyles(const XLRelationshipItem& item) {

    m_styles.reset(new XLStyles(*ParentDocument(), "xl/" + item.Target()));
    m_childXmlDocuments[m_styles->FilePath()] = m_styles.get();
}

/**
 * @details
 */
void Impl::XLWorkbook::CreateWorksheet(const XLRelationshipItem& item,
                                       const std::string& xmlData) {

    // Find the appropriate sheet node in the Workbook .xml file; get the name and id of the worksheet.
    string name;
    for (auto& node : m_sheetsNode.children()) {
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
        m_sheets[name]                             = nullptr;
        m_worksheets[name]                         = nullptr;
        m_childXmlDocuments["xl/" + item.Target()] = nullptr;
    }
    else {
        m_sheets[name]                             = make_unique<XLWorksheet>(*this, name, "xl/" + item.Target(),
                                                                              xmlData);
        m_worksheets[name]                         = m_sheets.at(name).get();
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
void Impl::XLWorkbook::CreateChartsheet(const XLRelationshipItem& item) {

    string name;
    auto   node = m_sheetsNode.first_child();
    while (node) {
        if (strcmp(node.attribute("r:id").value(),
                   item.Id().c_str()) == 0) {
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

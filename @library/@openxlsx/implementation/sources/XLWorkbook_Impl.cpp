//
// Created by Troldal on 02/09/16.
//

#include <cstring>
#include <pugixml.hpp>
#include <XLWorkbook_Impl.h>

#include "XLWorksheet_Impl.h"
#include "XLChartsheet_Impl.h"

using namespace std;
using namespace OpenXLSX;

/**
 * @details The constructor initializes the member variables and calls the loadXMLData from the
 * XLAbstractXMLFile base class.
 */
Impl::XLWorkbook::XLWorkbook(XLDocument& parent, const std::string& filePath)

        : XLAbstractXMLFile(parent, filePath),
          XLSpreadsheetElement(parent),
          m_sheetId(0),
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
    m_sheetsNode   = XmlDocument()->first_child().child("sheets");
    m_definedNames = XmlDocument()->first_child().child("definedNames");

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
 */
Impl::XLSheet* Impl::XLWorkbook::Sheet(const std::string& sheetName) {

    //    return const_cast<Impl::XLSheet*>(static_cast<const Impl::XLWorkbook*>(this)->Sheet(sheetName));

    auto sheetData = find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& data) {
        return sheetName == data.sheetName.value();
    });

    if (sheetData == m_sheets.end())
        throw XLException("Sheet \"" + sheetName + "\" does not exist");
    if (sheetData->sheetItem == nullptr) {
        switch (sheetData->sheetType) {
            case XLSheetType::WorkSheet:
                sheetData->sheetItem = make_unique<XLWorksheet>(*this, sheetData->sheetName, sheetData->sheetPath);
                break;

            case XLSheetType::ChartSheet:
                //sheetData->sheetItem = make_unique<XLChartsheet>(*this, sheetData->sheetName, sheetData->sheetPath);
                break;

            default:
                throw XLException("Unknown sheet type");
        }

        m_childXmlDocuments.at(sheetData->sheetPath) = sheetData->sheetItem.get();
    }

    return sheetData->sheetItem.get();
}

/**
 * @details
 */
const Impl::XLSheet* Impl::XLWorkbook::Sheet(const std::string& sheetName) const {

    auto sheetData = find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& data) {
        return sheetName == data.sheetName.value();
    });

    if (sheetData == m_sheets.end())
        throw XLException("Sheet \"" + sheetName + "\" does not exist");
    if (sheetData->sheetItem == nullptr) {
        switch (sheetData->sheetType) {
            case XLSheetType::WorkSheet:
                //sheetData->sheetItem = make_unique<XLWorksheet>(*this, sheetData->sheetName, sheetData->sheetPath);
                break;

            case XLSheetType::ChartSheet:
                //sheetData->sheetItem = make_unique<XLChartsheet>(*this, sheetData->sheetName, sheetData->sheetPath);
                break;

            default:
                throw XLException("Unknown sheet type");
        }

        m_childXmlDocuments.at(sheetData->sheetPath) = sheetData->sheetItem.get();
    }

    return sheetData->sheetItem.get();
}

/**
 * @details
 * @todo Ensure to throw if index is invalid.
 */
Impl::XLSheet* Impl::XLWorkbook::Sheet(unsigned int index) {

    //    return const_cast<Impl::XLSheet*>(static_cast<const Impl::XLWorkbook*>(this)->Sheet(index));
    string name = vector(m_sheetsNode.begin(), m_sheetsNode.end())[index - 1].attribute("name").as_string();
    return Sheet(name);
}

/**
 * @details
 */
const Impl::XLSheet* Impl::XLWorkbook::Sheet(unsigned int index) const {

    string name = vector(m_sheetsNode.begin(), m_sheetsNode.end())[index - 1].attribute("name").as_string();
    return Sheet(name);
}

/**
 * @details
 * @todo Currently, an XLWorksheet object cannot be created from the const method. This should be fixed.
 */
Impl::XLWorksheet* Impl::XLWorkbook::Worksheet(const std::string& sheetName) {

    auto sheetData = find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& data) {
        return (sheetName == data.sheetName.value() && data.sheetType == XLSheetType::WorkSheet);
    });

    if (sheetData == m_sheets.end())
        throw XLException("Worksheet \"" + sheetName + "\" does not exist");

    return dynamic_cast<XLWorksheet*>(Sheet(sheetName));
}

/**
 * @details
 */
const Impl::XLWorksheet* Impl::XLWorkbook::Worksheet(const std::string& sheetName) const {

    auto sheetData = find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& data) {
        return (sheetName == data.sheetName.value() && data.sheetType == XLSheetType::WorkSheet);
    });

    if (sheetData == m_sheets.end())
        throw XLException("Worksheet \"" + sheetName + "\" does not exist");

    return dynamic_cast<const XLWorksheet*>(Sheet(sheetName));
}

/**
 * @details
 */
Impl::XLChartsheet* Impl::XLWorkbook::Chartsheet(const std::string& sheetName) {

    auto sheetData = find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& data) {
        return (sheetName == data.sheetName.value() && data.sheetType == XLSheetType::ChartSheet);
    });

    if (sheetData == m_sheets.end())
        throw XLException("Chartsheet \"" + sheetName + "\" does not exist");

    return dynamic_cast<XLChartsheet*>(Sheet(sheetName));
}

/**
 * @details
 */
const Impl::XLChartsheet* Impl::XLWorkbook::Chartsheet(const std::string& sheetName) const {

    auto sheetData = find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& data) {
        return (sheetName == data.sheetName.value() && data.sheetType == XLSheetType::ChartSheet);
    });

    if (sheetData == m_sheets.end())
        throw XLException("Chartsheet \"" + sheetName + "\" does not exist");

    return dynamic_cast<const XLChartsheet*>(Sheet(sheetName));
}

/**
 * @details
 */
bool Impl::XLWorkbook::HasSharedStrings() const {

    return m_sharedStrings != nullptr;
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
    m_sheets.erase(find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& item) { return sheetName == item.sheetName.value(); }));

    // Delete the node from Workbook.xml
    m_sheetsNode.remove_child(m_sheetsNode.find_child_by_attribute("name", sheetName.c_str()));
}

/**
 * @details
 */
void Impl::XLWorkbook::AddWorksheet(const std::string& sheetName, unsigned int index) {

    CreateWorksheet(*InitiateWorksheet(sheetName, index), XLWorksheet::NewSheetXmlData());
}

/**
 * @details
 * @todo If the original sheet's tebSelected attribute is set, ensure it is un-set in the clone.
 */
void Impl::XLWorkbook::CloneWorksheet(const std::string& extName, const std::string& newName, unsigned int index) {

    CreateWorksheet(*InitiateWorksheet(newName, index), Worksheet(extName)->GetXmlData());
}

/**
 * @details
 */
Impl::XLRelationshipItem* Impl::XLWorkbook::InitiateWorksheet(const std::string& sheetName, unsigned int index) {

    auto        sheetID       = GetNewSheetID();
    std::string worksheetPath = "/xl/worksheets/sheet" + to_string(sheetID) + ".xml";

    // Add content item to document
    ParentDocument()->AddContentItem(worksheetPath, XLContentType::Worksheet);

    // Add relationship item
    XLRelationshipItem& item = *m_relationships
            ->AddRelationship(XLRelationshipType::Worksheet, "worksheets/sheet" + to_string(sheetID) + ".xml");

    // insert Sheet node at the given Index

    auto node  = XMLNode();
    auto nodes = vector(m_sheetsNode.begin(), m_sheetsNode.end());

    if (index == 0 || index > nodes.size())
        node = m_sheetsNode.append_child("sheet");
    else
        node = m_sheetsNode.insert_child_before("sheet", nodes[index - 1]);

    node.append_attribute("name")    = sheetName.c_str();
    node.append_attribute("sheetId") = to_string(sheetID).c_str();
    node.append_attribute("r:id")    = item.Id().c_str();

    // Add entry to the App Properties
    if (index == 0)
        ParentDocument()->AppProperties()->InsertSheetName(sheetName, WorksheetCount() + 1);
    else
        ParentDocument()->AppProperties()->InsertSheetName(sheetName, index);

    ParentDocument()->AppProperties()->SetHeadingPair("Worksheets", WorksheetCount() + 1);

    return &item;
}

int Impl::XLWorkbook::GetNewSheetID() {

    return ++m_sheetId;
}

/**
 * @details
 * @todo To be implemented.
 */
void Impl::XLWorkbook::AddChartsheet(const std::string& sheetName, unsigned int index) {

}

/**
 * @details
 * @todo To be implemented.
 */
void Impl::XLWorkbook::MoveSheet(const std::string& sheetName, unsigned int index) {

    auto node = m_sheetsNode.find_child_by_attribute("name", sheetName.c_str());

    if (index <= 1)
        m_sheetsNode.prepend_move(node);
    else if (index >= SheetCount())
        m_sheetsNode.append_move(node);
    else {
        auto     current = m_sheetsNode.first_child();
        for (int i       = 1; i < index; ++i)
            current = current.next_sibling();
        m_sheetsNode.insert_move_before(node, current);
    }
}

/**
 * @details
 */
unsigned int Impl::XLWorkbook::IndexOfSheet(const std::string& sheetName) const {

    auto node = m_sheetsNode.first_child();
    if (node.type() == pugi::node_null)
        throw runtime_error("Sheet does not exist.");

    unsigned int index = 1;
    while (node != nullptr) {
        if (strcmp(node.attribute("name").value(), sheetName.c_str()) == 0)
            break;
        node = node.next_sibling();
        if (!node)
            throw runtime_error("Sheet does not exist.");
        ++index;
    }

    return index;
}

XLSheetType Impl::XLWorkbook::TypeOfSheet(const std::string& sheetName) const {

    auto sheetData = find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& data) {
        return sheetName == data.sheetName.value();
    });

    if (sheetData == m_sheets.end())
        throw XLException("Sheet \"" + sheetName + "\" does not exist");

    return sheetData->sheetType;
}

XLSheetType Impl::XLWorkbook::TypeOfSheet(unsigned int index) const {

    string name = vector(m_sheetsNode.begin(), m_sheetsNode.end())[index - 1].attribute("name").as_string();
    return TypeOfSheet(name);
}

/**
 * @details
 */
unsigned int Impl::XLWorkbook::SheetCount() const {

    return m_sheets.size();
}

/**
 * @details
 */
unsigned int Impl::XLWorkbook::WorksheetCount() const {

    return count_if(m_sheets.begin(), m_sheets.end(), [](const XLSheetData& iter) {
        return iter.sheetType == XLSheetType::WorkSheet;
    });
}

/**
 * @details
 */
unsigned int Impl::XLWorkbook::ChartsheetCount() const {

    return count_if(m_sheets.begin(), m_sheets.end(), [](const XLSheetData& iter) {
        return iter.sheetType == XLSheetType::ChartSheet;
    });
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

    return find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& item) {
        return sheetName == item.sheetName.value();
    }) != m_sheets.end();
}

/**
 * @details
 */
bool Impl::XLWorkbook::WorksheetExists(const std::string& sheetName) const {

    return find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& item) {
        return (sheetName == item.sheetName.value() && item.sheetType == XLSheetType::WorkSheet);
    }) != m_sheets.end();
}

/**
 * @details
 */
bool Impl::XLWorkbook::ChartsheetExists(const std::string& sheetName) const {

    return find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& item) {
        return (sheetName == item.sheetName.value() && item.sheetType == XLSheetType::ChartSheet);
    }) != m_sheets.end();
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
void Impl::XLWorkbook::CreateWorksheet(const XLRelationshipItem& item, const std::string& xmlData) {

    if (m_sheetsNode.find_child_by_attribute("r:id", item.Id().c_str()) == nullptr)
        throw XLException("Invalid sheet ID");

    // Find the appropriate sheet node in the Workbook .xml file; get the name and id of the worksheet.
    // If xmlData is empty, set the m_sheets and m_childXmlDocuments elements to nullptr. The worksheet will then be
    // lazy-instantiated when the worksheet is requested using the 'Worksheet" function.
    // If xmlData is provided (i.e. a new sheet is created or an existing is cloned), create the new worksheets accordingly.
    auto& sheet = m_sheets.emplace_back();
    sheet.sheetName = m_sheetsNode.find_child_by_attribute("r:id", item.Id().c_str()).attribute("name");
    sheet.sheetPath = "xl/" + item.Target();
    sheet.sheetType = XLSheetType::WorkSheet;
    sheet.sheetItem = (xmlData.empty() ? nullptr : make_unique<XLWorksheet>(*this, sheet.sheetName, sheet.sheetPath, xmlData));

    m_childXmlDocuments[sheet.sheetPath] = sheet.sheetItem.get();

    // Update internal counters.
    m_sheetId++;
}

/**
 * @details
 */
void Impl::XLWorkbook::CreateChartsheet(const XLRelationshipItem& item) {

    //TODO: Create Chartsheet object here.

    m_sheetId++;
}

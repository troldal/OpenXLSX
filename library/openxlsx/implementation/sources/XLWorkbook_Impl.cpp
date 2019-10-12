//
// Created by Troldal on 02/09/16.
//

#include <algorithm>
#include <vector>
#include <utility>
#include <type_traits>

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
          m_sheetId(0),
          m_relationships(parent, "xl/_rels/workbook.xml.rels"),
          m_sharedStrings(parent),
          m_document(&parent) {

    ParseXMLData();
}

/**
 * @details The destructor is specifically defaulted in the implementation file. Otherwise compilation issues
 * will occur.
 */
Impl::XLWorkbook::~XLWorkbook() = default;

/**
 * @details
 */
bool Impl::XLWorkbook::ParseXMLData() {

    // ===== Find the "sheets" section in the Workbook.xml file
    m_sheetsNode = XmlDocument()->first_child().child("sheets");
    m_definedNamesNode = XmlDocument()->first_child().child("definedNames");

    // ===== Iterate through the relationship items and handle them accordingly
    for (const auto& relationship : m_relationships.Relationships()) {
        string path = relationship->Target().value();

        switch (relationship->Type()) {

            // ===== Set up Worksheet files
            case XLRelationshipType::Worksheet :
                CreateWorksheet(*relationship);
                break;

                // ===== Set up Chartsheet files
            case XLRelationshipType::ChartSheet :
                CreateChartsheet(*relationship);
                break;

                // ===== Default handler
            default:
                break;
        }
    }

    // ===== Read all defined names
    // ===== TODO: Consider factoring this out.
    if (m_definedNamesNode) {
        for (const auto& name : m_definedNamesNode.children()) {
            auto& definedName = m_definedNames.emplace_back(XLDefinedName());
            definedName.definedNameNode = name;
            definedName.name = name.attribute("name");
            definedName.localSheetId = name.attribute("localSheetId");

            // ===== If the defined name is scoped to a single worksheet, find the corresponding worksheet node
            if (definedName.localSheetId)
                definedName.sheetNode = GetChildByIndex(m_sheetsNode, definedName.localSheetId.as_uint());
        }
    }

    // ===== Find the sheet node corresponding to the active sheet
    // ===== TODO: Consider factoring this out
    if (!XmlDocument()->first_child().child("bookViews").first_child().attribute("activeTab")) {
        XmlDocument()->first_child().child("bookViews").first_child().append_attribute("activeTab").set_value(0);
    }
    unsigned int activeTab = XmlDocument()->first_child().child("bookViews").first_child().attribute("activeTab").as_uint();
    m_activeSheet = GetChildByIndex(m_sheetsNode, activeTab);

    return true;
}

/**
 * @details The function uses the Scott Meyers const_cast trick to call the const version of the Sheet function.
 */
Impl::XLSheet* Impl::XLWorkbook::Sheet(const std::string& sheetName) {

        return const_cast<Impl::XLSheet*>(as_const(*this).Sheet(sheetName));
}

/**
 * @details
 */
const Impl::XLSheet* Impl::XLWorkbook::Sheet(const std::string& sheetName) const {

    // ===== Find the sheet data corresponding to the sheet with the requested name
    auto sheetData = find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& data) {
        return sheetName == data.sheetNode.attribute("name").value();
    });

    // ===== If no sheet with the requested name exists, throw an exception.
    if (sheetData == m_sheets.end())
        throw XLException("Sheet \"" + sheetName + "\" does not exist");

    // ===== If the sheet has not been loaded, load it into memory
    if (sheetData->sheetItem == nullptr) {
        switch (sheetData->sheetType) {

            // ===== Handler for worksheets
            case XLSheetType::WorkSheet: // TODO: The const_cast here is REALLY ugly. Find a way to eliminate this somehow!
                sheetData->sheetItem = make_unique<XLWorksheet>(const_cast<XLWorkbook&>(*this),
                                                                sheetData->sheetNode.attribute("name"),
                                                                sheetData->sheetRelationship.Target().value());
                break;

            // ===== Handler for chartsheets
            case XLSheetType::ChartSheet:
                // TODO: Implement creation of chartsheet
                break;

            // ===== Default handler
            default:
                throw XLException("Unknown sheet type");
        }
    }

    return sheetData->sheetItem.get();
}

/**
 * @details The function uses the Scott Meyers const_cast trick to call the const version of the Sheet function.
 */
Impl::XLSheet* Impl::XLWorkbook::Sheet(unsigned int index) {

    return const_cast<Impl::XLSheet*>(as_const(*this).Sheet(index));
}

/**
 * @details Create a vector with sheet nodes, retrieve the node at the requested index, get sheet name and return the
 * corresponding sheet object.
 * @todo Throw if index is out of bounds
 */
const Impl::XLSheet* Impl::XLWorkbook::Sheet(unsigned int index) const {

    return Sheet(vector<XMLNode>(m_sheetsNode.begin(), m_sheetsNode.end())[index - 1].attribute("name").as_string());
}

/**
 * @details The function uses the Scott Meyers const_cast trick to call the const version of the Worksheet function.
 */
Impl::XLWorksheet* Impl::XLWorkbook::Worksheet(const std::string& sheetName) {

    return const_cast<Impl::XLWorksheet*>(as_const(*this).Worksheet(sheetName));
}

/**
 * @details
 */
const Impl::XLWorksheet* Impl::XLWorkbook::Worksheet(const std::string& sheetName) const {

    // ===== Find the worksheet data corresponding to the sheet with the requested name
    auto sheetData = find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& data) {
        return (sheetName == data.sheetNode.attribute("name").value() && data.sheetType == XLSheetType::WorkSheet);
    });

    // ===== If no sheet with the requested name exists, throw an exception.
    if (sheetData == m_sheets.end())
        throw XLException("Worksheet \"" + sheetName + "\" does not exist");

    return dynamic_cast<const XLWorksheet*>(Sheet(sheetName));
}

/**
 * @details The function uses the Scott Meyers const_cast trick to call the const version of the Chartsheet function.
 */
Impl::XLChartsheet* Impl::XLWorkbook::Chartsheet(const std::string& sheetName) {

    return const_cast<Impl::XLChartsheet*>(as_const(*this).Chartsheet(sheetName));
}

/**
 * @details
 */
const Impl::XLChartsheet* Impl::XLWorkbook::Chartsheet(const std::string& sheetName) const {

    // ===== Find the chartsheet data corresponding to the sheet with the requested name
    auto sheetData = find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& data) {
        return (sheetName == data.sheetNode.attribute("name").value() && data.sheetType == XLSheetType::ChartSheet);
    });

    // ===== If no sheet with the requested name exists, throw an exception.
    if (sheetData == m_sheets.end())
        throw XLException("Chartsheet \"" + sheetName + "\" does not exist");

    return dynamic_cast<const XLChartsheet*>(Sheet(sheetName));
}

/**
 * @details
 */
bool Impl::XLWorkbook::HasSharedStrings() const {

    return m_sharedStrings; //implicit conversion to bool
}

/**
 * @details
 */
Impl::XLSharedStrings* Impl::XLWorkbook::SharedStrings() const {

    return &m_sharedStrings;
}

/**
 * @details
 */
void Impl::XLWorkbook::DeleteNamedRanges() {

    for (auto& child : m_definedNamesNode.children())
        child.parent().remove_child(child);
}

/**
 * @details
 */
void Impl::XLWorkbook::DeleteSheet(const std::string& sheetName) {

    auto worksheetCount = count_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& data) {
        return data.sheetType == XLSheetType::WorkSheet;
    });

    // ===== Find the relevant XLSheetData item in the internal data structure
    auto sheetData = find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& data) {
        return sheetName == data.sheetNode.attribute("name").value();
    });

    // ===== If this is the last worksheet in the workbook, throw an exception.
    if (worksheetCount == 1 && sheetData->sheetType == XLSheetType::WorkSheet)
        throw XLException("Invalid operation. There must be at least one worksheet in the workbook.");

    // ===== Delete references to the sheet in the .xml files
    Document()->AppProperties()->DeleteSheetName(sheetName);
    Document()->DeleteXMLFile(sheetData->sheetContentItem.Path().substr(1));
    Document()->DeleteContentItem(sheetData->sheetContentItem);
    Relationships()->DeleteRelationship(sheetData->sheetRelationship);

    // ===== Delete the XLSheetData item from the internal data structure
    m_sheets.erase(find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& item) {
        return sheetName == item.sheetNode.attribute("name").value();
    }));

    // ===== Delete the node from Workbook.xml
    m_sheetsNode.remove_child(m_sheetsNode.find_child_by_attribute("name", sheetName.c_str()));

    // ===== Update the activeTab attribute. If the deleted sheet was the active sheet, set the property to zero
    if (!m_activeSheet) {
        XmlDocument()->first_child().child("bookViews").first_child().attribute("activeTab").set_value(0);
        m_activeSheet = m_sheetsNode.first_child();
    }
    else {
        unsigned int index = 0;
        for (auto& item : m_sheetsNode.children()) {
            if (m_activeSheet == item) {
                XmlDocument()->first_child().child("bookViews").first_child().attribute("activeTab").set_value(index);
                break;
            }
            index++;
        }
    }
}

/**
 * @details
 */
void Impl::XLWorkbook::AddWorksheet(const std::string& sheetName, unsigned int index) {

    CreateWorksheet(*InitiateWorksheet(sheetName, index), XLWorksheet::NewSheetXmlData());
}

/**
 * @details
 * @todo If the original sheet's tabSelected attribute is set, ensure it is un-set in the clone.
 */
void Impl::XLWorkbook::CloneWorksheet(const std::string& extName, const std::string& newName, unsigned int index) {

    CreateWorksheet(*InitiateWorksheet(newName, index), Worksheet(extName)->GetXmlData());
}

/**
 * @details
 */
Impl::XLRelationshipItem* Impl::XLWorkbook::InitiateWorksheet(const std::string& sheetName, unsigned int index) {

    auto sheetID = GetNewSheetID();
    std::string worksheetPath = "/xl/worksheets/sheet" + to_string(sheetID) + ".xml";

    // Add content item to document
    Document()->AddContentItem(worksheetPath, XLContentType::Worksheet);

    // Add relationship item
    XLRelationshipItem& item = *m_relationships.AddRelationship(XLRelationshipType::Worksheet,
                                                                "worksheets/sheet" + to_string(sheetID) + ".xml");

    // insert Sheet node at the given Index

    auto node = XMLNode();
    auto nodes = vector<XMLNode>(m_sheetsNode.begin(), m_sheetsNode.end());
    if (index == 0 || index > nodes.size())
        node = m_sheetsNode.append_child("sheet");
    else
        node = m_sheetsNode.insert_child_before("sheet", nodes[index - 1]);

    node.append_attribute("name") = sheetName.c_str();
    node.append_attribute("sheetId") = to_string(sheetID).c_str();
    node.append_attribute("r:id") = item.Id().value();

    // Add entry to the App Properties
    if (index == 0)
        Document()->AppProperties()->InsertSheetName(sheetName, WorksheetCount() + 1);
    else
        Document()->AppProperties()->InsertSheetName(sheetName, index);

    Document()->AppProperties()->SetHeadingPair("Worksheets", WorksheetCount() + 1);

    return &item;
}

int Impl::XLWorkbook::GetNewSheetID() {

    return max_element(m_sheets.begin(), m_sheets.end(), [](const XLSheetData& a, const XLSheetData& b) {
        return a.sheetNode.attribute("sheetId").as_int() < b.sheetNode.attribute("sheetId").as_int();
    })->sheetNode.attribute("sheetId").as_int() + 1;
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
void Impl::XLWorkbook::MoveSheet(const std::string& sheetName, unsigned int newIndex) {

    // ===== Checking that the input is valid
    if (newIndex < 1)
        newIndex = 1;
    if (newIndex > m_sheets.size())
        newIndex = m_sheets.size();

    // ===== If the new index is equal to the current, don't do anything
    unsigned int oldIndex = 1;
    for (const auto& item : m_sheets) {
        if (sheetName == item.sheetNode.attribute("name").value())
            break;
        ++oldIndex;
    }

    if (newIndex == oldIndex)
        return;

    // ===== Modify the node in the XML file
    auto node = m_sheetsNode.find_child_by_attribute("name", sheetName.c_str());

    if (newIndex <= 1)
        m_sheetsNode.prepend_move(node);
    else if (newIndex >= SheetCount())
        m_sheetsNode.append_move(node);
    else {
        auto currentSheet = m_sheetsNode.first_child();
        auto currentIndex = 1;
        while (currentIndex < newIndex) {
            currentSheet = currentSheet.next_sibling();
            ++currentIndex;
        }
        if (oldIndex > newIndex)
            m_sheetsNode.insert_move_before(node, currentSheet);
        else
            m_sheetsNode.insert_move_after(node, currentSheet);
    }

    // ===== Move the element to the right location in the vector
    auto first = m_sheets.begin() + min(newIndex, oldIndex) - 1;
    auto last = m_sheets.begin() + max(newIndex, oldIndex);
    auto n_first = m_sheets.begin() + (oldIndex > newIndex ? oldIndex - 1 : oldIndex);
    rotate(first, n_first, last);

    // ===== Updated defined names with worksheet scopes.
    for (auto& definedName : m_definedNames) {
        if (definedName.localSheetId) {
            definedName.localSheetId.set_value(IndexOfSheet(definedName.sheetNode.attribute("name").value()) - 1);
        }
    }

    // ===== Update the activeTab attribute.
    unsigned int index = 0;
    for (auto& item : m_sheetsNode.children()) {
        if (m_activeSheet == item) {
            XmlDocument()->first_child().child("bookViews").first_child().attribute("activeTab").set_value(index);
            break;
        }
        index++;
    }
}

/**
 * @details
 */
unsigned int Impl::XLWorkbook::IndexOfSheet(const std::string& sheetName) const {

    // ===== Iterate through sheet nodes. When a match is found, return the index;
    unsigned int index = 1;
    for (auto& sheet : m_sheetsNode.children()) {
        if (sheetName == sheet.attribute("name").value())
            return index;
        index++;
    }

    // ===== If a match is not found, throw an exception.
    throw XLException("Sheet does not exist");

}

XLSheetType Impl::XLWorkbook::TypeOfSheet(const std::string& sheetName) const {

    auto sheetData = find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& data) {
        return sheetName == data.sheetNode.attribute("name").value();
    });

    if (sheetData == m_sheets.end())
        throw XLException("Sheet \"" + sheetName + "\" does not exist");

    return sheetData->sheetType;
}

XLSheetType Impl::XLWorkbook::TypeOfSheet(unsigned int index) const {

    string name = vector<XMLNode>(m_sheetsNode.begin(), m_sheetsNode.end())[index - 1].attribute("name").as_string();
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
std::vector<std::string> Impl::XLWorkbook::SheetNames() const {

    vector<string> result;

    for (const auto& item : m_sheets)
        result.emplace_back(item.sheetNode.attribute("name").value());

    return result;
}

/**
 * @details
 */
std::vector<std::string> Impl::XLWorkbook::WorksheetNames() const {

    vector<string> result;

    for (const auto& item : m_sheets)
        if (item.sheetType == XLSheetType::WorkSheet)
            result.emplace_back(item.sheetNode.attribute("name").value());

    return result;
}

/**
 * @details
 */
std::vector<std::string> Impl::XLWorkbook::ChartsheetNames() const {

    vector<string> result;

    for (const auto& item : m_sheets)
        if (item.sheetType == XLSheetType::ChartSheet)
            result.emplace_back(item.sheetNode.attribute("name").value());

    return result;
}

Impl::XLDocument* Impl::XLWorkbook::Document() {

    return m_document;
}

const Impl::XLDocument* Impl::XLWorkbook::Document() const {

    return m_document;
}

/**
 * @details
 */
bool Impl::XLWorkbook::SheetExists(const std::string& sheetName) const {

    return find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& item) {
        return sheetName == item.sheetNode.attribute("name").value();
    }) != m_sheets.end();
}

/**
 * @details
 */
bool Impl::XLWorkbook::WorksheetExists(const std::string& sheetName) const {

    return find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& item) {
        return (sheetName == item.sheetNode.attribute("name").value() && item.sheetType == XLSheetType::WorkSheet);
    }) != m_sheets.end();
}

/**
 * @details
 */
bool Impl::XLWorkbook::ChartsheetExists(const std::string& sheetName) const {

    return find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& item) {
        return (sheetName == item.sheetNode.attribute("name").value() && item.sheetType == XLSheetType::ChartSheet);
    }) != m_sheets.end();
}

/**
 * @details The UpdateSheetName member function searches theroug the usages of the old name and replaces with the
 * new sheet name.
 * @todo Currently, this function only searches through defined names. Consider using this function to update the
 * actual sheet name as well.
 */
void Impl::XLWorkbook::UpdateSheetName(const std::string& oldName, const std::string& newName) {

    for (auto& sheet : m_sheets) {
        if (sheet.sheetType == XLSheetType::WorkSheet)
            Worksheet(sheet.sheetNode.attribute("name").value())->UpdateSheetName(oldName, newName);
    }

    // ===== Set up temporary variables
    std::string oldNameTemp = oldName;
    std::string newNameTemp = newName;
    std::string formula;

    // ===== If the sheet name contains spaces, it should be enclosed in single quotes (')
    if (oldName.find(' ') != std::string::npos)
        oldNameTemp = "\'" + oldName + "\'";
    if (newName.find(' ') != std::string::npos)
        newNameTemp = "\'" + newName + "\'";

    // ===== Ensure only sheet names are replaced (references to sheets always ends with a '!')
    oldNameTemp += '!';
    newNameTemp += '!';

    // ===== Iterate through all defined names
    for (auto& definedName : m_definedNames) {
        formula = definedName.definedNameNode.text().get();

        // ===== Skip if formula contains a '[' and ']' (means that the defined refers to external workbook)
        if (formula.find('[') == string::npos && formula.find(']') == string::npos) {

            // ===== For all instances of the old sheet name in the formula, replace with the new name.
            while (formula.find(oldNameTemp) != string::npos) {
                formula.replace(formula.find(oldNameTemp), oldNameTemp.length(), newNameTemp);
            }
            definedName.definedNameNode.text().set(formula.c_str());
        }
    }
}

/**
 * @details
 */
Impl::XLRelationships* Impl::XLWorkbook::Relationships() {

    return &m_relationships;
}

/**
 * @details
 */
const Impl::XLRelationships* Impl::XLWorkbook::Relationships() const {

    return &m_relationships;
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
void Impl::XLWorkbook::CreateWorksheet(const XLRelationshipItem& item, const std::string& xmlData) {

    if (m_sheetsNode.find_child_by_attribute("r:id", item.Id().value()) == nullptr)
        throw XLException("Invalid sheet ID");

    // Find the appropriate sheet node in the Workbook .xml file; get the name and id of the worksheet.
    // If xmlData is empty, set the m_sheets and m_childXmlDocuments elements to nullptr. The worksheet will then be
    // lazy-instantiated when the worksheet is requested using the 'Worksheet" function.
    // If xmlData is provided (i.e. a new sheet is created or an existing is cloned), create the new worksheets accordingly.

    unsigned int index = 0;
    for (const auto& elem : m_sheetsNode.children()) {
        if (string(item.Id().value()) == elem.attribute("r:id").value())
            break;
        ++index;
    }

    if (index > m_sheets.size())
        index = m_sheets.size();

    auto& sheet = *m_sheets.insert(m_sheets.begin() + index, XLSheetData());
    sheet.sheetNode = m_sheetsNode.find_child_by_attribute("r:id", item.Id().value());
    sheet.sheetRelationship = item;
    sheet.sheetContentItem = Document()->ContentItem(string("/xl/") + item.Target().value());
    sheet.sheetType = XLSheetType::WorkSheet;
    sheet.sheetItem = (xmlData.empty() ? nullptr : make_unique<XLWorksheet>(*this,
                                                                            sheet.sheetNode.attribute("name"),
                                                                            sheet.sheetRelationship.Target().value(),
                                                                            xmlData));

    // ===== Update the activeTab attribute. If the deleted sheet was the active sheet, set the property to zero
    unsigned int idx = 0;
    for (auto& it : m_sheetsNode.children()) {
        if (m_activeSheet == it) {
            XmlDocument()->first_child().child("bookViews").first_child().attribute("activeTab").set_value(idx);
            break;
        }
        index++;
    }
}

/**
 * @details
 */
void Impl::XLWorkbook::CreateChartsheet(const XLRelationshipItem& item) {

    //TODO: Create Chartsheet object here.
}

void Impl::XLWorkbook::WriteXMLData() {

    XLAbstractXMLFile::WriteXMLData();
    m_relationships.WriteXMLData();

    if (m_sharedStrings)
        m_sharedStrings.WriteXMLData();

    for (auto& sheet : m_sheets)
        if (sheet.sheetItem)
            sheet.sheetItem->WriteXMLData();
}


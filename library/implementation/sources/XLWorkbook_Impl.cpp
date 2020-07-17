//
// Created by Troldal on 02/09/16.
//

#include "XLWorkbook_Impl.hpp"

#include "XLChartsheet_Impl.hpp"
#include "XLDocument_Impl.hpp"
#include "XLSheet_Impl.hpp"
#include "XLWorksheet_Impl.hpp"

#include <algorithm>
#include <iterator>
#include <utility>
#include <vector>

using namespace std;
using namespace OpenXLSX;

namespace
{
    std::string getNewSheetXmlData()
    {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
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
               "</worksheet>";
    }

    inline XMLNode getSheetNodeByRID(XMLNode sheetsNode, const std::string& sheetRID)
    {
        return sheetsNode.find_child_by_attribute("r:id", sheetRID.c_str());
    }

    inline XMLNode getSheetNodeByID(XMLNode sheetsNode, uint16_t sheetID)
    {
        return sheetsNode.find_child_by_attribute("sheetId", std::to_string(sheetID).c_str());
    }

    inline XMLNode getSheetNodeByName(XMLNode sheetsNode, const std::string& sheetName)
    {
        return sheetsNode.find_child_by_attribute("name", sheetName.c_str());
    }

    inline uint16_t getSheetIndexByRID(XMLNode sheetsNode, const std::string& sheetRID)
    {
        auto sheetNode = getSheetNodeByRID(sheetsNode, sheetRID);
        if (!sheetsNode) throw XLException("Sheet does not exist");

        for (auto iter = sheetsNode.children().begin(); iter != sheetsNode.children().end(); ++iter)
            if (*iter == sheetNode) return std::distance(sheetNode.children().begin(), iter);
        throw XLException("Sheet with rID \"" + sheetRID + "\" does not exist.");
    }

    inline uint16_t getSheetIndexByID(XMLNode sheetsNode, uint16_t sheetID)
    {
        auto sheetNode = getSheetNodeByID(sheetsNode, sheetID);
        if (!sheetsNode) throw XLException("Sheet does not exist");

        for (auto iter = sheetsNode.children().begin(); iter != sheetsNode.children().end(); ++iter)
            if (*iter == sheetNode) return std::distance(sheetNode.children().begin(), iter);
        throw XLException("Sheet with ID \"" + to_string(sheetID) + "\" does not exist.");
    }

    inline uint16_t getSheetIndexByName(XMLNode sheetsNode, const std::string& sheetName)
    {
        auto sheetNode = getSheetNodeByName(sheetsNode, sheetName);
        if (!sheetsNode) throw XLException("Sheet does not exist");

        for (auto iter = sheetsNode.children().begin(); iter != sheetsNode.children().end(); ++iter)
            if (*iter == sheetNode) return std::distance(sheetNode.children().begin(), iter);
        throw XLException("Sheet with name \"" + sheetName + "\" does not exist.");
    }

    inline std::string getSheetNameByRID(XMLNode sheetsNode, const std::string& sheetRID)
    {
        return getSheetNodeByRID(sheetsNode, sheetRID).attribute("name").value();
    }

    inline std::string getSheetRIDByName(XMLNode sheetsNode, const std::string& sheetName)
    {
        return getSheetNodeByName(sheetsNode, sheetName).attribute("r:id").value();
    }

    inline uint16_t GetNewSheetID(XMLNode sheetsNode)
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
Impl::XLWorkbook::XLWorkbook(XLXmlData* xmlData)

    : XLAbstractXMLFile(xmlData),
      m_sheetId(0),
      m_relationships(ParentDoc().getXmlData("xl/_rels/workbook.xml.rels")),
      m_sharedStrings(ParentDoc().getXmlData("xl/sharedStrings.xml"))
{
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
bool Impl::XLWorkbook::ParseXMLData()
{
    // ===== Iterate through the relationship items and handle them accordingly
    //    for (const auto& relationship : m_relationships.Relationships()) {
    //        string path = relationship.Target();
    //
    //        switch (relationship.Type()) {
    //        // ===== Set up Worksheet files
    //        case XLRelationshipType::Worksheet:
    //            CreateWorksheet(relationship);
    //            break;
    //
    //            // ===== Set up Chartsheet files
    //        case XLRelationshipType::ChartSheet:
    //            CreateChartsheet(relationship);
    //            break;
    //
    //            // ===== Default handler
    //        default: break;
    //        }
    //    }

    // ===== Find the sheet node corresponding to the active sheet
    // ===== TODO: Consider factoring this out
    //    if (!XmlDocument()->first_child().child("bookViews").first_child().attribute("activeTab")) {
    //        XmlDocument()->first_child().child("bookViews").first_child().append_attribute("activeTab").set_value(0);
    //    }
    //    unsigned int activeTab =
    //        XmlDocument()->first_child().child("bookViews").first_child().attribute("activeTab").as_uint();
    //    m_activeSheet = GetChildByIndex(getSheetsNode(), activeTab);

    return true;
}

/**
 * @details
 */
Impl::XLSheet Impl::XLWorkbook::Sheet(const std::string& sheetName)
{
    // ===== First determine if the sheet exists.
    if (XmlDocument().document_element().child("sheets").find_child_by_attribute("name", sheetName.c_str()) == nullptr)
        throw XLException("Sheet \"" + sheetName + "\" does not exist");

    // ===== Find the sheet data corresponding to the sheet with the requested name
    std::string xmlID =
        XmlDocument().document_element().child("sheets").find_child_by_attribute("name", sheetName.c_str()).attribute("r:id").value();
    std::string xmlPath = m_relationships.RelationshipByID(xmlID).Target();
    return XLSheet(ParentDoc().getXmlData("xl/" + xmlPath));

    //    auto sheetData = find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& data) {
    //        return sheetName == data.sheetNode.attribute("name").value();
    //    });

    //    // ===== If the sheet has not been loaded, load it into memory
    //    if (sheetData->sheetItem == nullptr) {
    //        switch (sheetData->sheetType) {
    //        // ===== Handler for worksheets
    //        case XLSheetType::WorkSheet:    // TODO: The const_cast here is REALLY ugly. Find a way to eliminate this
    //                                        // somehow!
    //            sheetData->sheetItem = make_unique<XLWorksheet>(
    //                &const_cast<XLWorkbook&>(*this).ParentDoc().getXmlData(sheetData->sheetRelationship.Target()),
    //                sheetData->sheetNode.attribute("r:id").value());
    //            break;
    //
    //        // ===== Handler for chartsheets
    //        case XLSheetType::ChartSheet:
    //            // TODO: Implement creation of chartsheet
    //            break;
    //
    //        // ===== Default handler
    //        default: throw XLException("Unknown sheet type");
    //        }
    //    }
    //
    //    return sheetData->sheetItem.get();
}

/**
 * @details Create a vector with sheet nodes, retrieve the node at the requested index, get sheet name and return the
 * corresponding sheet object.
 * @todo Throw if index is out of bounds
 */
Impl::XLSheet Impl::XLWorkbook::Sheet(unsigned int index)
{
    return Sheet(vector<XMLNode>(getSheetsNode().begin(), getSheetsNode().end())[index - 1].attribute("name").as_string());
}

/**
 * @details
 */
Impl::XLWorksheet Impl::XLWorkbook::Worksheet(const std::string& sheetName)
{
    return Sheet(sheetName).Get<XLWorksheet>();

    //    // ===== Find the worksheet data corresponding to the sheet with the requested name
    //    auto sheetData = find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& data) {
    //        return (sheetName == data.sheetNode.attribute("name").value() && data.sheetType ==
    //        XLSheetType::WorkSheet);
    //    });
    //
    //    // ===== If no sheet with the requested name exists, throw an exception.
    //    if (sheetData == m_sheets.end()) throw XLException("Worksheet \"" + sheetName + "\" does not exist");
    //
    //    return dynamic_cast<const XLWorksheet*>(Sheet(sheetName));
}

/**
 * @details
 */
Impl::XLChartsheet Impl::XLWorkbook::Chartsheet(const std::string& sheetName)
{
    return Sheet(sheetName).Get<XLChartsheet>();
    //
    //    // ===== Find the chartsheet data corresponding to the sheet with the requested name
    //    auto sheetData = find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& data) {
    //        return (sheetName == data.sheetNode.attribute("name").value() && data.sheetType ==
    //        XLSheetType::ChartSheet);
    //    });
    //
    //    // ===== If no sheet with the requested name exists, throw an exception.
    //    if (sheetData == m_sheets.end()) throw XLException("Chartsheet \"" + sheetName + "\" does not exist");
    //
    //    return dynamic_cast<const XLChartsheet*>(Sheet(sheetName));
}

/**
 * @details
 */
bool Impl::XLWorkbook::HasSharedStrings() const
{
    return true;    // m_sharedStrings;    // implicit conversion to bool
}

/**
 * @details
 */
Impl::XLSharedStrings* Impl::XLWorkbook::SharedStrings() const
{
    return &m_sharedStrings;
}

/**
 * @details
 */
void Impl::XLWorkbook::DeleteNamedRanges()
{
    for (auto& child : XmlDocument().document_element().child("definedNames").children()) child.parent().remove_child(child);
}

/**
 * @details
 */
void Impl::XLWorkbook::DeleteSheet(const std::string& sheetName)
{
    auto sheetID = getSheetsNode().find_child_by_attribute("name", sheetName.c_str()).attribute("r:id").value();

    auto worksheetCount = count_if(getSheetsNode().children().begin(), getSheetsNode().children().end(), [&](const XMLNode& item) {
        return ParentDoc().executeQuery(
                   XLQuery(XLQueryType::GetSheetType, XLQueryParams { { "sheetID", item.attribute("r:id").value() } })) == "WORKSHEET";
    });

    auto worksheetType = ParentDoc().executeQuery(XLQuery(
        XLQueryType::GetSheetType,
        XLQueryParams { { "sheetID", getSheetsNode().find_child_by_attribute("name", sheetName.c_str()).attribute("r:id").value() } }));

    // ===== If this is the last worksheet in the workbook, throw an exception.
    if (worksheetCount == 1 && worksheetType == "WORKSHEET")
        throw XLException("Invalid operation. There must be at least one worksheet in the workbook.");

    ParentDoc().executeCommand(
        XLCommand(XLCommandType::DeleteSheet, XLCommandParams { { "sheetName", sheetName }, { "sheetID", sheetID } }, ""));

    // ===== Delete the node from Workbook.xml
    getSheetsNode().remove_child(getSheetsNode().find_child_by_attribute("name", sheetName.c_str()));

    // ===== Update the activeTab attribute. If the deleted sheet was the active sheet, set the property to zero
    //        if (!m_activeSheet) {
    //            XmlDocument()->first_child().child("bookViews").first_child().attribute("activeTab").set_value(0);
    //            m_activeSheet = getSheetsNode().first_child();
    //        }
    //        else {
    //            unsigned int index = 0;
    //            for (auto& item : getSheetsNode().children()) {
    //                if (m_activeSheet == item) {
    //                    XmlDocument()->first_child().child("bookViews").first_child().attribute("activeTab").set_value(index);
    //                    break;
    //                }
    //                index++;
    //            }
    //        }
}

/**
 * @details
 */
void Impl::XLWorkbook::AddWorksheet(const std::string& sheetName, unsigned int index)
{
    // ===== If a sheet with the given name already exists, throw an exception.
    if (XmlDocument().document_element().child("sheets").find_child_by_attribute("name", sheetName.c_str()))
        throw XLException("Sheet named \"" + sheetName + "\" already exists.");

    // CreateWorksheet(*InitiateWorksheet(sheetName, index), getNewSheetXmlData());
    InitiateWorksheet(sheetName, index), getNewSheetXmlData();
}

/**
 * @details
 * @todo If the original sheet's tabSelected attribute is set, ensure it is un-set in the clone.
 */
void Impl::XLWorkbook::CloneWorksheet(const std::string& extName, const std::string& newName, unsigned int index)
{
    // ===== If a sheet with the given name already exists, throw an exception.
    if (XmlDocument().document_element().child("sheets").find_child_by_attribute("name", newName.c_str()))
        throw XLException("Sheet named \"" + newName + "\" already exists.");

    CreateWorksheet(*InitiateWorksheet(newName, index), Worksheet(extName).GetXmlData());
}

/**
 * @details
 */
Impl::XLRelationshipItem* Impl::XLWorkbook::InitiateWorksheet(const std::string& sheetName, unsigned int index)
{
    auto        sheetID       = GetNewSheetID(XmlDocument().document_element().child("sheets"));
    std::string worksheetPath = "/xl/worksheets/sheet" + to_string(sheetID) + ".xml";

    ParentDoc().executeCommand(
        XLCommand(XLCommandType::AddWorksheet,
                  XLCommandParams { { "sheetPath", worksheetPath }, { "sheetName", sheetName }, { "sheetIndex", to_string(index) } },
                  ""));

    auto node  = XMLNode();
    auto nodes = vector<XMLNode>(getSheetsNode().begin(), getSheetsNode().end());
    if (index == 0 || index > nodes.size())
        node = getSheetsNode().append_child("sheet");
    else
        node = getSheetsNode().insert_child_before("sheet", nodes[index - 1]);

    node.append_attribute("name")    = sheetName.c_str();
    node.append_attribute("sheetId") = to_string(sheetID).c_str();
    node.append_attribute("r:id") =
        ParentDoc().executeQuery(XLQuery(XLQueryType::GetSheetID, XLQueryParams { { "sheetPath", worksheetPath } })).c_str();

    // Add content item to document
    // ParentDoc().AddContentItem(worksheetPath, XLContentType::Worksheet);

    // Add relationship item
    //    XLRelationshipItem item = m_relationships.AddRelationship(XLRelationshipType::Worksheet,
    //                                                              "worksheets/sheet" + to_string(sheetID) + ".xml");

    // insert Sheet node at the given Index

    // Add entry to the App Properties
    //    if (index == 0)
    //        ParentDoc().AppProperties()->InsertSheetName(sheetName, WorksheetCount() + 1);
    //    else
    //        ParentDoc().AppProperties()->InsertSheetName(sheetName, index);
    //
    //    ParentDoc().AppProperties()->SetHeadingPair("Worksheets", WorksheetCount() + 1);

    return nullptr;
}

XMLNode Impl::XLWorkbook::getSheetsNode() const
{
    return XmlDocument().first_child().child("sheets");
}

void Impl::XLWorkbook::setSheetName(const string& sheetRID, const string& newName)
{
    auto sheetName = XmlDocument().document_element().child("sheets").find_child_by_attribute("r:id", sheetRID.c_str()).attribute("name");

    UpdateSheetName(sheetName.value(), newName);
    sheetName.set_value(newName.c_str());
}

void Impl::XLWorkbook::setSheetVisibility(const string& sheetRID, const string& state)
{
    auto stateAttribute =
        XmlDocument().document_element().child("sheets").find_child_by_attribute("r:id", sheetRID.c_str()).attribute("state");

    if (!stateAttribute) {
        stateAttribute =
            XmlDocument().document_element().child("sheets").find_child_by_attribute("r:id", sheetRID.c_str()).append_attribute("state");
    }

    stateAttribute.set_value(state.c_str());
}

std::string Impl::XLWorkbook::getSheetName(const string& sheetRID) const
{
    return XmlDocument().document_element().child("sheets").find_child_by_attribute("r:id", sheetRID.c_str()).attribute("name").value();
}

/**
 * @details
 * @todo To be implemented.
 */
void Impl::XLWorkbook::AddChartsheet(const std::string& sheetName, unsigned int index) {}

/**
 * @details
 * @todo To be implemented.
 */
void Impl::XLWorkbook::MoveSheet(const std::string& sheetName, unsigned int newIndex)
{
    // ===== Check that the input is valid
    if (newIndex < 1 || newIndex > std::distance(XmlDocument().document_element().child("sheets").children().begin(),
                                                 XmlDocument().document_element().child("sheets").children().end()))
        throw XLException("Invalid sheet index");

    // ===== If the new index is equal to the current, don't do anything
    if (newIndex == std::distance(XmlDocument().document_element().child("sheets").children().begin(),
                                  std::find_if(XmlDocument().document_element().child("sheets").children().begin(),
                                               XmlDocument().document_element().child("sheets").children().end(),
                                               [&](const XMLNode& item) { return sheetName == item.attribute("name").value(); })))
        return;

    // ===== Modify the node in the XML file
    if (newIndex <= 1)
        getSheetsNode().prepend_move(getSheetsNode().find_child_by_attribute("name", sheetName.c_str()));
    else if (newIndex >= SheetCount())
        getSheetsNode().append_move(getSheetsNode().find_child_by_attribute("name", sheetName.c_str()));
    else {
        auto currentSheet = getSheetsNode().first_child();
        for (auto i = 1; i < newIndex; ++i) currentSheet = currentSheet.next_sibling();
        getSheetsNode().insert_move_before(getSheetsNode().find_child_by_attribute("name", sheetName.c_str()), currentSheet);
    }

    // ===== Updated defined names with worksheet scopes.
    for (auto& definedName : XmlDocument().document_element().child("definedNames").children()) {
        definedName.attribute("localSheetId").set_value(IndexOfSheet(sheetName) - 1);
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
unsigned int Impl::XLWorkbook::IndexOfSheet(const std::string& sheetName) const
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

Impl::XLSheetType Impl::XLWorkbook::TypeOfSheet(const std::string& sheetName) const
{
    //    auto sheetData = find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& data) {
    //        return sheetName == data.sheetNode.attribute("name").value();
    //    });
    //
    //    if (sheetData == m_sheets.end()) throw XLException("Sheet \"" + sheetName + "\" does not exist");
    //
    //    return sheetData->sheetType;
}

Impl::XLSheetType Impl::XLWorkbook::TypeOfSheet(unsigned int index) const
{
    string name = vector<XMLNode>(getSheetsNode().begin(), getSheetsNode().end())[index - 1].attribute("name").as_string();
    return TypeOfSheet(name);
}

/**
 * @details
 */
unsigned int Impl::XLWorkbook::SheetCount() const
{
    return std::distance(getSheetsNode().children().begin(), getSheetsNode().children().end());
    // return m_sheets.size();
}

/**
 * @details
 */
unsigned int Impl::XLWorkbook::WorksheetCount() const
{
    //    return count_if(m_sheets.begin(), m_sheets.end(), [](const XLSheetData& iter) {
    //        return iter.sheetType == XLSheetType::WorkSheet;
    //    });
}

/**
 * @details
 */
unsigned int Impl::XLWorkbook::ChartsheetCount() const
{
    //    return count_if(m_sheets.begin(), m_sheets.end(), [](const XLSheetData& iter) {
    //        return iter.sheetType == XLSheetType::ChartSheet;
    //    });
}

/**
 * @details
 */
std::vector<std::string> Impl::XLWorkbook::SheetNames() const
{
    //    vector<string> result;
    //
    //    for (const auto& item : m_sheets) result.emplace_back(item.sheetNode.attribute("name").value());
    //
    //    return result;
}

/**
 * @details
 */
std::vector<std::string> Impl::XLWorkbook::WorksheetNames() const
{
    std::vector<std::string> results;

    for (const auto& item : getSheetsNode().children()) {
        if (ParentDoc().executeQuery(XLQuery(XLQueryType::GetSheetType, XLQueryParams { { "sheetID", item.attribute("r:id").value() } })) ==
            "WORKSHEET")
            results.emplace_back(item.attribute("name").value());
    }

    return results;
}

/**
 * @details
 */
std::vector<std::string> Impl::XLWorkbook::ChartsheetNames() const
{
    //    vector<string> result;
    //
    //    for (const auto& item : m_sheets)
    //        if (item.sheetType == XLSheetType::ChartSheet)
    //        result.emplace_back(item.sheetNode.attribute("name").value());
    //
    //    return result;
}

/**
 * @details
 */
bool Impl::XLWorkbook::SheetExists(const std::string& sheetName) const
{
    //    return find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& item) {
    //               return sheetName == item.sheetNode.attribute("name").value();
    //           }) != m_sheets.end();
}

/**
 * @details
 */
bool Impl::XLWorkbook::WorksheetExists(const std::string& sheetName) const
{
    //    return find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& item) {
    //               return (sheetName == item.sheetNode.attribute("name").value() &&
    //                       item.sheetType == XLSheetType::WorkSheet);
    //           }) != m_sheets.end();
}

/**
 * @details
 */
bool Impl::XLWorkbook::ChartsheetExists(const std::string& sheetName) const
{
    //    return find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& item) {
    //               return (sheetName == item.sheetNode.attribute("name").value() &&
    //                       item.sheetType == XLSheetType::ChartSheet);
    //           }) != m_sheets.end();
}

/**
 * @details The UpdateSheetName member function searches theroug the usages of the old name and replaces with the
 * new sheet name.
 * @todo Currently, this function only searches through defined names. Consider using this function to update the
 * actual sheet name as well.
 */
void Impl::XLWorkbook::UpdateSheetName(const std::string& oldName, const std::string& newName)
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

/**
 * @details
 */
Impl::XLRelationships* Impl::XLWorkbook::Relationships()
{
    return &m_relationships;
}

/**
 * @details
 */
const Impl::XLRelationships* Impl::XLWorkbook::Relationships() const
{
    return &m_relationships;
}

/**
 * @details
 */
void Impl::XLWorkbook::CreateWorksheet(const XLRelationshipItem& item, const std::string& xmlData)
{
    //        if (getSheetsNode().find_child_by_attribute("r:id", item.Id().c_str()) == nullptr)
    //            throw XLException("Invalid sheet ID");
    //
    //        // Find the appropriate sheet node in the Workbook .xml file; get the name and id of the worksheet.
    //        // If xmlData is empty, set the m_sheets and m_childXmlDocuments elements to nullptr. The worksheet will then be
    //        // lazy-instantiated when the worksheet is requested using the 'Worksheet" function.
    //        // If xmlData is provided (i.e. a new sheet is created or an existing is cloned), create the new worksheets
    //        // accordingly.
    //
    //        unsigned int index = 0;
    //        for (const auto& elem : getSheetsNode().children()) {
    //            if (string(item.Id()) == elem.attribute("r:id").value()) break;
    //            ++index;
    //        }
    //
    //        if (index > m_sheets.size()) index = m_sheets.size();
    //
    //        auto& sheet             = *m_sheets.insert(m_sheets.begin() + index, XLSheetData());
    //        sheet.sheetNode         = getSheetsNode().find_child_by_attribute("r:id", item.Id().c_str());
    //        sheet.sheetRelationship = item;
    //        sheet.sheetContentItem  = ParentDoc().ContentItem(string("/xl/") + item.Target());
    //        sheet.sheetType         = XLSheetType::WorkSheet;
    //        sheet.sheetItem =
    //            (xmlData.empty() ? nullptr
    //                             : make_unique<XLWorksheet>(&ParentDoc().getXmlData(sheet.sheetRelationship.Target()),
    //                                                        sheet.sheetNode.attribute("r:id").value()));
    //
    //        // ===== Update the activeTab attribute. If the deleted sheet was the active sheet, set the property to zero
    //        unsigned int idx = 0;
    //        for (auto& it : getSheetsNode().children()) {
    //            if (m_activeSheet == it) {
    //                XmlDocument()->first_child().child("bookViews").first_child().attribute("activeTab").set_value(idx);
    //                break;
    //            }
    //            index++;
    //        }
}

/**
 * @details
 */
void Impl::XLWorkbook::CreateChartsheet(const XLRelationshipItem& item)
{
    // TODO: Create Chartsheet object here.
}

void Impl::XLWorkbook::WriteXMLData()
{
    //    XLAbstractXMLFile::WriteXMLData();
    //    m_relationships.WriteXMLData();
    //
    //    if (m_sharedStrings) m_sharedStrings.WriteXMLData();
    //
    //    for (auto& sheet : m_sheets)
    //        if (sheet.sheetItem) sheet.sheetItem->WriteXMLData();
}

void Impl::XLWorkbook::executeCommand(Impl::XLCommand command)
{
    switch (command.commandType()) {
        case XLCommandType::SetSheetName:
            setSheetName(command.sender(), command.parameters().at("sheetName"));
            break;

        default:
            break;
    }
}

std::string Impl::XLWorkbook::queryCommand(Impl::XLQuery query) const
{
    switch (query.queryType()) {
        case XLQueryType::GetSheetName:
            return getSheetName(query.parameters().at("sheetID"));

        case XLQueryType::GetSheetVisibility:
            return getSheetNodeByRID(XmlDocument().document_element().child("sheets"), query.parameters().at("sheetID"))
                .attribute("state")
                .value();

        default:
            return std::string();
    }
}

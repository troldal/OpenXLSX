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
XLWorkbook::XLWorkbook(XLXmlData* xmlData)

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
XLWorkbook::~XLWorkbook() = default;

/**
 * @details
 */
bool XLWorkbook::ParseXMLData()
{
    return true;
}

/**
 * @details
 */
XLSheet XLWorkbook::Sheet(const std::string& sheetName)
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
XLSheet XLWorkbook::Sheet(unsigned int index)
{
    return Sheet(vector<XMLNode>(getSheetsNode().begin(), getSheetsNode().end())[index - 1].attribute("name").as_string());
}

/**
 * @details
 */
XLWorksheet XLWorkbook::Worksheet(const std::string& sheetName)
{
    return Sheet(sheetName).Get<XLWorksheet>();
}

/**
 * @details
 */
XLChartsheet XLWorkbook::Chartsheet(const std::string& sheetName)
{
    return Sheet(sheetName).Get<XLChartsheet>();
}

/**
 * @details
 */
bool XLWorkbook::HasSharedStrings() const
{
    return true;    // m_sharedStrings;    // implicit conversion to bool
}

/**
 * @details
 */
XLSharedStrings* XLWorkbook::SharedStrings() const
{
    return &m_sharedStrings;
}

/**
 * @details
 */
void XLWorkbook::DeleteNamedRanges()
{
    for (auto& child : XmlDocument().document_element().child("definedNames").children()) child.parent().remove_child(child);
}

/**
 * @details
 */
void XLWorkbook::DeleteSheet(const std::string& sheetName)
{
    auto sheetID = getSheetsNode().find_child_by_attribute("name", sheetName.c_str()).attribute("r:id").value();

    auto worksheetCount = count_if(getSheetsNode().children().begin(), getSheetsNode().children().end(), [&](const XMLNode& item) {
        return ParentDoc().executeQuery(XLQuerySheetType(item.attribute("r:id").value())).sheetType() == XLContentType::Worksheet;
    });

    auto worksheetType = ParentDoc().executeQuery(XLQuerySheetType(getRID())).sheetType();

    // ===== If this is the last worksheet in the workbook, throw an exception.
    if (worksheetCount == 1 && worksheetType == XLContentType::Worksheet)
        throw XLException("Invalid operation. There must be at least one worksheet in the workbook.");

    ParentDoc().executeCommand(XLCommandDeleteSheet(sheetID, sheetName));

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
void XLWorkbook::AddWorksheet(const std::string& sheetName, unsigned int index)
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
void XLWorkbook::CloneWorksheet(const std::string& extName, const std::string& newName, unsigned int index)
{
    // ===== If a sheet with the given name already exists, throw an exception.
    if (XmlDocument().document_element().child("sheets").find_child_by_attribute("name", newName.c_str()))
        throw XLException("Sheet named \"" + newName + "\" already exists.");

    CreateWorksheet(*InitiateWorksheet(newName, index), Worksheet(extName).GetXmlData());
}

/**
 * @details
 */
XLRelationshipItem* XLWorkbook::InitiateWorksheet(const std::string& sheetName, unsigned int index)
{
    auto        sheetID       = GetNewSheetID(XmlDocument().document_element().child("sheets"));
    std::string worksheetPath = "/xl/worksheets/sheet" + to_string(sheetID) + ".xml";

    ParentDoc().executeCommand(XLCommandAddWorksheet(sheetName, worksheetPath, index));

    auto node  = XMLNode();
    auto nodes = vector<XMLNode>(getSheetsNode().begin(), getSheetsNode().end());
    if (index == 0 || index > nodes.size())
        node = getSheetsNode().append_child("sheet");
    else
        node = getSheetsNode().insert_child_before("sheet", nodes[index - 1]);

    node.append_attribute("name")    = sheetName.c_str();
    node.append_attribute("sheetId") = to_string(sheetID).c_str();
    node.append_attribute("r:id")    = ParentDoc().executeQuery(XLQuerySheetRelsID(worksheetPath)).sheetID().c_str();

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

XMLNode XLWorkbook::getSheetsNode() const
{
    return XmlDocument().first_child().child("sheets");
}

void XLWorkbook::setSheetName(const string& sheetRID, const string& newName)
{
    auto sheetName = XmlDocument().document_element().child("sheets").find_child_by_attribute("r:id", sheetRID.c_str()).attribute("name");

    UpdateSheetName(sheetName.value(), newName);
    sheetName.set_value(newName.c_str());
}

void XLWorkbook::setSheetVisibility(const string& sheetRID, const string& state)
{
    auto stateAttribute =
        XmlDocument().document_element().child("sheets").find_child_by_attribute("r:id", sheetRID.c_str()).attribute("state");

    if (!stateAttribute) {
        stateAttribute =
            XmlDocument().document_element().child("sheets").find_child_by_attribute("r:id", sheetRID.c_str()).append_attribute("state");
    }

    stateAttribute.set_value(state.c_str());
}

std::string XLWorkbook::getSheetName(const string& sheetRID) const
{
    return XmlDocument().document_element().child("sheets").find_child_by_attribute("r:id", sheetRID.c_str()).attribute("name").value();
}

/**
 * @details
 * @todo To be implemented.
 */
void XLWorkbook::AddChartsheet(const std::string& sheetName, unsigned int index) {}

/**
 * @details
 * @todo To be implemented.
 */
void XLWorkbook::MoveSheet(const std::string& sheetName, unsigned int newIndex)
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
unsigned int XLWorkbook::IndexOfSheet(const std::string& sheetName) const
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

XLSheetType XLWorkbook::TypeOfSheet(const std::string& sheetName) const
{
    //    auto sheetData = find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& data) {
    //        return sheetName == data.sheetNode.attribute("name").value();
    //    });
    //
    //    if (sheetData == m_sheets.end()) throw XLException("Sheet \"" + sheetName + "\" does not exist");
    //
    //    return sheetData->sheetType;
}

XLSheetType XLWorkbook::TypeOfSheet(unsigned int index) const
{
    string name = vector<XMLNode>(getSheetsNode().begin(), getSheetsNode().end())[index - 1].attribute("name").as_string();
    return TypeOfSheet(name);
}

/**
 * @details
 */
unsigned int XLWorkbook::SheetCount() const
{
    return std::distance(getSheetsNode().children().begin(), getSheetsNode().children().end());
}

/**
 * @details
 */
unsigned int XLWorkbook::WorksheetCount() const
{
    //    return count_if(m_sheets.begin(), m_sheets.end(), [](const XLSheetData& iter) {
    //        return iter.sheetType == XLSheetType::WorkSheet;
    //    });
}

/**
 * @details
 */
unsigned int XLWorkbook::ChartsheetCount() const
{
    //    return count_if(m_sheets.begin(), m_sheets.end(), [](const XLSheetData& iter) {
    //        return iter.sheetType == XLSheetType::ChartSheet;
    //    });
}

/**
 * @details
 */
std::vector<std::string> XLWorkbook::SheetNames() const
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
std::vector<std::string> XLWorkbook::WorksheetNames() const
{
    std::vector<std::string> results;

    for (const auto& item : getSheetsNode().children()) {
        if (ParentDoc().executeQuery(XLQuerySheetType(item.attribute("r:id").value())).sheetType() == XLContentType::Worksheet)
            results.emplace_back(item.attribute("name").value());
    }

    return results;
}

/**
 * @details
 */
std::vector<std::string> XLWorkbook::ChartsheetNames() const
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
bool XLWorkbook::SheetExists(const std::string& sheetName) const
{
    //    return find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& item) {
    //               return sheetName == item.sheetNode.attribute("name").value();
    //           }) != m_sheets.end();
}

/**
 * @details
 */
bool XLWorkbook::WorksheetExists(const std::string& sheetName) const
{
    //    return find_if(m_sheets.begin(), m_sheets.end(), [&](const XLSheetData& item) {
    //               return (sheetName == item.sheetNode.attribute("name").value() &&
    //                       item.sheetType == XLSheetType::WorkSheet);
    //           }) != m_sheets.end();
}

/**
 * @details
 */
bool XLWorkbook::ChartsheetExists(const std::string& sheetName) const
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
void XLWorkbook::UpdateSheetName(const std::string& oldName, const std::string& newName)
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
XLRelationships* XLWorkbook::Relationships()
{
    return &m_relationships;
}

/**
 * @details
 */
const XLRelationships* XLWorkbook::Relationships() const
{
    return &m_relationships;
}

/**
 * @details
 */
void XLWorkbook::CreateWorksheet(const XLRelationshipItem& item, const std::string& xmlData)
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
void XLWorkbook::CreateChartsheet(const XLRelationshipItem& item)
{
    // TODO: Create Chartsheet object here.
}

void XLWorkbook::executeCommand(XLCommand command)
{
    std::visit(overloaded { [this](XLCommandSetSheetName cmd) { setSheetName(cmd.sheetID(), cmd.sheetName()); },
                            [this](XLCommandSetSheetVisibility cmd) {},
                            [this](XLCommandSetSheetColor cmd) {},
                            [this](XLCommandAddWorksheet cmd) {},
                            [this](XLCommandAddChartsheet cmd) {},
                            [this](XLCommandDeleteSheet cmd) {},
                            [this](XLCommandCloneSheet cmd) {} },
               command);
}

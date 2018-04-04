//
// Created by Troldal on 02/09/16.
//

#include "XLWorkbook.h"
#include "XLWorksheet.h"
#include "XLDocument.h"

using namespace std;
using namespace OpenXLSX;


/**
 * @details The constructor initializes the member variables and calls the loadXMLData from the
 * XLAbstractXMLFile base class.
 */
XLWorkbook::XLWorkbook(XLDocument &parent,
                       const std::string &filePath)

    : XLAbstractXMLFile(parent.RootDirectory()->string(), filePath),
      XLSpreadsheetElement(parent),
      m_sheetsNode(nullptr),
      m_sheetNodes(),
      m_sheets(),
      m_sheetPaths(),
      m_sheetId(0),
      m_sheetCount(0),
      m_worksheetCount(0),
      m_chartsheetCount(0),
      m_definedNames(nullptr),
      m_relationships(nullptr),
      m_sharedStrings(nullptr),
      m_styles(nullptr)
{
    SetModified();
    LoadXMLData(); // This calls the ParseXMLData method!
}

/**
 * @details
 */
bool XLWorkbook::ParseXMLData()
{

    // Set up the Workbook Relationships.
    m_relationships.reset(new XLRelationships(*ParentDocument(), "xl/_rels/workbook.xml.rels"));

    // Find the "sheets" section in the Workbook.xml file
    auto node = XmlDocument()->firstNode();
    while (node) {
        if (node->name() == "sheets") m_sheetsNode = node;
        if (node->name() == "definedNames") m_definedNames = node;
        node = node->nextSibling();
    }

    auto sheetNode = m_sheetsNode->childNode();
    while (sheetNode) {
        m_sheetNodes[sheetNode->attribute("name")->value()] = sheetNode;

        sheetNode = sheetNode->nextSibling();
    }

    for (auto const &item : m_relationships->Relationships()) {
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
    // Calling the const function from the non-const overload.
    //return const_cast<XLWorksheet &>(static_cast<const XLWorkbook *>(this)->Worksheet(sheetName));

    if (m_sheets.find(sheetName) == m_sheets.end()) throw std::range_error("Sheet does not exist");

    if (m_sheets.at(sheetName) == nullptr) {
        m_sheets.at(sheetName) = make_unique<XLWorksheet>(*this, sheetName, m_sheetPaths.at(sheetName));
        m_childXmlDocuments.at(m_sheetPaths.at(sheetName)) = m_sheets.at(sheetName).get();
    }

    XLAbstractSheet *sheet = m_sheets.at(sheetName).get();
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

    // if (m_sheets.at(sheetName) == nullptr)
    //     m_sheets.at(sheetName) = make_unique<XLWorksheet>(*this, sheetName, m_sheetPaths.at(sheetName));

    XLAbstractSheet *sheet = m_sheets.at(sheetName).get();
    if (sheet->Type() == XLSheetType::WorkSheet)
        return dynamic_cast<XLWorksheet *>(sheet);
    else
        throw std::range_error("Sheet does not exist");
}

/**
 * @details
 */
XLAbstractSheet *XLWorkbook::Sheet(unsigned int index)
{
    return nullptr;
}

/**
 * @details
 */
XLAbstractSheet *XLWorkbook::Sheet(const std::string &sheetName)
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
    m_definedNames->deleteChildNodes();
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
    m_sheetNodes.at(sheetName)->deleteNode();
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

    // Standard XML header for Worksheet files.
    string content = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
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

    auto item = CreateWorksheetFile(sheetName, index, content);
    CreateWorksheet(*item);

}

/**
 * @details
 * @todo If the original sheet's tebSelected attribute is set, ensure it is un-set in the clone.
 */
void XLWorkbook::CloneWorksheet(const std::string &extName,
                                const std::string &newName,
                                unsigned int index)
{

    string content = Worksheet(extName)->GetXmlData();
    auto item = CreateWorksheetFile(newName, index, content);
    CreateWorksheet(*item);
}

/**
 * @details
 */
XLRelationshipItem *XLWorkbook::CreateWorksheetFile(const std::string &sheetName,
                                                    unsigned int index,
                                                    const std::string &content)
{

    // Create file
    ParentDocument()->AddFile("xl/worksheets/sheet" + to_string(m_sheetId + 1) + ".xml", content);

    // Add content item to document
    ParentDocument()->AddContentItem("/xl/worksheets/sheet" + to_string(m_sheetId + 1) + ".xml",
                                    XLContentType::Worksheet);

    // Add relationship item
    XLRelationshipItem &item = m_relationships->AddRelationship(XLRelationshipType::Worksheet,
                                                                "worksheets/sheet" + to_string(m_sheetId + 1) + ".xml");

    // Add a Sheet node to the Workbook.xml file
    auto node = m_xmlDocument->createNode("sheet");
    auto name = m_xmlDocument->createAttribute("name", sheetName);
    auto sheetId = m_xmlDocument->createAttribute("sheetId", to_string(m_sheetId + 1));
    auto rId = m_xmlDocument->createAttribute("r:id", item.Id());

    node->appendAttribute(name);
    node->appendAttribute(sheetId);
    node->appendAttribute(rId);
    m_sheetNodes[sheetName] = node;

    // insert Sheet node at the given Index
    if (index == 0)
        m_sheetsNode->appendNode(node);
    else {
        auto curNode = m_sheetsNode->childNode();
        if (!curNode)
            m_sheetsNode->appendNode(node);
        else {
            unsigned int idx = 1;
            while (idx < index) {
                curNode = curNode->nextSibling();
                ++idx;
                if (!curNode) m_sheetsNode->appendNode(node);
            }
            m_sheetsNode->insertNode(curNode, node);
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

    auto node = m_sheetsNode->childNode();
    if (!node) throw runtime_error("Sheet does not exist.");

    unsigned int index = 1;
    while (node) {
        if (node->attribute("name")->value() == sheetName) break;
        node = node->nextSibling();
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
XMLNode *XLWorkbook::SheetNode(const string &sheetName)
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
void XLWorkbook::CreateWorksheet(const XLRelationshipItem &item)
{

    string name;
    auto node = m_sheetsNode->childNode();
    while (node) {
        if (node->attribute("r:id")->value() == item.Id()) {
            name = node->attribute("name")->value();
            break;
        }
        node = node->nextSibling();
    }

    m_sheets[name] = nullptr; //make_unique<XLWorksheet>(*this, name, "xl/" + item.Target());
    m_sheetPaths[name] = "xl/" + item.Target();

    //m_childXmlDocuments["xl/" + item.Target()] = m_sheets.at(name).get();
    m_childXmlDocuments["xl/" + item.Target()] = nullptr;

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
    auto node = m_sheetsNode->childNode();
    while (node) {
        if (node->attribute("r:id")->value() == item.Id()) {
            name = node->attribute("name")->value();
            break;
        }
        node = node->nextSibling();
    }

    //TODO: Create Chartsheet object here.


    m_sheetCount++;
    m_chartsheetCount++;
    m_sheetId++;
}

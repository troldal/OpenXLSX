//
// Created by Troldal on 15/08/16.
//

#include <cstring>
#include <iterator>
#include <utility>
//#include <pugixml.hpp>

#include "XLAppProperties_Impl.hpp"
#include "XLDocument_Impl.hpp"

using namespace std;
using namespace OpenXLSX;

namespace
{
    inline XMLAttribute headingPairsSize(XMLNode docNode)
    {
        return docNode.child("HeadingPairs").first_child().attribute("size");
    }

    inline XMLNode headingPairsCategories(XMLNode docNode)
    {
        return docNode.child("HeadingPairs").first_child().first_child();
    }

    inline XMLNode headingPairsCounts(XMLNode docNode)
    {
        return headingPairsCategories(docNode).next_sibling();
    }

    inline XMLNode sheetNames(XMLNode docNode)
    {
        return docNode.child("TitleOfParts").first_child();
    }

    inline XMLAttribute sheetCount(XMLNode docNode)
    {
        return sheetNames(docNode).attribute("size");
    }
}    // namespace

/**
 * @details
 */
Impl::XLAppProperties::XLAppProperties(XLXmlData* xmlData) : XLAbstractXMLFile(xmlData)
{
    ParseXMLData();
}

/**
 * @details
 */
bool Impl::XLAppProperties::ParseXMLData()
{
    return true;
}

/**
 * @details
 */
XMLNode Impl::XLAppProperties::AddSheetName(const string& title)
{
    auto theNode = sheetNames(XmlDocument().document_element()).append_child("vt:lpstr");
    theNode.text().set(title.c_str());
    sheetCount(XmlDocument().document_element()).set_value(sheetCount(XmlDocument().document_element()).as_uint() + 1);

    return theNode;
}

/**
 * @details
 */
void Impl::XLAppProperties::DeleteSheetName(const string& title)
{
    for (auto& iter : sheetNames(XmlDocument().document_element()).children()) {
        if (iter.child_value() == title) {
            sheetNames(XmlDocument().document_element()).remove_child(iter);
            sheetCount(XmlDocument().document_element()).set_value(sheetCount(XmlDocument().document_element()).as_uint() - 1);
            return;
        }
    }
}

/**
 * @details
 */
void Impl::XLAppProperties::SetSheetName(const string& oldTitle, const string& newTitle)
{
    for (auto& iter : sheetNames(XmlDocument().document_element()).children()) {
        if (iter.child_value() == oldTitle) {
            iter.text().set(newTitle.c_str());
            return;
        }
    }
}

/**
 * @details
 */
XMLNode Impl::XLAppProperties::SheetNameNode(const string& title)
{
    for (auto& sheet : sheetNames(XmlDocument().document_element()).children()) {
        if (string_view(sheet.child_value()) == title) {
            return sheet;
        }
    }

    return XMLNode();
}

/**
 * @details
 */
void Impl::XLAppProperties::AddHeadingPair(const string& name, int value)
{
    for (auto& item : headingPairsCategories(XmlDocument().document_element()).children()) {
        if (item.child_value() == name) return;
    }

    auto pairCategory = headingPairsCategories(XmlDocument().document_element()).append_child("vt:lpstr");
    pairCategory.set_value(name.c_str());

    auto pairCount = headingPairsCounts(XmlDocument().document_element()).append_child("vt:i4");
    pairCount.set_value(to_string(value).c_str());

    headingPairsSize(XmlDocument().document_element())
        .set_value(std::distance(headingPairsCategories(XmlDocument().document_element()).begin(),
                                 headingPairsCategories(XmlDocument().document_element()).end()));
}

/**
 * @details
 */
void Impl::XLAppProperties::DeleteHeadingPair(const string& name)
{
    auto category = headingPairsCategories(XmlDocument().document_element()).begin();
    auto count    = headingPairsCounts(XmlDocument().document_element()).begin();
    while (category != headingPairsCategories(XmlDocument().document_element()).end() &&
           count != headingPairsCounts(XmlDocument().document_element()).end())
    {
        if (category->child_value() == name) {
            headingPairsCategories(XmlDocument().document_element()).remove_child(*category);
            headingPairsCounts(XmlDocument().document_element()).remove_child(*count);
            break;
        }
        ++category;
        ++count;
    }
}

/**
 * @details
 */
void Impl::XLAppProperties::SetHeadingPair(const string& name, int newValue)
{
    auto category = headingPairsCategories(XmlDocument().document_element()).begin();
    auto count    = headingPairsCounts(XmlDocument().document_element()).begin();
    while (category != headingPairsCategories(XmlDocument().document_element()).end() &&
           count != headingPairsCounts(XmlDocument().document_element()).end())
    {
        if (category->child_value() == name) {
            count->text().set(to_string(newValue).c_str());
            break;
        }
        ++category;
        ++count;
    }
}

/**
 * @details
 */
bool Impl::XLAppProperties::SetProperty(const std::string& name, const std::string& value)
{
    Property(name).text().set(value.c_str());
    return true;
}

/**
 * @details
 */
XMLNode Impl::XLAppProperties::Property(const string& name) const
{
    auto property = XmlDocument().first_child().child(name.c_str());
    if (!property) XmlDocument().first_child().append_child(name.c_str());

    return property;
}

/**
 * @details
 */
void Impl::XLAppProperties::DeleteProperty(const string& name)
{
    XmlDocument().first_child().remove_child(Property(name));
}

/**
 * @details
 */
XMLNode Impl::XLAppProperties::AppendSheetName(const std::string& sheetName)
{
    auto theNode = sheetNames(XmlDocument().document_element()).append_child("vt:lpstr");
    theNode.text().set(sheetName.c_str());
    sheetCount(XmlDocument().document_element()).set_value(sheetCount(XmlDocument().document_element()).as_uint() + 1);

    return theNode;
}

/**
 * @details
 */
XMLNode Impl::XLAppProperties::PrependSheetName(const std::string& sheetName)
{
    auto theNode = sheetNames(XmlDocument().document_element()).prepend_child("vt:lpstr");
    theNode.text().set(sheetName.c_str());
    sheetCount(XmlDocument().document_element()).set_value(sheetCount(XmlDocument().document_element()).as_uint() + 1);

    return theNode;
}

/**
 * @details
 */
XMLNode Impl::XLAppProperties::InsertSheetName(const std::string& sheetName, unsigned int index)
{
    if (index <= 1) return PrependSheetName(sheetName);
    if (index > sheetCount(XmlDocument().document_element()).as_uint()) return AppendSheetName(sheetName);

    auto     curNode = sheetNames(XmlDocument().document_element()).first_child();
    unsigned idx     = 1;
    while (curNode != nullptr) {
        if (idx == index) break;
        curNode = curNode.next_sibling();
        ++idx;
    }

    if (!curNode) return AppendSheetName(sheetName);

    auto theNode = sheetNames(XmlDocument().document_element()).insert_child_before("vt:lpstr", curNode);
    theNode.text().set(sheetName.c_str());

    sheetCount(XmlDocument().document_element()).set_value(sheetCount(XmlDocument().document_element()).as_uint() + 1);

    return theNode;
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode Impl::XLAppProperties::MoveSheetName(const std::string& sheetName, unsigned int index)
{
    return XMLNode();    // TODO: Dummy implementation.
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode Impl::XLAppProperties::AppendWorksheetName(const std::string& sheetName)
{
    return XMLNode();    // TODO: Dummy implementation.
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode Impl::XLAppProperties::PrependWorksheetName(const std::string& sheetName)
{
    return XMLNode();    // TODO: Dummy implementation.
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode Impl::XLAppProperties::InsertWorksheetName(const std::string& sheetName, unsigned int index)
{
    return XMLNode();    // TODO: Dummy implementation.
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode Impl::XLAppProperties::MoveWorksheetName(const std::string& sheetName, unsigned int index)
{
    return XMLNode();    // TODO: Dummy implementation.
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode Impl::XLAppProperties::AppendChartsheetName(const std::string& sheetName)
{
    return XMLNode();    // TODO: Dummy implementation.
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode Impl::XLAppProperties::PrependChartsheetName(const std::string& sheetName)
{
    return XMLNode();    // TODO: Dummy implementation.
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode Impl::XLAppProperties::InsertChartsheetName(const std::string& sheetName, unsigned int index)
{
    return XMLNode();    // TODO: Dummy implementation.
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode Impl::XLAppProperties::MoveChartsheetName(const std::string& sheetName, unsigned int index)
{
    return XMLNode();    // TODO: Dummy implementation.
}

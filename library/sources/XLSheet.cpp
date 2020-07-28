//
// Created by Troldal on 02/09/16.
//

#include "XLSheet.hpp"

#include "XLChartsheet.hpp"
#include "XLContentTypes.hpp"
#include "XLDocument.hpp"
#include "XLRelationships.hpp"
#include "XLWorksheet.hpp"

using namespace std;
using namespace OpenXLSX;

/**
 * @details The constructor begins by constructing an instance of its superclass, XLAbstractXMLFile. The default
 * sheet type is WorkSheet and the default sheet state is Visible.
 * @todo Consider to let the sheet type be determined by the subclasses.
 */
XLSheet::XLSheet(XLXmlData* xmlData) : XLXmlFile(xmlData), m_sheet()
{
    if (xmlData->getXmlType() == XLContentType::Worksheet)
        m_sheet = XLWorksheet(xmlData);
    else
        m_sheet = XLChartsheet(xmlData);
}

XLSheet::~XLSheet() = default;

/**
 * @details This method returns the m_sheetName property.
 */
string XLSheet::name() const
{
    return parentDoc().executeQuery(XLQuerySheetName(getRID())).sheetName();
}

/**
 * @details This method sets the name of the sheet to a new name, in the following way:
 * - Set the m_sheetName property to the new name.
 * - In the sheet node in the workbook.xml file, set the name attribute to the new name.
 * - Set the value of the title node in the app.xml file to the new name
 * - Set the m_isModified property to true.
 */
void XLSheet::setName(const std::string& name)
{
    parentDoc().executeCommand(XLCommandSetSheetName(getRID(), this->name(), name));
}

/**
 * @details This method returns the m_sheetState property.
 */
XLSheetState XLSheet::visibility() const
{
    auto state  = parentDoc().executeQuery(XLQuerySheetVisibility(getRID())).sheetVisibility();
    auto result = XLSheetState::Visible;

    if (state == "visible" || state.empty()) {
        result = XLSheetState::Visible;
    }
    else if (state == "hidden") {
        result = XLSheetState::Hidden;
    }
    else if (state == "veryHidden") {
        result = XLSheetState::VeryHidden;
    }

    return result;
}

/**
 * @details This method sets the visibility state of the sheet to a new value, in the following manner:
 * - Set the m_sheetState to the new value.
 * - If the state is set to Hidden or VeryHidden, create a state attribute in the sheet node in the workbook.xml file
 * (if it doesn't exist already) and set the value to the new state.
 * - If the state is set to Visible, delete the state attribute from the sheet node in the workbook.xml file, if it
 * exists.
 */
void XLSheet::setVisibility(XLSheetState state)
{
    auto stateString = std::string();
    switch (state) {
        case XLSheetState::Visible:
            stateString = "visible";
            break;

        case XLSheetState::Hidden:
            stateString = "hidden";
            break;

        case XLSheetState::VeryHidden:
            stateString = "veryHidden";
            break;
    }

    parentDoc().executeCommand(XLCommandSetSheetVisibility(getRID(), name(), stateString));
}

/**
 * @details
 */
XLColor XLSheet::color()
{
    return XLColor();
}

/**
 * @details
 */
void XLSheet::setColor(const XLColor& color) {}

/**
 * @details
 */
void XLSheet::setSelected(bool selected)
{
    unsigned int value = (selected ? 1 : 0);
    xmlDocument().first_child().child("sheetViews").first_child().attribute("tabSelected").set_value(value);
}

/**
 * @details
 */
XLSheetType XLSheet::type() const
{
    if (std::holds_alternative<XLWorksheet>(m_sheet))
        return XLSheetType::Worksheet;
    else
        return XLSheetType::Chartsheet;
}

XLSheet* XLSheet::clone(const string& newName)
{
    std::visit([&](auto&& arg) { arg.clone(newName); }, m_sheet);
    return nullptr;
}

/**
 * @details
 * @todo This method is currently unimplemented.
 */
unsigned int XLSheet::index() const
{
    return parentDoc().executeQuery(XLQuerySheetIndex(getRID())).sheetIndex();
}

/**
 * @details
 * @todo This method is currently unimplemented.
 */
void XLSheet::setIndex() {}

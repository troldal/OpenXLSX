//
// Created by Troldal on 04/10/16.
//

#include "XLDocument.hpp"
#include "XLSheet.hpp"

using namespace OpenXLSX;

/**
 * @details
 */
XLChartsheet::XLChartsheet(XLXmlData* xmlData) : XLXmlFile(xmlData) {}

/**
 * @details
 */
XLSheetState XLChartsheet::visibility() const
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
 * @details
 */
void XLChartsheet::setVisibility(XLSheetState state)
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
std::string XLChartsheet::name() const
{
    return parentDoc().executeQuery(XLQuerySheetName(getRID())).sheetName();
}

/**
 * @details
 */
void XLChartsheet::setName(const std::string& sheetName)
{
    parentDoc().executeCommand(XLCommandSetSheetName(getRID(), name(), sheetName));
}

/**
 * @details
 */
XLColor XLChartsheet::color() const
{
    return XLColor();
}

/**
 * @details
 */
void XLChartsheet::setColor(const XLColor& color) {}

/**
 * @details
 */
uint16_t XLChartsheet::index() const
{
    return parentDoc().executeQuery(XLQuerySheetIndex(getRID())).sheetIndex();
}

/**
 * @details
 */
void XLChartsheet::setIndex(uint16_t index) {}

/**
 * @details
 */
void XLChartsheet::clone(const std::string& newName)
{
    parentDoc().executeCommand(XLCommandCloneSheet(getRID(), newName));
}

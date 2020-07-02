//
// Created by Troldal on 02/09/16.
//

#include "XLSheet_Impl.hpp"
#include "XLContentTypes_Impl.hpp"
#include "XLRelationships_Impl.hpp"
#include "XLDocument_Impl.hpp"

#include <pugixml.hpp>

using namespace std;
using namespace OpenXLSX;

/**
 * @details The constructor begins by constructing an instance of its superclass, XLAbstractXMLFile. The default
 * sheet type is WorkSheet and the default sheet state is Visible.
 * @todo Consider to let the sheet type be determined by the subclasses.
 */
Impl::XLSheet::XLSheet(XLDocument& parent, const std::string& sheetRID, XMLAttribute name, const std::string& filepath, const std::string& xmlData)
        : XLAbstractXMLFile(parent, "xl/" + filepath, xmlData),
          m_sheetRID(sheetRID) {

}

Impl::XLSheet::~XLSheet() = default;

/**
 * @details This method returns the m_sheetName property.
 */
string Impl::XLSheet::Name() const {

    return ParentDoc().executeQuery(XLQuery(XLQueryType::GetSheetName, m_sheetRID));
}

/**
 * @details This method sets the name of the sheet to a new name, in the following way:
 * - Set the m_sheetName property to the new name.
 * - In the sheet node in the workbook.xml file, set the name attribute to the new name.
 * - Set the value of the title node in the app.xml file to the new name
 * - Set the m_isModified property to true.
 */
void Impl::XLSheet::SetName(const std::string& name) {

    ParentDoc().executeCommand(XLCommand(XLCommandType::SetSheetName, m_sheetRID, name));
}

/**
 * @details This method returns the m_sheetState property.
 */
Impl::XLSheetState Impl::XLSheet::State() const {

    auto state = ParentDoc().executeQuery(XLQuery(XLQueryType::GetSheetVisibility, m_sheetRID));
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
 * - If the state is set to Visible, delete the state attribute from the sheet node in the workbook.xml file, if it exists.
 */
void Impl::XLSheet::SetState(XLSheetState state) {

    auto stateString = std::string();
    switch (state) {
        case XLSheetState::Visible :
            stateString = "visible";
            break;

        case XLSheetState::Hidden :
            stateString = "hidden";
            break;

        case XLSheetState::VeryHidden :
            stateString = "veryHidden";
            break;
    }

    ParentDoc().executeCommand(XLCommand(XLCommandType::SetSheetVisibility, m_sheetRID, stateString));
}

/**
 * @details
 */
Impl::XLColor Impl::XLSheet::Color() {
    return Impl::XLColor();
}

/**
 * @details
 */
void Impl::XLSheet::SetColor(const Impl::XLColor& color) {

}

/**
 * @details
 */
void Impl::XLSheet::SetSelected(bool selected) {
    unsigned int value = (selected ? 1 : 0);
    XmlDocument()->first_child().child("sheetViews").first_child().attribute("tabSelected").set_value(value);
}

/**
 * @details
 * @todo This method is currently unimplemented.
 */
unsigned int Impl::XLSheet::Index() const {

    return stoi(ParentDoc().executeQuery(XLQuery(XLQueryType::GetSheetIndex, m_sheetRID)));
}

/**
 * @details
 * @todo This method is currently unimplemented.
 */
void Impl::XLSheet::SetIndex() {

}


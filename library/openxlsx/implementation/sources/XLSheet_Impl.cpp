//
// Created by Troldal on 02/09/16.
//

#include "XLSheet_Impl.h"
#include "XLContentTypes_Impl.h"
#include "XLRelationships_Impl.h"
#include "XLDocument_Impl.h"

#include <pugixml.hpp>

using namespace std;
using namespace OpenXLSX;

/**
 * @details The constructor begins by constructing an instance of its superclass, XLAbstractXMLFile. The default
 * sheet type is WorkSheet and the default sheet state is Visible.
 * @todo Consider to let the sheet type be determined by the subclasses.
 */
Impl::XLSheet::XLSheet(XLWorkbook& parent, XMLAttribute name, const std::string& filepath, const std::string& xmlData)
        : XLAbstractXMLFile(*parent.Document(), "xl/" + filepath, xmlData),
          m_sheetName(name),
          m_sheetType(parent.TypeOfSheet(name.value())),
          m_sheetState(XLSheetState::Visible),
          m_nodeInWorkbook(parent.SheetNode(name.value())),
          m_nodeInApp(parent.Document()->m_docAppProperties->SheetNameNode(name.value())),
          m_parentWorkbook(parent),
          m_nodeInContentTypes(parent.Document()->ContentItem("/xl/" + filepath)),
          m_nodeInWorkbookRels(parent.Relationships()->RelationshipByTarget(filepath)) {

}

Impl::XLSheet::~XLSheet() = default;

/**
 * @details This method returns the m_sheetName property.
 */
string const Impl::XLSheet::Name() const {

    return m_sheetName.value();
}

/**
 * @details This method sets the name of the sheet to a new name, in the following way:
 * - Set the m_sheetName property to the new name.
 * - In the sheet node in the workbook.xml file, set the name attribute to the new name.
 * - Set the value of the title node in the app.xml file to the new name
 * - Set the m_isModified property to true.
 */
void Impl::XLSheet::SetName(const std::string& name) {

    // ===== Update defined names
    m_parentWorkbook.UpdateSheetName(m_sheetName.value(), name);

    // ===== Change sheet name in main .xml files
    m_sheetName.set_value(name.c_str());
    m_nodeInWorkbook.attribute("name").set_value(name.c_str());
    m_nodeInApp.text().set(name.c_str());

}

/**
 * @details This method returns the m_sheetState property.
 */
const XLSheetState& Impl::XLSheet::State() const {

    return m_sheetState;
}

/**
 * @details This method sets the visibility state of the sheet to a new value, in the following manner:
 * - Set the m_sheetState to the new value.
 * - If the state is set to Hidden or VeryHidden, create a state attribute in the sheet node in the workbook.xml file
 * (if it doesn't exist already) and set the value to the new state.
 * - If the state is set to Visible, delete the state attribute from the sheet node in the workbook.xml file, if it exists.
 */
void Impl::XLSheet::SetState(XLSheetState state) {

    m_sheetState = state;

    switch (m_sheetState) {
        case XLSheetState::Hidden : {
            auto att = m_nodeInWorkbook.attribute("state");
            if (!m_nodeInWorkbook.attribute("state"))
                m_nodeInWorkbook.append_attribute("state") = "hidden";
            else
                att.set_value("hidden");
            break;
        }

        case XLSheetState::VeryHidden : {
            auto att = m_nodeInWorkbook.attribute("state");
            if (!att)
                m_nodeInWorkbook.append_attribute("state") = "veryhidden";
            else
                att.set_value("veryhidden"); // todo: Check that this actually works
            break;
        }

        case XLSheetState::Visible : {
            auto att = m_nodeInWorkbook.attribute("state");
            if (att)
                m_nodeInWorkbook.remove_attribute("state");
        }
    }
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
 * @details This method simply returns the m_sheetType property.
 */
const XLSheetType& Impl::XLSheet::Type() const {

    return m_sheetType;
}

/**
 * @details
 * @todo This method is currently unimplemented.
 */
unsigned int Impl::XLSheet::Index() const {

    return m_parentWorkbook.IndexOfSheet(Name());
}

/**
 * @details
 * @todo This method is currently unimplemented.
 */
void Impl::XLSheet::SetIndex() {

}

Impl::XLWorkbook* Impl::XLSheet::Workbook() {

    return &m_parentWorkbook;
}

const Impl::XLWorkbook* Impl::XLSheet::Workbook() const {

    return &m_parentWorkbook;
}

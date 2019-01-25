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
Impl::XLSheet::XLSheet(XLWorkbook& parent,
                       const std::string& name,
                       const std::string& filepath,
                       const std::string& xmlData)
        : XLAbstractXMLFile(*parent.ParentDocument(),
                            filepath,
                            xmlData),
          XLSpreadsheetElement(*parent.ParentDocument()),
          m_sheetName(name),
          m_sheetType(XLSheetType::WorkSheet),
          m_sheetState(XLSheetState::Visible),
          m_nodeInWorkbook(std::make_unique<XMLNode>(parent.SheetNode(name))),
          m_nodeInApp(std::make_unique<XMLNode>(parent.ParentDocument()->m_docAppProperties->SheetNameNode(name))),
          m_nodeInContentTypes(parent.ParentDocument()->ContentItem("/" + filepath)),
          m_nodeInWorkbookRels(parent.Relationships()->RelationshipByTarget(filepath.substr(3))) {

}

Impl::XLSheet::~XLSheet() {

}

/**
 * @details This method returns the m_sheetName property.
 */
const std::string& Impl::XLSheet::Name() const {

    return m_sheetName;
}

/**
 * @details This method sets the name of the sheet to a new name, in the following way:
 * - Set the m_sheetName property to the new name.
 * - In the sheet node in the workbook.xml file, set the name attribute to the new name.
 * - Set the value of the title node in the app.xml file to the new name
 * - Set the m_isModified property to true.
 */
void Impl::XLSheet::SetName(const std::string& name) {

    //ParentDocument()->AppProperties()->SetSheetName(m_sheetName, name);
    m_sheetName = name;
    m_nodeInWorkbook->attribute("name").set_value(name.c_str());
    m_nodeInApp->text().set(name.c_str());

    ParentWorkbook()->UpdateSheetNames();
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
            auto att = m_nodeInWorkbook->attribute("state");
            if (!m_nodeInWorkbook->attribute("state"))
                m_nodeInWorkbook->append_attribute("state") = "hidden";
            else
                att.set_value("hidden");
            break;
        }

        case XLSheetState::VeryHidden : {
            auto att = m_nodeInWorkbook->attribute("state");
            if (!att)
                m_nodeInWorkbook->append_attribute("state") = "veryhidden";
            else
                att.set_value("veryhidden"); // todo: Check that this actually works
            break;
        }

        case XLSheetState::Visible : {
            auto att = m_nodeInWorkbook->attribute("state");
            if (att)
                m_nodeInWorkbook->remove_attribute("state");
        }
    }
}

/**
 * @details The deleteSheet method deletes the sheet in the following steps:
 * - Delete the relevant node in the workbook.xml file.
 * - Delete the title node in the app.xml file.
 * - Delete the content type element in the ContentTypes.xml file.
 * - Delete the relevant item in the relationships file
 * @todo Find the exact names of the .xml fils that will be modified.
 * @todo Consider if this can be moved to the destructor.
 * @todo Delete the associated xml file for the sheet.
 */
void Impl::XLSheet::Delete() {

    // Delete the node in AppProperties.
    ParentDocument()->AppProperties()->DeleteSheetName(m_sheetName);
    m_nodeInApp = std::make_unique<XMLNode>();

    // Delete the item in content_types.xml
    m_nodeInContentTypes->DeleteItem();
    m_nodeInContentTypes = nullptr;

    // Delete the item in Workbook.xml.rels
    ParentDocument()->Workbook()->Relationships()->DeleteRelationship(m_nodeInWorkbookRels->Id());
    m_nodeInWorkbookRels = nullptr;

    // Delete the underlying XML file.
    DeleteXMLData();
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

    return 0;
}

/**
 * @details
 * @todo This method is currently unimplemented.
 */
void Impl::XLSheet::SetIndex() {

}


//
// Created by Troldal on 02/09/16.
//

#include "XLAbstractSheet.h"
#include "XLContentTypes.h"
#include "XLRelationships.h"
#include "XLDocument.h"

using namespace std;
using namespace OpenXLSX;


/**
 * @details The constructor begins by constructing an instance of its superclass, XLAbstractXMLFile. The default
 * sheet type is WorkSheet and the default sheet state is Visible.
 * @todo Consider to let the sheet type be determined by the subclasses.
 */
XLAbstractSheet::XLAbstractSheet(XLWorkbook &parent,
                                 const std::string &name,
                                 const std::string &filepath)
    : XLAbstractXMLFile(parent.ParentDocument()->RootDirectory()->string(), filepath),
      XLSpreadsheetElement(*parent.ParentDocument()),
      m_sheetName(name),
      m_sheetType(XLSheetType::WorkSheet),
      m_sheetState(XLSheetState::Visible),
      m_nodeInWorkbook(parent.SheetNode(name)),
      m_nodeInApp(&parent.ParentDocument()->m_docAppProperties->SheetNameNode(name)),
      m_nodeInContentTypes(parent.ParentDocument()->ContentItem("/" + filepath)),
      m_nodeInWorkbookRels(&parent.Relationships()->RelationshipByTarget(filepath.substr(3)))
{

}

/**
 * @details This method returns the m_sheetName property.
 */
const std::string &XLAbstractSheet::Name() const
{
    return m_sheetName;
}

/**
 * @details This method sets the name of the sheet to a new name, in the following way:
 * - Set the m_sheetName property to the new name.
 * - In the sheet node in the workbook.xml file, set the name attribute to the new name.
 * - Set the value of the title node in the app.xml file to the new name
 * - Set the m_isModified property to true.
 */
void XLAbstractSheet::SetName(const std::string &name)
{
    m_sheetName = name;
    m_nodeInWorkbook->attribute("name")->setValue(name);
    m_nodeInApp->setValue(name);
    SetModified();
}

/**
 * @details This method returns the m_sheetState property.
 */
const XLSheetState &XLAbstractSheet::State() const
{
    return m_sheetState;
}

/**
 * @details This method sets the visibility state of the sheet to a new value, in the following manner:
 * - Set the m_sheetState to the new value.
 * - If the state is set to Hidden or VeryHidden, create a state attribute in the sheet node in the workbook.xml file
 * (if it doesn't exist already) and set the value to the new state.
 * - If the state is set to Visible, delete the state attribute from the sheet node in the workbook.xml file, if it exists.
 */
void XLAbstractSheet::SetState(XLSheetState state)
{
    m_sheetState = state;

    switch (m_sheetState) {
        case XLSheetState::Hidden : {
            auto att = m_nodeInWorkbook->attribute("state");
            if (!att) {
                att = XmlDocument()->createAttribute("state", "hidden");
                m_nodeInWorkbook->appendAttribute(att);
            }
            else {
                att->setValue("hidden");
            }
            break;
        }

        case XLSheetState::VeryHidden : {
            auto att = m_nodeInWorkbook->attribute("state");
            if (!att) {
                att = XmlDocument()->createAttribute("state", "veryhidden");
                m_nodeInWorkbook->appendAttribute(att);
            }
            else {
                att->setValue("veryhidden"); // todo: Check that this actually works
            }
            break;
        }

        case XLSheetState::Visible : {
            auto att = m_nodeInWorkbook->attribute("state");
            if (att) {
                att->deleteAttribute();
            }
        }
    }
    SetModified();
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
void XLAbstractSheet::Delete()
{

    // Delete the node in AppProperties.
    ParentDocument()->AppProperties()->DeleteSheetName(m_sheetName);
    m_nodeInApp = nullptr;

    // Delete the item in content_types.xml
    m_nodeInContentTypes->DeleteItem();
    m_nodeInContentTypes = nullptr;

    // Delete the item in Workbook.xml.rels
    ParentDocument()->Workbook()->Relationships()->DeleteRelationship(m_nodeInWorkbookRels->Id());
    m_nodeInWorkbookRels = nullptr;

    // Delete the underlying XML file.
    //TODO: Doesn't seem to work; However, Excel can still open the file.
    DeleteFile();
}

/**
 * @details This method simply returns the m_sheetType property.
 */
const XLSheetType &XLAbstractSheet::Type() const
{
    return m_sheetType;
}

/**
 * @details
 * @todo This method is currently unimplemented.
 */
unsigned int XLAbstractSheet::Index() const
{
    return 0;
}

/**
 * @details
 * @todo This method is currently unimplemented.
 */
void XLAbstractSheet::SetIndex()
{

}

/**
 * @details
 */
const std::string &XLAbstractSheet::SheetName()
{
    return m_sheetName;
}

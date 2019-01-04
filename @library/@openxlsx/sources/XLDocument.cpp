//
// Created by Troldal on 24/07/16.
//

#include "XLAbstractXMLFile.h"
#include "XLDocument.h"
#include "XLWorksheet.h"
#include "XLTemplate.h"
#include "XLException.h"


using namespace std;
using namespace libzippp;
using namespace OpenXLSX;

/**
 * @details The default constructor, with no arguments.
 */
XLDocument::XLDocument()
    : m_filePath(""),
      m_documentRelationships(nullptr),
      m_contentTypes(nullptr),
      m_docAppProperties(nullptr),
      m_docCoreProperties(nullptr),
      m_workbook(nullptr),
      m_xmlFiles(),
      m_archive(nullptr),
      m_xmlData()
{
}

/**
 * @details An alternative constructor, taking a std::string with the path to the .xlsx package as an argument.
 */
XLDocument::XLDocument(const std::string &docPath)
    : m_filePath(""),
      m_documentRelationships(nullptr),
      m_contentTypes(nullptr),
      m_docAppProperties(nullptr),
      m_docCoreProperties(nullptr),
      m_workbook(nullptr),
      m_xmlFiles(),
      m_archive(nullptr),
      m_xmlData()
{
    OpenDocument(docPath);
}

/**
 * @details The destructor calls the closeDocument method before the object is destroyed.
 */
XLDocument::~XLDocument()
{
    CloseDocument();
}

/**
 * @details The openDocument method opens the .xlsx package in the following manner:
 * - Check if a document is already open. If yes, close it.
 * - Create a temporary folder for the contents of the .xlsx package
 * - Unzip the contents of the package to the temporary folder.
 * - load the contents into the datastructure for manipulation.
 */
void XLDocument::OpenDocument(const string &fileName)
{
    // Check if a document is already open. If yes, close it.
    if(m_archive && m_archive->isOpen()) CloseDocument();

    m_filePath = fileName;
    m_archive = make_unique<ZipArchive>(m_filePath);
    m_archive->open(ZipArchive::WRITE);

    // Open the Relationships and Content_Types files for the document level.
    m_documentRelationships = make_unique<XLRelationships>(*this, "_rels/.rels");
    m_contentTypes = make_unique<XLContentTypes>(*this, "[Content_Types].xml");

    // Create workbook and property items
    m_docCoreProperties = CreateItem<XLCoreProperties>("docProps/core.xml");
    m_docAppProperties = CreateItem<XLAppProperties>("docProps/app.xml");
    m_workbook = CreateItem<XLWorkbook>("xl/workbook.xml");
}

/**
 * @details Create a new document. This is done by saving the data in XLTemplate.h in binary format.
 */
void XLDocument::CreateDocument(const std::string &fileName)
{
    std::ofstream outfile(fileName, std::ios::binary);
    outfile.write(reinterpret_cast<char const *>(excelTemplate), excelTemplateSize);
    outfile.close();

    OpenDocument(fileName);
}

/**
 * @details The document is closed by deleting the temporary folder structure.
 * @todo Consider deleting all the internal objects as well.
 */
void XLDocument::CloseDocument()
{
    if(m_archive) m_archive->discard();
    m_archive.reset(nullptr);
    m_filePath.clear();
    m_documentRelationships.reset(nullptr);
    m_contentTypes.reset(nullptr);
    m_docAppProperties.reset(nullptr);
    m_docCoreProperties.reset(nullptr);
    m_workbook.reset(nullptr);
    m_xmlFiles.clear();
    m_xmlData.clear();
}

/**
 * @details Save the document with the same name. The existing file will be overwritten.
 */
bool XLDocument::SaveDocument()
{
    return SaveDocumentAs(m_filePath);
}

/**
 * @details Save the document with a new name. If present, the 'calcChain.xml file will be ignored. The reason for this
 * is that changes to the document may invalidate the calcChain.xml file. Deleting will force Excel to re-create the
 * file. This will happen automatically, without the user noticing.
 */
bool XLDocument::SaveDocumentAs(const string &fileName)
{
    // If the filename is different than the name of the current file, copy the current file to new destination,
    // close the current zip file and open the new one.
    if (fileName != m_filePath) {
        m_archive->discard();
        std::ifstream src(m_filePath, std::ios::binary);
        std::ofstream dst(fileName, std::ios::binary);
        dst << src.rdbuf();

        m_filePath = fileName;
        m_archive = make_unique<ZipArchive>(m_filePath);
        m_archive->open(ZipArchive::WRITE);
    }

    // Commit all XML files, i.e. save to the zip file.
    m_documentRelationships->CommitXMLData();
    m_contentTypes->CommitXMLData();
    for (auto file : m_xmlFiles) {
        file.second->CommitXMLData();
    }

    // Close and re-open the zip file, in order to save changes.
    m_archive->close();
    m_archive->open(ZipArchive::WRITE);
    m_xmlData.clear();

    return true;
}

/**
 * @details
 * @todo Currently, this method returns the full path, which is not the intention.
 */
std::string XLDocument::DocumentName() const
{
    return m_filePath;
}

/**
 * @details
 */
std::string XLDocument::DocumentPath() const
{
    return m_filePath;
}

/**
 * @details Get a pointer to the underlying XLWorkbook object.
 */
XLWorkbook *XLDocument::Workbook()
{
    return m_workbook.get();
}

/**
 * @details Get a const pointer to the underlying XLWorkbook object.
 */
const XLWorkbook *XLDocument::Workbook() const
{
    return m_workbook.get();
}

/**
 * @details Get the value for a property.
 */
std::string XLDocument::GetProperty(XLDocumentProperties theProperty) const
{
    switch (theProperty) {
        case XLDocumentProperties::Application :
            return m_docAppProperties->Property("Application").value();

        case XLDocumentProperties::AppVersion :
            return m_docAppProperties->Property("AppVersion").value();

        case XLDocumentProperties::Category :
            return m_docCoreProperties->Property("cp:category").value();

        case XLDocumentProperties::Company :
            return m_docAppProperties->Property("Company").value();

        case XLDocumentProperties::CreationDate :
            return m_docCoreProperties->Property("dcterms:created").value();

        case XLDocumentProperties::Creator :
            return m_docCoreProperties->Property("dc:creator").value();

        case XLDocumentProperties::Description :
            return m_docCoreProperties->Property("dc:description").value();

        case XLDocumentProperties::DocSecurity :
            return m_docAppProperties->Property("DocSecurity").value();

        case XLDocumentProperties::HyperlinkBase :
            return m_docAppProperties->Property("HyperlinkBase").value();

        case XLDocumentProperties::HyperlinksChanged :
            return m_docAppProperties->Property("HyperlinksChanged").value();

        case XLDocumentProperties::Keywords :
            return m_docCoreProperties->Property("cp:keywords").value();

        case XLDocumentProperties::LastModifiedBy :
            return m_docCoreProperties->Property("cp:lastModifiedBy").value();

        case XLDocumentProperties::LastPrinted :
            return m_docCoreProperties->Property("cp:lastPrinted").value();

        case XLDocumentProperties::LinksUpToDate :
            return m_docAppProperties->Property("LinksUpToDate").value();

        case XLDocumentProperties::Manager :
            return m_docAppProperties->Property("Manager").value();

        case XLDocumentProperties::ModificationDate :
            return m_docCoreProperties->Property("dcterms:modified").value();

        case XLDocumentProperties::ScaleCrop :
            return m_docAppProperties->Property("ScaleCrop").value();

        case XLDocumentProperties::SharedDoc :
            return m_docAppProperties->Property("SharedDoc").value();

        case XLDocumentProperties::Subject :
            return m_docCoreProperties->Property("dc:subject").value();

        case XLDocumentProperties::Title :
            return m_docCoreProperties->Property("dc:title").value();

    }

    return ""; // To silence compiler warning (Weffc++).
}

/**
 * @details Set the value for a property.
 */
void XLDocument::SetProperty(XLDocumentProperties theProperty,
                             const string &value)
{
    switch (theProperty) {
        case XLDocumentProperties::Application :
            m_docAppProperties->SetProperty("Application", value);
            break;

        case XLDocumentProperties::AppVersion :
            m_docAppProperties->SetProperty("AppVersion", value);
            break;

        case XLDocumentProperties::Category :
            m_docCoreProperties->SetProperty("cp:category", value);
            break;

        case XLDocumentProperties::Company :
            m_docAppProperties->SetProperty("Company", value);
            break;

        case XLDocumentProperties::CreationDate :
            m_docCoreProperties->SetProperty("dcterms:created", value);
            break;

        case XLDocumentProperties::Creator :
            m_docCoreProperties->SetProperty("dc:creator", value);
            break;

        case XLDocumentProperties::Description :
            m_docCoreProperties->SetProperty("dc:description", value);
            break;

        case XLDocumentProperties::DocSecurity :
            m_docAppProperties->SetProperty("DocSecurity", value);
            break;

        case XLDocumentProperties::HyperlinkBase :
            m_docAppProperties->SetProperty("HyperlinkBase", value);
            break;

        case XLDocumentProperties::HyperlinksChanged :
            m_docAppProperties->SetProperty("HyperlinksChanged", value);
            break;

        case XLDocumentProperties::Keywords :
            m_docCoreProperties->SetProperty("cp:keywords", value);
            break;

        case XLDocumentProperties::LastModifiedBy :
            m_docCoreProperties->SetProperty("cp:lastModifiedBy", value);
            break;

        case XLDocumentProperties::LastPrinted :
            m_docCoreProperties->SetProperty("cp:lastPrinted", value);
            break;

        case XLDocumentProperties::LinksUpToDate :
            m_docAppProperties->SetProperty("LinksUpToDate", value);
            break;

        case XLDocumentProperties::Manager :
            m_docAppProperties->SetProperty("Manager", value);
            break;

        case XLDocumentProperties::ModificationDate :
            m_docCoreProperties->SetProperty("dcterms:modified", value);
            break;

        case XLDocumentProperties::ScaleCrop :
            m_docAppProperties->SetProperty("ScaleCrop", value);
            break;

        case XLDocumentProperties::SharedDoc :
            m_docAppProperties->SetProperty("SharedDoc", value);
            break;

        case XLDocumentProperties::Subject :
            m_docCoreProperties->SetProperty("dc:subject", value);
            break;

        case XLDocumentProperties::Title :
            m_docCoreProperties->SetProperty("dc:title", value);
            break;
    }
}

/**
 * @details Delete a property
 */
void XLDocument::DeleteProperty(const string &propertyName)
{
    m_docAppProperties->DeleteProperty(propertyName);
}

/**
 * @details Get a pointer to the sheet node in the app.xml file.
 */
XMLNode XLDocument::SheetNameNode(const std::string &sheetName)
{
    return m_docAppProperties->SheetNameNode(sheetName);
}

/**
 * @details Get a pointer to the content item in the [Content_Types].xml file.
 */
XLContentItem *XLDocument::ContentItem(const std::string &path)
{
    return m_contentTypes->ContentItem(path);
}

/**
 * @details Ad a new ContentItem and return the resulting object.
 */
XLContentItem *XLDocument::AddContentItem(const std::string &contentPath,
                                           XLContentType contentType)
{

    m_contentTypes->addOverride(contentPath, contentType);
    return m_contentTypes->ContentItem(contentPath);
}

/**
 * @details Add a xml file to the package.
 */
void XLDocument::AddOrReplaceXMLFile(const std::string &path,
                                     const std::string &content)
{
    m_xmlData.push_back(content);
    m_archive->addData(path, m_xmlData.back().data(), m_xmlData.back().size());
}

/**
 * @details
 */
std::string XLDocument::GetXMLFile(const std::string &path)
{
    return m_archive->getEntry(path).readAsText();
}

/**
 * @details
 */
void XLDocument::DeleteXMLFile(const std::string &path)
{
    m_archive->deleteEntry(path);
}

/**
 * @details Get a pointer to the XLAppProperties object
 */
XLAppProperties *XLDocument::AppProperties()
{
    return m_docAppProperties.get();
}

/**
 * @details Get a pointer to the (const) XLAppProperties object
 */
const XLAppProperties *XLDocument::AppProperties() const
{
    return m_docAppProperties.get();
}

/**
 * @details Get a pointer to the XLCoreProperties object
 */
XLCoreProperties *XLDocument::CoreProperties()
{
    return m_docCoreProperties.get();
}

/**
 * @details Get a pointer to the (const) XLCoreProperties object
 */
const XLCoreProperties * XLDocument::CoreProperties() const
{
    return m_docCoreProperties.get();
}



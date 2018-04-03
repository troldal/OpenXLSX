//
// Created by Troldal on 24/07/16.
//

#include "XLDocument.h"
#include "XLWorksheet.h"
#include "XLTemplate.h"


using namespace std;
using namespace boost::filesystem;
using namespace libzippp;
using namespace OpenXLSX;

/**
 * @details The default constructor, with no arguments.
 */
XLDocument::XLDocument()
    : m_filePath(""),
      m_tempPath(""),
      m_documentRelationships(nullptr),
      m_contentTypes(nullptr),
      m_docAppProperties(nullptr),
      m_docCoreProperties(nullptr),
      m_workbook(nullptr),
      m_xmlFiles(),
      m_archive(nullptr)
{
}

/**
 * @details An alternative constructor, taking a std::string with the path to the .xlsx package as an argument.
 */
XLDocument::XLDocument(const std::string &docPath)
    : m_filePath(""),
      m_tempPath(""),
      m_documentRelationships(nullptr),
      m_contentTypes(nullptr),
      m_docAppProperties(nullptr),
      m_docCoreProperties(nullptr),
      m_workbook(nullptr),
      m_xmlFiles(),
      m_archive(make_unique<XLArchive>(docPath))
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
    // If the tempPath Property has a ValueAsString (i.e. a file is currently open), close the current file and delete the temporary folder.
    if (!m_tempPath.empty()) CloseDocument();

    m_filePath = fileName;

    // If the temporary folder exists, append a $ to the prefix.
    string prefix = ".~$";
    while (true) {
        m_tempPath = m_filePath.parent_path() /= prefix + m_filePath.stem().string();
        if (!exists(m_tempPath)) break;
        prefix = prefix + "$";
    }

    // Create the root directory of the package contents
    create_directory(m_tempPath);

    ZipArchive archive(m_filePath.string());
    archive.open(ZipArchive::READ_ONLY);

    // Create all package sub-directories in the root
    for (auto it : archive.getEntries()) {
        path temp = m_tempPath;
        temp /= it.getName();
        temp = temp.parent_path();

        create_directories(temp);
    }

    // Unzip all package contents
    for (auto it : archive.getEntries()) {
        if (it.getName() == "xl/calcChain.xml") continue;
        path temp = m_tempPath;
        temp /= it.getName();

        // Read the zip entry as binary and SaveDocument to file
        std::ofstream file(temp.string(), std::ios::out | std::ios::binary | std::ios::app);
        file.write(reinterpret_cast<const char *>(it.readAsBinary()), it.getSize());
        file.close();
    }

    archive.close();

    // Open the Relationships file for the document level.
    m_documentRelationships = make_unique<XLRelationships>(*this, "_rels/.rels");

    // Open the content types XML file for the document
    m_contentTypes = make_unique<XLContentTypes>(*this, "[Content_Types].xml");

    LoadAllFiles();
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
void XLDocument::CloseDocument() const
{
    remove_all(m_tempPath);
}

/**
 * @details Save the document with the same name. The existing file will be overwritten.
 */
bool XLDocument::SaveDocument()
{
    return SaveDocumentAs(m_filePath.string());
}

/**
 * @details Save the document with a new name. If present, the 'calcChain.xml file will be ignored. The reason for this
 * is that changes to the document may invalidate the calcChain.xml file. Deleting will force Excel to re-create the
 * file. This will happen automatically, without the user noticing.
 */
bool XLDocument::SaveDocumentAs(const string &fileName)
{
    (*m_documentRelationships).SaveXMLData();
    (*m_contentTypes).SaveXMLData();
    for (auto file : m_xmlFiles) {
        file.second->SaveXMLData();
    }

    // Create a new archive
    ZipArchive archive(fileName);
    archive.open(ZipArchive::NEW);

    // Iterate through the files in the temporary folder and SaveDocument to the archive
    recursive_directory_iterator iter{m_tempPath};
    while (iter != recursive_directory_iterator{}) {
        if (!is_directory(iter->path()) && iter->path().filename() != "calcChain.xml") {
            archive.addFile(relative(iter->path(), m_tempPath).string(), iter->path().string());
        }
        ++iter;
    }

    archive.close();

    return true;
}

/**
 * @details
 */
std::string XLDocument::DocumentName() const
{
    return m_filePath.filename().string();
}

/**
 * @details
 */
std::string XLDocument::DocumentPath() const
{
    return m_filePath.string();
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
XMLNode *XLDocument::SheetNameNode(const std::string &sheetName)
{
    return &m_docAppProperties->SheetNameNode(sheetName);
}

/**
 * @details Get a pointer to the content item in the [Content_Types].xml file.
 */
XLContentItem *XLDocument::ContentItem(const std::string &path)
{
    return &m_contentTypes->ContentItem(path);
}

/**
 * @details Ad a new ContentItem and return the resulting object.
 */
XLContentItem *XLDocument::AddContentItem(const std::string &contentPath,
                                           XLContentType contentType)
{

    m_contentTypes->addOverride(contentPath, contentType);
    return &m_contentTypes->ContentItem(contentPath);
}

/**
 * @details Get a path object with the path to the root directory of the temporary folder.
 */
const XLPath *XLDocument::RootDirectory() const
{
    return &m_tempPath;
}

/**
 * @details Loads all the XML files in the package. For the workbook.xml file, child files will be loaded
 * automatically.
 */
void XLDocument::LoadAllFiles()
{
    // Iterate through the document relationship items.
    for (const auto &iter : m_documentRelationships->Relationships()) {
        if (iter.second->Type() == XLRelationshipType::CoreProperties) {
            m_docCoreProperties = make_unique<XLCoreProperties>(*this, iter.second->Target());
            m_xmlFiles[m_docCoreProperties->FilePath()] = m_docCoreProperties.get();
        }
    }

    for (const auto &iter : m_documentRelationships->Relationships()) {
        if (iter.second->Type() == XLRelationshipType::ExtendedProperties) {
            m_docAppProperties = make_unique<XLAppProperties>(*this, iter.second->Target());
            m_xmlFiles[m_docAppProperties->FilePath()] = m_docAppProperties.get();
        }
    }

    for (const auto &iter : m_documentRelationships->Relationships()) {
        if (iter.second->Type() == XLRelationshipType::Workbook) {
            m_workbook = make_unique<XLWorkbook>(*this, iter.second->Target());
            m_xmlFiles[m_workbook->FilePath()] = m_workbook.get();
        }
    }
}

/**
 * @details Add a xml file to the package.
 */
void XLDocument::AddFile(const std::string &path,
                         const std::string &content)
{
    auto fullPath = m_tempPath;
    fullPath /= path;

    std::ofstream file(fullPath.string());
    file << content;
    file.close();
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


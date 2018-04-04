//
// Created by Troldal on 05/08/16.
//

#include <boost/filesystem.hpp>
#include "XLWorkbook.h"
#include "XLWorksheet.h"

using namespace std;
using namespace OpenXLSX;
using namespace boost::filesystem;

/**
 * @details The constructor creates a new object with the parent XLDocument and the file path as input, with
 * an optional input being a std::string with the XML data. I
 * @warning If the XML data is provided by a string, the file path will be the path of the new file,
 * once created; any existing files will be overwritten.
 * @todo Consider if the file needs to be saved immediately if XML data is provided as a string.
 */
XLAbstractXMLFile::XLAbstractXMLFile(const std::string &root,
                                     const std::string &filePath,
                                     const std::string &data)
    : m_xmlDocument(make_unique<XMLDocument>()),
      m_filePath(filePath),
      m_root(root),
      m_childXmlDocuments(),
      m_isModified(false)
{
    if (data != "") SetXmlData(data);
}

/**
 * @details This method sets the XML data with a std::string as input. The underlying XMLDocument reads the data.
 * @todo Consider if the file needs to be saved immediately. Also, call parseXMLData.
 */
void XLAbstractXMLFile::SetXmlData(const std::string &xmlData)
{
    m_xmlDocument->readData(xmlData);
}

/**
 * @details This method retrieves the underlying XML data as a std::string.
 */
std::string XLAbstractXMLFile::GetXmlData() const
{
    return m_xmlDocument->getData();
}

/**
 * @details This method loads the XML data from the file and fills the object data structure by a call to the
 * parseXMLData method. If the file does not exist, false is returned; otherwise, true.
 */
bool XLAbstractXMLFile::LoadXMLData()
{
    if (m_filePath.empty()) return false;

    path thePath = m_root;
    thePath /= m_filePath;

    m_xmlDocument->loadFile(thePath.string());

    return ParseXMLData();
}

/**
 * @details This method saves the XML data to file, but only if the XMLDocument object exists (which may not be the
 * case if it has been explicitly deleted by the user) and is in a modified state.
 * After the file has been saved, all child XML files are saved to disk as well.
 */
void XLAbstractXMLFile::SaveXMLData() const
{
    path thePath = m_root;
    thePath /= m_filePath;

    if (m_isModified && m_xmlDocument)
        (*m_xmlDocument).saveFile(thePath.string());

    m_isModified = false;

    for (auto file : m_childXmlDocuments)
        if (file.second) file.second->SaveXMLData();
}

/**
 * @details This method returns the path of the XML file as a std::string.
 */
const string &XLAbstractXMLFile::FilePath() const
{
    return m_filePath;
}

/**
 * @details This method sets the file path given as a std::string, by modifying the m_filePath property.
 */
void XLAbstractXMLFile::SetFilePath(const string &filePath)
{
    m_filePath = filePath;
}

/**
 * @details This method is mainly meant for debugging, by enabling printing of the xml file to cout.
 */
void XLAbstractXMLFile::Print() const
{
    XmlDocument()->print();
}

/**
 * @details This method returns a pointer to the underlying XMLDocument resource.
 */
XMLDocument * XLAbstractXMLFile::XmlDocument()
{
    return const_cast<XMLDocument *>(static_cast<const XLAbstractXMLFile *>(this)->XmlDocument());
}

/**
 * @details
 */
const XMLDocument *XLAbstractXMLFile::XmlDocument() const
{
    return m_xmlDocument.get();
}

/**
 * @details This method "deletes" the object in the following manner:
 * - Delete the underlying XMLDocument resource.
 * - Delete the XML file with the raw data.
 * The XLAbstractXMLFile object itself is not deleted.
 * @todo Consider if there is a more elegant way to do this.
 */
bool XLAbstractXMLFile::DeleteFile()
{
    m_xmlDocument.reset(nullptr);
    path p(m_filePath);
    remove(p);
    return true;
}

/**
 * @details This method returns a pointer to the parent XLDocument object.
 */ /*
XLDocument &XLAbstractXMLFile::ParentDocument() const
{
    return m_parentDocument;
} */

/**
 * @details Set the XLWorksheet object to 'modified'. This is done by setting the m_is Modified member to true.
 */
void XLAbstractXMLFile::SetModified()
{
    m_isModified = true;
}

/**
 * @details
 */
bool XLAbstractXMLFile::IsModified()
{
    return m_isModified;
}

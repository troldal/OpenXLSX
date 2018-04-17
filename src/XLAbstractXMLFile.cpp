//
// Created by Troldal on 05/08/16.
//

#include "XLDocument.h"
#include "XLWorkbook.h"
#include "XLWorksheet.h"

using namespace std;
using namespace OpenXLSX;

/**
 * @details The constructor creates a new object with the parent XLDocument and the file path as input, with
 * an optional input being a std::string with the XML data. If the XML data is provided by a string, any file with
 * the same path in the .zip file will be overwritten upon saving of the document. If no xmlData is provided,
 * the data will be read from the .zip file, using the given path.
 */
XLAbstractXMLFile::XLAbstractXMLFile(XLDocument &parent,
                                     const std::string &filePath,
                                     const std::string &xmlData)
    :   m_parentDocument(parent),
        m_path(filePath),
        m_xmlDocument(make_unique<XMLDocument>()),
        m_childXmlDocuments(),
        m_isModified(false)
{
    if (xmlData.empty()) SetXmlData(m_parentDocument.GetXMLFile(m_path));
    else SetXmlData(xmlData);
}

/**
 * @details This method sets the XML data with a std::string as input. The underlying XMLDocument reads the data.
 */
void XLAbstractXMLFile::SetXmlData(const std::string &xmlData)
{
    m_xmlDocument->ReadData(xmlData);
}

/**
 * @details This method retrieves the underlying XML data as a std::string.
 */
std::string XLAbstractXMLFile::GetXmlData() const
{
    return m_xmlDocument->GetData();
}

/**
 * @details The CommitXMLData method calls the AddOrReplaceXMLFile method for the current object and all child objects.
 * This, in turn, will add or replace the XML data files in the zipped .xlsx package.
 */
void XLAbstractXMLFile::CommitXMLData()
{
    m_parentDocument.AddOrReplaceXMLFile(m_path, GetXmlData());
    for (auto file : m_childXmlDocuments) {
        if(file.second) file.second->CommitXMLData();
    }
}

/**
 * @details The DeleteXMLData method calls the DeleteXMLFile method for the current object and all child objects.
 * This, in turn, delete the XML data files in the zipped .xlsx package.
 */
void XLAbstractXMLFile::DeleteXMLData()
{
    m_parentDocument.DeleteXMLFile(m_path);
    for (auto file : m_childXmlDocuments) {
        if(file.second) file.second->DeleteXMLData();
    }
}

/**
 * @details This method returns the path in the .zip file of the XML file as a std::string.
 */
const string &XLAbstractXMLFile::FilePath() const
{
    return m_path;
}

/**
 * @details This method is mainly meant for debugging, by enabling printing of the xml file to cout.
 */
void XLAbstractXMLFile::Print() const
{
    XmlDocument()->Print();
}

/**
 * @details This method returns a pointer to the underlying XMLDocument resource.
 */
XMLDocument * XLAbstractXMLFile::XmlDocument()
{
    return const_cast<XMLDocument *>(static_cast<const XLAbstractXMLFile *>(this)->XmlDocument());
}

/**
 * @details This method returns a pointer to the underlying XMLDocument resource as const.
 */
const XMLDocument *XLAbstractXMLFile::XmlDocument() const
{
    return m_xmlDocument.get();
}

/**
 * @details Set the XLWorksheet object to 'modified'. This is done by setting the m_is Modified member to true.
 */
void XLAbstractXMLFile::SetModified()
{
    m_isModified = true;
}

/**
 * @details Returns the value of the m_isModified member variable.
 */
bool XLAbstractXMLFile::IsModified()
{
    return m_isModified;
}

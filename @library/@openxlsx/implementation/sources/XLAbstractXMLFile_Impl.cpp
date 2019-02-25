#include <utility>

//
// Created by Troldal on 05/08/16.
//

#include "XLDocument_Impl.h"
#include <sstream>
#include <pugixml.hpp>

using namespace std;
using namespace OpenXLSX;

/**
 * @details The constructor creates a new object with the parent XLDocument and the file path as input, with
 * an optional input being a std::string with the XML data. If the XML data is provided by a string, any file with
 * the same path in the .zip file will be overwritten upon saving of the document. If no xmlData is provided,
 * the data will be read from the .zip file, using the given path.
 */
Impl::XLAbstractXMLFile::XLAbstractXMLFile(XLDocument& parent, std::string filePath, const std::string& xmlData)
        : m_xmlDocument(std::make_unique<XMLDocument>()),
          m_parentDocument(parent),
          m_path(filePath),
          m_childXmlDocuments() {

    if (xmlData.empty())
        SetXmlData(m_parentDocument.GetXMLFile(m_path));
    else
        SetXmlData(xmlData);

    CommitXMLData();
}

/**
 * @details This method sets the XML data with a std::string as input. The underlying XMLDocument reads the data.
 */
void Impl::XLAbstractXMLFile::SetXmlData(const std::string& xmlData) {

    m_xmlDocument->load_string(xmlData.c_str());
}

/**
 * @details This method retrieves the underlying XML data as a std::string.
 */
std::string Impl::XLAbstractXMLFile::GetXmlData() const {

    ostringstream ostr;
    m_xmlDocument->print(ostr);
    return ostr.str();
}

/**
 * @details The CommitXMLData method calls the AddOrReplaceXMLFile method for the current object and all child objects.
 * This, in turn, will add or replace the XML data files in the zipped .xlsx package.
 */
void Impl::XLAbstractXMLFile::CommitXMLData() {

    m_parentDocument.AddOrReplaceXMLFile(m_path, GetXmlData());
    for (auto file : m_childXmlDocuments) {
        if (file.second != nullptr)
            file.second->CommitXMLData();
    }
}

/**
 * @details The DeleteXMLData method calls the DeleteXMLFile method for the current object and all child objects.
 * This, in turn, delete the XML data files in the zipped .xlsx package.
 */
void Impl::XLAbstractXMLFile::DeleteXMLData() {

    m_parentDocument.DeleteXMLFile(m_path);
    for (auto file : m_childXmlDocuments) {
        if (file.second != nullptr)
            file.second->DeleteXMLData();
    }
}

/**
 * @details This method returns the path in the .zip file of the XML file as a std::string.
 */
const string& Impl::XLAbstractXMLFile::FilePath() const {

    return m_path;
}

/**
 * @details This method is mainly meant for debugging, by enabling printing of the xml file to cout.
 */
void Impl::XLAbstractXMLFile::Print() const {

    XmlDocument()->print(cout);
}

/**
 * @details This method returns a pointer to the underlying XMLDocument resource.
 */
XMLDocument* Impl::XLAbstractXMLFile::XmlDocument() {

    return const_cast<XMLDocument*>(static_cast<const XLAbstractXMLFile*>(this)->XmlDocument());
}

/**
 * @details This method returns a pointer to the underlying XMLDocument resource as const.
 */
const XMLDocument* Impl::XLAbstractXMLFile::XmlDocument() const {

    return m_xmlDocument.get();
}


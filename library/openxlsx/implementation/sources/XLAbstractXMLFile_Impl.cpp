#include <utility>

//
// Created by Troldal on 05/08/16.
//

#include "XLDocument_Impl.h"
#include <sstream>
#include <pugixml.hpp>
#include <XLAbstractXMLFile_Impl.h>

using namespace std;
using namespace OpenXLSX;

/**
 * @details The constructor creates a new object with the parent XLDocument and the file path as input, with
 * an optional input being a std::string with the XML data. If the XML data is provided by a string, any file with
 * the same path in the .zip file will be overwritten upon saving of the document. If no xmlData is provided,
 * the data will be read from the .zip file, using the given path.
 */
Impl::XLAbstractXMLFile::XLAbstractXMLFile(XLDocument& parent, std::string filePath, const std::string& xmlData)
        : m_path(std::move(filePath)),
          m_parentDocument(parent),
          m_xmlDocument(XMLDocument()) {

    if (xmlData.empty())
        SetXmlData(m_parentDocument.GetXMLFile(m_path));
    else
        SetXmlData(xmlData);

    m_parentDocument.AddOrReplaceXMLFile(m_path, GetXmlData());
}

Impl::XLAbstractXMLFile::operator bool() const {

    return !GetXmlData().empty();
}

/**
 * @details This method sets the XML data with a std::string as input. The underlying XMLDocument reads the data.
 */
void Impl::XLAbstractXMLFile::SetXmlData(const std::string& xmlData) {

    m_xmlDocument.load_string(xmlData.c_str(),pugi::parse_default | pugi::parse_ws_pcdata);
}

/**
 * @details This method retrieves the underlying XML data as a std::string.
 */
const string& Impl::XLAbstractXMLFile::GetXmlData() const {

    ostringstream ostr;
    m_xmlDocument.save(ostr, "", pugi::format_raw);
    m_xmlData = ostr.str();
    return m_xmlData;
}

/**
 * @details The CommitXMLData method calls the AddOrReplaceXMLFile method for the current object and all child objects.
 * This, in turn, will add or replace the XML data files in the zipped .xlsx package.
 */
void Impl::XLAbstractXMLFile::WriteXMLData() {

    m_parentDocument.AddOrReplaceXMLFile(m_path, GetXmlData());
}

/**
 * @details The DeleteXMLData method calls the DeleteXMLFile method for the current object and all child objects.
 * This, in turn, delete the XML data files in the zipped .xlsx package.
 */
void Impl::XLAbstractXMLFile::DeleteXMLData() {

    m_parentDocument.DeleteXMLFile(m_path);
}

/**
 * @details This method returns the path in the .zip file of the XML file as a std::string.
 */
const string& Impl::XLAbstractXMLFile::FilePath() const {

    return m_path;
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

    return &m_xmlDocument;
}


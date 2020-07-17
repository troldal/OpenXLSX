//
// Created by Kenneth Balslev on 14/07/2020.
//

#include "XLXmlData.hpp"

#include "XLDocument.hpp"

OpenXLSX::Impl::XLXmlData::XLXmlData(OpenXLSX::Impl::XLDocument*   parentDoc,
                                     const std::string&            xmlPath,
                                     const std::string&            xmlId,
                                     OpenXLSX::Impl::XLContentType xmlType)
    : m_parentDoc(parentDoc),
      m_xmlPath(xmlPath),
      m_xmlID(xmlId),
      m_xmlType(xmlType),
      m_xmlDoc()
{
    m_xmlDoc.reset();
}

void OpenXLSX::Impl::XLXmlData::setRawData(const std::string& data)
{
    m_xmlDoc.load_string(data.c_str(), pugi::parse_default | pugi::parse_ws_pcdata);
}

std::string OpenXLSX::Impl::XLXmlData::getRawData() const
{
    std::ostringstream ostr;
    getXmlDocument()->save(ostr, "", pugi::format_raw);
    return ostr.str();
}

OpenXLSX::Impl::XLDocument* OpenXLSX::Impl::XLXmlData::getParentDoc()
{
    return m_parentDoc;
}

const OpenXLSX::Impl::XLDocument* OpenXLSX::Impl::XLXmlData::getParentDoc() const
{
    return m_parentDoc;
}

std::string OpenXLSX::Impl::XLXmlData::getXmlPath() const
{
    return m_xmlPath;
}

std::string OpenXLSX::Impl::XLXmlData::getXmlID() const
{
    return m_xmlID;
}

OpenXLSX::Impl::XLContentType OpenXLSX::Impl::XLXmlData::getXmlType() const
{
    return m_xmlType;
}

OpenXLSX::XMLDocument* OpenXLSX::Impl::XLXmlData::getXmlDocument()
{
    if (!m_xmlDoc.document_element())
        m_xmlDoc.load_string(m_parentDoc->getXmlDataFromArchive(m_xmlPath).c_str(), pugi::parse_default | pugi::parse_ws_pcdata);

    return &m_xmlDoc;
}

const OpenXLSX::XMLDocument* OpenXLSX::Impl::XLXmlData::getXmlDocument() const
{
    if (!m_xmlDoc.document_element())
        m_xmlDoc.load_string(m_parentDoc->getXmlDataFromArchive(m_xmlPath).c_str(), pugi::parse_default | pugi::parse_ws_pcdata);

    return &m_xmlDoc;
}

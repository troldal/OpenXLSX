#include "XLXmlFile.hpp"

#include "XLDocument.hpp"

#include <sstream>
#include <utility>

using namespace std;
using namespace OpenXLSX;

/**
 * @details The constructor creates a new object with the parent XLDocument and the file path as input, with
 * an optional input being a std::string with the XML data. If the XML data is provided by a string, any file with
 * the same path in the .zip file will be overwritten upon saving of the document. If no xmlData is provided,
 * the data will be read from the .zip file, using the given path.
 */
XLXmlFile::XLXmlFile(XLXmlData* xmlData) : m_xmlData(xmlData) {}

/**
 * @details
 */
XLXmlFile::operator bool() const
{
    return !xmlData().empty();
}

/**
 * @details This method sets the XML data with a std::string as input. The underlying XMLDocument reads the data.
 * When envoking the load_string method in PugiXML, the flag 'parse_ws_pcdata' is passed along with the default flags.
 * This will enable parsing of whitespace characters. If not set, Excel cells with only spaces will be returned as
 * empty strings, which is not what we want. The downside is that whitespace characters such as \\n and \\t in the
 * input xml file may mess up the parsing.
 */
void XLXmlFile::setXmlData(const std::string& xmlData)
{
    m_xmlData->setRawData(xmlData);
}

/**
 * @details This method retrieves the underlying XML data as a std::string.
 */
std::string XLXmlFile::xmlData() const
{
    return m_xmlData->getRawData();
}

/**
 * @details This method returns the path in the .zip file of the XML file as a std::string.
 */
string XLXmlFile::filePath() const
{
    return m_xmlData->getXmlPath();
}

/**
 * @details This method returns a pointer to the underlying XMLDocument resource.
 */
XMLDocument& XLXmlFile::xmlDocument()
{
    return const_cast<XMLDocument&>(static_cast<const XLXmlFile*>(this)->xmlDocument());    // NOLINT
}

/**
 * @details This method returns a pointer to the underlying XMLDocument resource as const.
 */
const XMLDocument& XLXmlFile::xmlDocument() const
{
    return *m_xmlData->getXmlDocument();
}

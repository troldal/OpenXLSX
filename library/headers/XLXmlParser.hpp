//
// Created by Kenneth Balslev on 31/07/2020.
//

#ifndef OPENXLSX_XLXMLPARSER_HPP
#define OPENXLSX_XLXMLPARSER_HPP

namespace pugi
{
    class xml_node;
    class xml_attribute;
    class xml_document;
}    // namespace pugi

namespace OpenXLSX
{
    using XMLNode      = pugi::xml_node;
    using XMLAttribute = pugi::xml_attribute;
    using XMLDocument  = pugi::xml_document;
}    // namespace OpenXLSX
#endif    // OPENXLSX_XLXMLPARSER_HPP

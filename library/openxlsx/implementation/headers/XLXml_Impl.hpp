//
// Created by Troldal on 2018-12-25.
//

#ifndef OPENXLSX_IMPL_XLXML_H
#define OPENXLSX_IMPL_XLXML_H

// ===== Third Party Library Includes ===== //
#include <pugixml.hpp>

namespace OpenXLSX
{
    using XMLNode = pugi::xml_node;
    using XMLAttribute = pugi::xml_attribute;
    using XMLDocument = pugi::xml_document;

    inline XMLNode GetChildByIndex(XMLNode parent, unsigned long index) {

        XMLNode result = parent.first_child();
        for (unsigned long it = 0; it < index; ++it)
            result = result.next_sibling();

        return result;
    }

} // namespace OpenXLSX

#endif //OPENXLSX_IMPL_XLXML_H

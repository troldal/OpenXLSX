/*

   ____                               ____      ___ ____       ____  ____      ___
  6MMMMb                              `MM(      )M' `MM'      6MMMMb\`MM(      )M'
 8P    Y8                              `MM.     d'   MM      6M'    ` `MM.     d'
6M      Mb __ ____     ____  ___  __    `MM.   d'    MM      MM        `MM.   d'
MM      MM `M6MMMMb   6MMMMb `MM 6MMb    `MM. d'     MM      YM.        `MM. d'
MM      MM  MM'  `Mb 6M'  `Mb MMM9 `Mb    `MMd       MM       YMMMMb     `MMd
MM      MM  MM    MM MM    MM MM'   MM     dMM.      MM           `Mb     dMM.
MM      MM  MM    MM MMMMMMMM MM    MM    d'`MM.     MM            MM    d'`MM.
YM      M9  MM    MM MM       MM    MM   d'  `MM.    MM            MM   d'  `MM.
 8b    d8   MM.  ,M9 YM    d9 MM    MM  d'    `MM.   MM    / L    ,M9  d'    `MM.
  YMMMM9    MMYMMM9   YMMMM9 _MM_  _MM_M(_    _)MM_ _MMMMMMM MYMMMM9 _M(_    _)MM_
            MM
            MM
           _MM_

  Copyright (c) 2018, Kenneth Troldal Balslev

  All rights reserved.

  Redistribution and use in source and binary forms, with or without
  modification, are permitted provided that the following conditions are met:
  - Redistributions of source code must retain the above copyright
    notice, this list of conditions and the following disclaimer.
  - Redistributions in binary form must reproduce the above copyright
    notice, this list of conditions and the following disclaimer in the
    documentation and/or other materials provided with the distribution.
  - Neither the name of the author nor the
    names of any contributors may be used to endorse or promote products
    derived from this software without specific prior written permission.

  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
  ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
  WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
  DISCLAIMED. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY
  DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
  (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
  LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
  ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
  (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
  SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

 */

// ===== External Includes ===== //
#include <cstring>
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLContentTypes.hpp"
#include "XLDocument.hpp"

#include <XLException.hpp>

using namespace OpenXLSX;

namespace
{
    const std::string applicationOpenXmlOfficeDocument = "application/vnd.openxmlformats-officedocument";
    const std::string applicationOpenXmlPackage        = "application/vnd.openxmlformats-package";
    const std::string applicationMicrosoftExcel        = "application/vnd.ms-excel";
    const std::string applicationMicrosoftOffice       = "application/vnd.ms-office";

    /**
     * @details
     * @note 2024-08-31: In line with a change to hardcoded XML relationship domains in XLRelationships.cpp, replaced the repeating
     *          hardcoded strings here with the above declared constants, preparing a potential need for a similar "dumb" fallback implementation
     */
    XLContentType GetTypeFromString(const std::string& typeString)
    {
        XLContentType type { XLContentType::Unknown };

        if (typeString == applicationMicrosoftExcel + ".Sheet.macroEnabled.main+xml")
            type = XLContentType::WorkbookMacroEnabled;
        else if (typeString == applicationOpenXmlOfficeDocument + ".spreadsheetml.sheet.main+xml")
            type = XLContentType::Workbook;
        else if (typeString == applicationOpenXmlOfficeDocument + ".spreadsheetml.worksheet+xml")
            type = XLContentType::Worksheet;
        else if (typeString == applicationOpenXmlOfficeDocument + ".spreadsheetml.chartsheet+xml")
            type = XLContentType::Chartsheet;
        else if (typeString == applicationOpenXmlOfficeDocument + ".spreadsheetml.externalLink+xml")
            type = XLContentType::ExternalLink;
        else if (typeString == applicationOpenXmlOfficeDocument + ".theme+xml")
            type = XLContentType::Theme;
        else if (typeString == applicationOpenXmlOfficeDocument + ".spreadsheetml.styles+xml")
            type = XLContentType::Styles;
        else if (typeString == applicationOpenXmlOfficeDocument + ".spreadsheetml.sharedStrings+xml")
            type = XLContentType::SharedStrings;
        else if (typeString == applicationOpenXmlOfficeDocument + ".drawing+xml")
            type = XLContentType::Drawing;
        else if (typeString == applicationOpenXmlOfficeDocument + ".drawingml.chart+xml")
            type = XLContentType::Chart;
        else if (typeString == applicationMicrosoftOffice + ".chartstyle+xml")
            type = XLContentType::ChartStyle;
        else if (typeString == applicationMicrosoftOffice + ".chartcolorstyle+xml")
            type = XLContentType::ChartColorStyle;
        else if (typeString == applicationMicrosoftExcel + ".controlproperties+xml")
            type = XLContentType::ControlProperties;
        else if (typeString == applicationOpenXmlOfficeDocument + ".spreadsheetml.calcChain+xml")
            type = XLContentType::CalculationChain;
        else if (typeString == applicationMicrosoftOffice + ".vbaProject")
            type = XLContentType::VBAProject;
        else if (typeString == applicationOpenXmlPackage + ".core-properties+xml")
            type = XLContentType::CoreProperties;
        else if (typeString == applicationOpenXmlOfficeDocument + ".extended-properties+xml")
            type = XLContentType::ExtendedProperties;
        else if (typeString == applicationOpenXmlOfficeDocument + ".custom-properties+xml")
            type = XLContentType::CustomProperties;
        else if (typeString == applicationOpenXmlOfficeDocument + ".spreadsheetml.comments+xml")
            type = XLContentType::Comments;
        else if (typeString == applicationOpenXmlOfficeDocument + ".spreadsheetml.table+xml")
            type = XLContentType::Table;
        else if (typeString == applicationOpenXmlOfficeDocument + ".vmlDrawing")
            type = XLContentType::VMLDrawing;
        else
            type = XLContentType::Unknown;

        return type;
    }

    /**
     * @details
     */
    std::string GetStringFromType(XLContentType type)
    {
        std::string typeString;

        if (type == XLContentType::WorkbookMacroEnabled)
            typeString = applicationMicrosoftExcel + ".Sheet.macroEnabled.main+xml";
        else if (type == XLContentType::Workbook)
            typeString = applicationOpenXmlOfficeDocument + ".spreadsheetml.sheet.main+xml";
        else if (type == XLContentType::Worksheet)
            typeString = applicationOpenXmlOfficeDocument + ".spreadsheetml.worksheet+xml";
        else if (type == XLContentType::Chartsheet)
            typeString = applicationOpenXmlOfficeDocument + ".spreadsheetml.chartsheet+xml";
        else if (type == XLContentType::ExternalLink)
            typeString = applicationOpenXmlOfficeDocument + ".spreadsheetml.externalLink+xml";
        else if (type == XLContentType::Theme)
            typeString = applicationOpenXmlOfficeDocument + ".theme+xml";
        else if (type == XLContentType::Styles)
            typeString = applicationOpenXmlOfficeDocument + ".spreadsheetml.styles+xml";
        else if (type == XLContentType::SharedStrings)
            typeString = applicationOpenXmlOfficeDocument + ".spreadsheetml.sharedStrings+xml";
        else if (type == XLContentType::Drawing)
            typeString = applicationOpenXmlOfficeDocument + ".drawing+xml";
        else if (type == XLContentType::Chart)
            typeString = applicationOpenXmlOfficeDocument + ".drawingml.chart+xml";
        else if (type == XLContentType::ChartStyle)
            typeString = applicationMicrosoftOffice + ".chartstyle+xml";
        else if (type == XLContentType::ChartColorStyle)
            typeString = applicationMicrosoftOffice + ".chartcolorstyle+xml";
        else if (type == XLContentType::ControlProperties)
            typeString = applicationMicrosoftExcel + ".controlproperties+xml";
        else if (type == XLContentType::CalculationChain)
            typeString = applicationOpenXmlOfficeDocument + ".spreadsheetml.calcChain+xml";
        else if (type == XLContentType::VBAProject)
            typeString = applicationMicrosoftOffice + ".vbaProject";
        else if (type == XLContentType::CoreProperties)
            typeString = applicationOpenXmlPackage + ".core-properties+xml";
        else if (type == XLContentType::ExtendedProperties)
            typeString = applicationOpenXmlOfficeDocument + ".extended-properties+xml";
        else if (type == XLContentType::CustomProperties)
            typeString = applicationOpenXmlOfficeDocument + ".custom-properties+xml";
        else if (type == XLContentType::Comments)
            typeString = applicationOpenXmlOfficeDocument + ".spreadsheetml.comments+xml";
        else if (type == XLContentType::Table)
            typeString = applicationOpenXmlOfficeDocument + ".spreadsheetml.table+xml";
        else if (type == XLContentType::VMLDrawing)
            typeString = applicationOpenXmlOfficeDocument + ".vmlDrawing";
        else
            throw XLInternalError("Unknown ContentType");

        return typeString;
    }
}    // namespace

/**
 * @details
 */
XLContentItem::XLContentItem() : m_contentNode(std::make_unique<XMLNode>()) {}

/**
 * @details
 */
XLContentItem::XLContentItem(const XMLNode& node) : m_contentNode(std::make_unique<XMLNode>(node)) {}

/**
 * @details
 */
XLContentItem::XLContentItem(const XLContentItem& other) : m_contentNode(std::make_unique<XMLNode>(*other.m_contentNode)) {}

/**
 * @details
 */
XLContentItem::XLContentItem(XLContentItem&& other) noexcept = default;

/**
 * @details
 */
XLContentItem::~XLContentItem() = default;

/**
 * @details
 */
XLContentItem& XLContentItem::operator=(const XLContentItem& other)
{
    if (&other != this) *m_contentNode = *other.m_contentNode;
    return *this;
}

XLContentItem& XLContentItem::operator=(XLContentItem&& other) noexcept = default;

/**
 * @details
 */
XLContentType XLContentItem::type() const { return GetTypeFromString(m_contentNode->attribute("ContentType").value()); }

/**
 * @details
 */
std::string XLContentItem::path() const { return m_contentNode->attribute("PartName").value(); }

/**
 * @details
 */
XLContentTypes::XLContentTypes() = default;

/**
 * @details
 */
XLContentTypes::XLContentTypes(XLXmlData* xmlData) : XLXmlFile(xmlData) {}

/**
 * @details
 */
XLContentTypes::~XLContentTypes() = default;

/**
 * @details
 */
XLContentTypes::XLContentTypes(const XLContentTypes& other) = default;

/**
 * @details
 */
XLContentTypes::XLContentTypes(XLContentTypes&& other) noexcept = default;

/**
 * @details
 */
XLContentTypes& XLContentTypes::operator=(const XLContentTypes& other) = default;

/**
 * @details
 */
XLContentTypes& XLContentTypes::operator=(XLContentTypes&& other) noexcept = default;

/**
 * @details
 * @note 2024-07-22: added more intelligent whitespace support
 */
void XLContentTypes::addOverride(const std::string& path, XLContentType type)
{
    const std::string typeString = GetStringFromType(type);

    XMLNode lastOverride = xmlDocument().document_element().last_child_of_type(pugi::node_element); // see if there's a last element
    XMLNode node{};    // scope declaration

    // Create new node in the [Content_Types].xml file
    if (lastOverride.empty())
        node = xmlDocument().document_element().prepend_child("Override");
    else { // if last element found
        // ===== Insert node after previous override
        node = xmlDocument().document_element().insert_child_after("Override", lastOverride);

        // ===== Using whitespace nodes prior to lastOverride as a template, insert whitespaces between lastOverride and the new node
        XMLNode copyWhitespaceFrom = lastOverride;    // start looking for whitespace nodes before previous override
        XMLNode insertBefore = node;                  // start inserting the same whitespace nodes before new override
        while (copyWhitespaceFrom.previous_sibling().type() == pugi::node_pcdata) { // empty node returns pugi::node_null
            // Advance to previous "template" whitespace node, ensured to exist in while-condition
            copyWhitespaceFrom = copyWhitespaceFrom.previous_sibling();
            // ===== Insert a whitespace node
            insertBefore = xmlDocument().document_element().insert_child_before(pugi::node_pcdata, insertBefore);
            insertBefore.set_value(copyWhitespaceFrom.value());            // copy the a whitespace in sequence node value
        }
    }
    node.append_attribute("PartName").set_value(path.c_str());
    node.append_attribute("ContentType").set_value(typeString.c_str());
}

/**
 * @details
 */
void XLContentTypes::deleteOverride(const std::string& path)
{
    xmlDocument().document_element().remove_child(xmlDocument().document_element().find_child_by_attribute("PartName", path.c_str()));
}

/**
 * @details
 */
void XLContentTypes::deleteOverride(const XLContentItem& item) { deleteOverride(item.path()); }

/**
 * @details
 */
XLContentItem XLContentTypes::contentItem(const std::string& path)
{
    return XLContentItem(xmlDocument().document_element().find_child_by_attribute("PartName", path.c_str()));
}

/**
 * @details
 */
std::vector<XLContentItem> XLContentTypes::getContentItems()
{
    std::vector<XLContentItem> result;
    XMLNode                    item = xmlDocument().document_element().first_child_of_type(pugi::node_element);
    while (not item.empty()) {
        if (strcmp(item.name(), "Override") == 0) result.emplace_back(item);
        item = item.next_sibling_of_type(pugi::node_element);
    }

    return result;
}

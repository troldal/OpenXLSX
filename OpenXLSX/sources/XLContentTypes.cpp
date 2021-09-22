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

using namespace OpenXLSX;

namespace
{
    /**
     * @details
     */
    XLContentType GetTypeFromString(const std::string& typeString)
    {
        XLContentType type;

        if (typeString == "application/vnd.ms-excel.Sheet.macroEnabled.main+xml")
            type = XLContentType::WorkbookMacroEnabled;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
            type = XLContentType::Workbook;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
            type = XLContentType::Worksheet;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml")
            type = XLContentType::Chartsheet;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml")
            type = XLContentType::ExternalLink;
        else if (typeString == "application/vnd.openxmlformats-officedocument.theme+xml")
            type = XLContentType::Theme;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
            type = XLContentType::Styles;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")
            type = XLContentType::SharedStrings;
        else if (typeString == "application/vnd.openxmlformats-officedocument.drawing+xml")
            type = XLContentType::Drawing;
        else if (typeString == "application/vnd.openxmlformats-officedocument.drawingml.chart+xml")
            type = XLContentType::Chart;
        else if (typeString == "application/vnd.ms-office.chartstyle+xml")
            type = XLContentType::ChartStyle;
        else if (typeString == "application/vnd.ms-office.chartcolorstyle+xml")
            type = XLContentType::ChartColorStyle;
        else if (typeString == "application/vnd.ms-excel.controlproperties+xml")
            type = XLContentType::ControlProperties;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml")
            type = XLContentType::CalculationChain;
        else if (typeString == "application/vnd.ms-office.vbaProject")
            type = XLContentType::VBAProject;
        else if (typeString == "application/vnd.openxmlformats-package.core-properties+xml")
            type = XLContentType::CoreProperties;
        else if (typeString == "application/vnd.openxmlformats-officedocument.extended-properties+xml")
            type = XLContentType::ExtendedProperties;
        else if (typeString == "application/vnd.openxmlformats-officedocument.custom-properties+xml")
            type = XLContentType::CustomProperties;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml")
            type = XLContentType::Comments;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml")
            type = XLContentType::Table;
        else if (typeString == "application/vnd.openxmlformats-officedocument.vmlDrawing")
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
            typeString = "application/vnd.ms-excel.Sheet.macroEnabled.main+xml";
        else if (type == XLContentType::Workbook)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        else if (type == XLContentType::Worksheet)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
        else if (type == XLContentType::Chartsheet)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";
        else if (type == XLContentType::ExternalLink)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
        else if (type == XLContentType::Theme)
            typeString = "application/vnd.openxmlformats-officedocument.theme+xml";
        else if (type == XLContentType::Styles)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
        else if (type == XLContentType::SharedStrings)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
        else if (type == XLContentType::Drawing)
            typeString = "application/vnd.openxmlformats-officedocument.drawing+xml";
        else if (type == XLContentType::Chart)
            typeString = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
        else if (type == XLContentType::ChartStyle)
            typeString = "application/vnd.ms-office.chartstyle+xml";
        else if (type == XLContentType::ChartColorStyle)
            typeString = "application/vnd.ms-office.chartcolorstyle+xml";
        else if (type == XLContentType::ControlProperties)
            typeString = "application/vnd.ms-excel.controlproperties+xml";
        else if (type == XLContentType::CalculationChain)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";
        else if (type == XLContentType::VBAProject)
            typeString = "application/vnd.ms-office.vbaProject";
        else if (type == XLContentType::CoreProperties)
            typeString = "application/vnd.openxmlformats-package.core-properties+xml";
        else if (type == XLContentType::ExtendedProperties)
            typeString = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
        else if (type == XLContentType::CustomProperties)
            typeString = "application/vnd.openxmlformats-officedocument.custom-properties+xml";
        else if (type == XLContentType::Comments)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
        else if (type == XLContentType::Table)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
        else if (type == XLContentType::VMLDrawing)
            typeString = "application/vnd.openxmlformats-officedocument.vmlDrawing";
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
XLContentType XLContentItem::type() const
{
    return GetTypeFromString(m_contentNode->attribute("ContentType").value());
}

/**
 * @details
 */
std::string XLContentItem::path() const
{
    return m_contentNode->attribute("PartName").value();
}

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
 */
void XLContentTypes::addOverride(const std::string& path, XLContentType type)
{
    std::string typeString = GetStringFromType(type);

    auto node = xmlDocument().first_child().append_child("Override");
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
void XLContentTypes::deleteOverride(XLContentItem& item)
{
    deleteOverride(item.path());
}

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
    for (auto item : xmlDocument().document_element().children()) {
        if (strcmp(item.name(), "Override") == 0) result.emplace_back(item);
    }

    return result;
}

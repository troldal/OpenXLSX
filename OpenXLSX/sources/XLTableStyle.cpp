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

    Written by Akira SHIMAHARA

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
#include <algorithm>
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLTableStyle.hpp"
#include "XLTemplates.hpp"
#include "XLTable.hpp"
#include "XLCommandQuery.hpp"

using namespace OpenXLSX;

XLTableStyle::XLTableStyle(const XMLNode& dataNode, XLTable* table): 
            m_dataNode(std::make_unique<XMLNode>(dataNode)),
            m_table(table)
{
    /* TODO implement the custom table style (with external xml file if required)
    */

}

XLTableStyle::XLTableStyle(const XLTableStyle& other)
        : m_dataNode(other.m_dataNode ? std::make_unique<XMLNode>(*other.m_dataNode) : nullptr),
         m_table(other.m_table)
{}

XLTableStyle::XLTableStyle(XLTableStyle&& other) noexcept
        : m_dataNode(std::move(other.m_dataNode)),
          m_table(other.m_table)
{}


XLTableStyle::~XLTableStyle() = default;

XLTableStyle& XLTableStyle::operator=(const XLTableStyle& other)
{
    if (&other != this) {
        auto temp = XLTableStyle(other);
        std::swap(*this, temp);
    }
    return *this;
}

XLTableStyle& XLTableStyle::operator=(XLTableStyle&& other) noexcept
{
    if (&other != this) {
        m_dataNode = std::move(other.m_dataNode);
        m_table  = other.m_table;
    }
    return *this;
}

bool XLTableStyle::columnStripes() const
{
    if(std::string(m_dataNode->attribute("showColumnStripes").value()) == "1")
        return true;
    
    return false;
}

void XLTableStyle::showColumnStripes(bool show) const
{
    auto node = m_dataNode->attribute("showColumnStripes");
    if(!node)
        node = m_dataNode->append_attribute("showColumnStripes");
    
    if(show)
        node.set_value("1");
    else
        node.set_value("0");
}


bool XLTableStyle::rowStripes() const
{
    if(std::string(m_dataNode->attribute("showRowStripes").value()) == "1")
        return true;
    
    return false;
}

void XLTableStyle::showRowStripes(bool show) const
{
    auto node = m_dataNode->attribute("showRowStripes");
    if(!node)
        node = m_dataNode->append_attribute("showRowStripes");
    
    if(show)
        node.set_value("1");
    else
       
       
    node.set_value("0");
}

bool XLTableStyle::firstColumnHighlighted() const
{
    if(std::string(m_dataNode->attribute("showFirstColumn").value()) == "1")
        return true;
    
    return false;
}

void XLTableStyle::showFirstColumnHighlighted(bool show) const
{
    auto node = m_dataNode->attribute("showFirstColumn");
    if(!node)
        node = m_dataNode->append_attribute("showFirstColumn");
    
    if(show)
        node.set_value("1");
    else
       
       
    node.set_value("0");
}

bool XLTableStyle::lastColumnHighlighted() const
{
    if(std::string(m_dataNode->attribute("showLastColumn").value()) == "1")
        return true;
    
    return false;
}

void XLTableStyle::showLastColumnHighlighted(bool show) const
{
    auto node = m_dataNode->attribute("showLastColumn");
    if(!node)
        node = m_dataNode->append_attribute("showLastColumn");
    
    if(show)
        node.set_value("1");
    else
       
       
    node.set_value("0");
}


const std::string XLTableStyle::style() const
{
    return std::string(m_dataNode->attribute("name").value());
}

void XLTableStyle::setStyle(const std::string& style) const
{
    if (!isTableStyleAvailable(style)){
        XLLogError("Specified style \"" + style + "\" is not available in this file");
        return;
    }
    
    m_dataNode->attribute("name").set_value(style.c_str());
}

bool XLTableStyle::isTableStyleAvailable(const std::string& style) const
{
    // check if the style is in the standard default list
    auto it = std::find(std::begin(XLTemplate::defaultTableStyle), 
                std::end(XLTemplate::defaultTableStyle), style);
    if( it != std::end(XLTemplate::defaultTableStyle))
        return true;

    // If not check also in the custom styles in the file
    m_table->m_pXmlData->getParentDoc();

    XLQuery query(XLQueryType::QueryXmlData);
    query.setParam<std::string>("xmlPath", "xl/styles.xml");

    XLXmlData* p = m_table->m_pXmlData->getParentDoc()->execQuery(query).template result<XLXmlData*>();
    if (p)
        for (const auto& node : p->getXmlDocument()->child("styleSheet").child("tableStyles"))
            if( style == std::string(node.attribute("name").value()))
                return true;

    return false;

}
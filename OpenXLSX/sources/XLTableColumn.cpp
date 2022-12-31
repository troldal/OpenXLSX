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
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "XLTableColumn.hpp"
#include "XLTemplates.hpp"
#include "XLTable.hpp"
using namespace OpenXLSX;

XLTableColumn::XLTableColumn(const XMLNode& dataNode, XLTable* table): 
            m_dataNode(std::make_unique<XMLNode>(dataNode)),
            m_table(table)
{
    /* TODO implement the followings
    m_name              = m_dataNode->attribute("name").value();
    m_xr3uid            = m_dataNode->attribute("xr3:uid").value();
    m_dataCellStyle     = m_dataNode->attribute("dataCellStyle").value();
    m_headerRowCellStyle= m_dataNode->attribute("headerRowCellStyle").value();
    m_totalsRowFunction = m_dataNode->attribute("totalsRowFunction").value();
    m_totalsRowLabel    = m_dataNode->attribute("totalsRowLabel").value();

    m_Id                = (m_dataNode->attribute("id"))              ? std::stoi(m_dataNode->attribute("id").value())             : (uint32_t) -1;
    m_dataDxfId         = (m_dataNode->attribute("dataDxfId"))       ? std::stoi(m_dataNode->attribute("dataDxfId").value())      : (uint32_t) -1;
    m_headerRowDxfId    = (m_dataNode->attribute("headerRowDxfId"))  ? std::stoi(m_dataNode->attribute("headerRowDxfId").value()) : (uint32_t) -1;
    m_totalsRowDxfId    = (m_dataNode->attribute("totalsRowDxfId"))  ? std::stoi(m_dataNode->attribute("totalsRowDxfId").value()) : (uint32_t) -1;

    if (m_dataNode->child("calculatedColumnFormula"))
        m_calculatedColumnFormula = m_dataNode->child("calculatedColumnFormula").child_value();
    if (m_dataNode->child("totalsRowFormula"))
        m_totalsRowFormula = m_dataNode->child("totalsRowFormula").child_value();
    */

}

XLTableColumn::XLTableColumn(const XLTableColumn& other)
        : m_dataNode(other.m_dataNode ? std::make_unique<XMLNode>(*other.m_dataNode) : nullptr),
         m_table(other.m_table)
{}

XLTableColumn::XLTableColumn(XLTableColumn&& other) noexcept
        : m_dataNode(std::move(other.m_dataNode)),
          m_table(other.m_table)
{}


XLTableColumn::~XLTableColumn() = default;

XLTableColumn& XLTableColumn::operator=(const XLTableColumn& other)
{
    if (&other != this) {
        auto temp = XLTableColumn(other);
        std::swap(*this, temp);
    }
    return *this;
}

XLTableColumn& XLTableColumn::operator=(XLTableColumn&& other) noexcept
{
    if (&other != this) {
        m_dataNode = std::move(other.m_dataNode);
        m_table  = other.m_table;
    }
    return *this;
}


std::string XLTableColumn::name() const
{
    return (m_dataNode->attribute("name").value());
}

void XLTableColumn::setName(const std::string& columnName) const
{
    m_dataNode->attribute("name").set_value(columnName.c_str());
    // TODO change the formulas  in table.xml
    // TODO change the formulas in the sheet
}

std::string XLTableColumn::totalsRowFunction() const
{
    return std::string(m_dataNode->attribute("totalsRowFunction").value());
}

void XLTableColumn::setTotalsRowFunction(const std::string& function)
{
    // TODO implement <totalsRowLabel>
    // TODO implement <totalsRowFormula> == custom function

    // check if the formula is allowed otherwise quit
    auto it = std::find(std::begin(XLTemplate::totalsRowFunctionList), 
                std::end(XLTemplate::totalsRowFunctionList), function);
    if( it == std::end(XLTemplate::totalsRowFunctionList))
        return;

    auto node = m_dataNode->attribute("totalsRowFunction");
    if(!node)
        node = m_dataNode->append_attribute("totalsRowFunction");
    
    node.set_value(function.c_str());

    // Update the fonctions in the worksheet
    m_table->setTotalFormulas();

}


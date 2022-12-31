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
#include "XLAutofilter.hpp"
#include "XLCellRange.hpp"
#include "XLXmlData.hpp"



using namespace OpenXLSX;


 XLAutofilter:: XLAutofilter(const XMLNode& dataNode, XLXmlData* parent): 
            m_dataNode(std::make_unique<XMLNode>(dataNode)),
            m_parent(parent)
{
}
// TODO check if Autofilter is valid
XLAutofilter::~XLAutofilter() = default;

std::string XLAutofilter::ref() const
{
    return std::string(m_dataNode->attribute("ref").value());
}

 
void XLAutofilter::setRef(const std::string& ref) const
{
    auto node = m_dataNode->attribute("ref");
    if(!node)
        node = m_dataNode->append_attribute("ref");
    
    node.set_value(ref.c_str());  
}

void XLAutofilter::hideArrow(uint16_t indexCol, bool hide) const
{
    auto node = m_dataNode->find_child_by_attribute("colId", std::to_string(indexCol).c_str());
    if (!node){
        node = m_dataNode->append_child("filterColumn");
        node.append_attribute("colId")
            .set_value(std::to_string(indexCol).c_str());
    }
    if(!node.attribute("hiddenButton"))
        node.append_attribute("hiddenButton");

    if(hide)
        node.attribute("hiddenButton").set_value("1");
    else
        node.attribute("hiddenButton").set_value("0");
}

void XLAutofilter::hideArrows(bool hide) const
{
    uint16_t nCol = XLCellRange::columnsCount(ref());
    for (uint16_t i = 0; i < nCol;i++)
        hideArrow(i,hide);
}
/*
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
*/

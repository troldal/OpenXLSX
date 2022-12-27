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
#include <string>
#include <pugixml.hpp>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "XLTable.hpp"


using namespace OpenXLSX;


XLTable::XLTable(XLXmlData* xmlData) : m_pXmlData(xmlData)
{
  /*
    using XMLNode      = pugi::xml_node;
    using XMLAttribute = pugi::xml_attribute;
    using XMLDocument  = pugi::xml_document;
    */


   
  


  // ===== Deal with the columns
  XMLNode pTblColumns =  m_pXmlData->getXmlDocument()->child("table").child("tableColumns");
  m_nColumnCount = std::stoi(pTblColumns.attribute("count").value());

  for (XMLNode col : pTblColumns.children()){
    std::string test = col.attribute("name").value();
    m_columns.emplace_back(XLTableColumn(&col));
  }
  
  for (auto item : m_columns){
    std::string test = item.name();
    int i =0;
  } 
    /*
    XMLDocument* pTableDoc = xmlData->getXmlDocument();
  m_name = pTableDoc->child("table").attribute("name").value();
    */
//for (const auto& node : getXmlData("xl/sharedStrings.xml")->getXmlDocument()->document_element().children()

//<tableColumns count="5">
  int i = 0;
}

std::string XLTable::name() const
{
  return (m_pXmlData->getXmlDocument()->child("table").attribute("name").value());
}

std::string XLTable::ref() const
{
  return (m_pXmlData->getXmlDocument()->child("table").attribute("ref").value());
}

uint16_t XLTable::columnCount() const
{
  return (std::stoi(m_pXmlData->getXmlDocument()->child("table")
            .child("tableColumns").attribute("count").value()));
}

void XLTable::setName(const std::string& tableName)
{
  XMLNode tableNode = m_pXmlData->getXmlDocument()->child("table");
  tableNode.attribute("name").set_value(tableName.c_str());
  tableNode.attribute("displayName").set_value(tableName.c_str());

  // TODO change the formulas  in table.xml
  // TODO change the formulas in the sheet


}



/*XLCellRange XLWorksheet::range(const XLCellReference& topLeft, const XLCellReference& bottomRight) const
{
    return XLCellRange(xmlDocument().first_child().child("sheetData"),
                       topLeft,
                       bottomRight,
                       parentDoc().execQuery(XLQuery(XLQueryType::QuerySharedStrings)).result<XLSharedStrings>());
}
*/
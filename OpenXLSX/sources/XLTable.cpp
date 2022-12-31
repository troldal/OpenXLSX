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
#include <string>
#include <pugixml.hpp>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "XLSheet.hpp"
#include "XLTable.hpp"
#include "XLCellRange.hpp"
#include "XLCellReference.hpp"


using namespace OpenXLSX;


XLTable::XLTable(XLXmlData* xmlData) 
      : m_pXmlData(xmlData),m_sheet(XLWorksheet(xmlData->getParentNode()))
{
  
  // ===== Deal with the columns
  XMLNode pTblColumns =  m_pXmlData->getXmlDocument()->child("table").child("tableColumns");

  for (const XMLNode& col : pTblColumns.children())
    m_columns.emplace_back(std::shared_ptr<XLTableColumn>(new XLTableColumn(col)));

}

XLTable::~XLTable()
{}

const std::string XLTable::name() const
{
  // Check if displayName is better
  return (m_pXmlData->getXmlDocument()->child("table").attribute("name").value());
}

const std::string XLTable::ref() const
{
  return (m_pXmlData->getXmlDocument()->child("table").attribute("ref").value());
}

std::vector<std::string>  XLTable:: columnNames() const
{
  std::vector<std::string> colNames;
  for (auto col : m_columns)
    colNames.push_back(col->name());
  
  return colNames;
}

uint16_t XLTable::columnIndex(const std::string& name) const
{
  auto it = std::find_if(m_columns.begin(),
    m_columns.end(), 
    [&](const auto& val){ return val->name() == name; } );
  
  if (it !=m_columns.end())
    return (it - m_columns.begin());
  
  return (uint16_t)(-1);
}

XLWorksheet* XLTable::getWorksheet()
{
  return &m_sheet;
}

XLCellRange XLTable::tableRange()
{
  return getWorksheet()->range(ref());
}

XLTableRows XLTable::tableRows()
{
    std::pair<std::string,std::string> p = XLCellRange::topLeftBottomRight(ref());
    XLCellReference topLeft(p.first);
    XLCellReference bottomRight(p.second);

    uint32_t firstRow = topLeft.row();
    uint32_t lastRow = bottomRight.row();

    uint16_t firstCol = topLeft.column(); 
    uint16_t lastCol = bottomRight.column();

    if (isHeaderVisible())
        firstRow +=1;
    
    if (isTotalVisible())
        lastRow -=1;
    
    return XLTableRows(m_pXmlData->getParentNode()->getXmlDocument()->first_child().child("sheetData"),
            firstRow, lastRow, firstCol, lastCol, getWorksheet());
}

XLCellRange XLTable::dataBodyRange()
{
  std::pair<std::string,std::string> p = XLCellRange::topLeftBottomRight(ref());
  XLCellReference topLeft(p.first);
  XLCellReference bottomRight(p.second);

  if (isHeaderVisible())
    topLeft.offset(1);
  
  if (isTotalVisible())
    bottomRight.offset(-1);
  
  return getWorksheet()->range(topLeft,bottomRight);
}

bool XLTable::isHeaderVisible() const
{
  std::string header = m_pXmlData->getXmlDocument()->child("table").attribute("headerRowCount").value();
  if (header == "0")
    return false;
  return true; // if missing, visible
}

bool XLTable::isTotalVisible() const
{
  std::string header = m_pXmlData->getXmlDocument()->child("table").attribute("totalsRowCount").value();
  if (header == "1")
    return true;
  return false; // if missing, not visible
}



uint16_t XLTable::columnsCount() const
{
  return (std::stoi(m_pXmlData->getXmlDocument()->child("table")
            .child("tableColumns").attribute("count").value()));
}

uint32_t XLTable::rowsCount() const
{
    std::pair<std::string,std::string> p = XLCellRange::topLeftBottomRight(ref());
    XLCellReference topLeft(p.first);
    XLCellReference bottomRight(p.second);

    uint32_t firstRow = topLeft.row();
    uint32_t lastRow = bottomRight.row();

    if (isHeaderVisible())
        firstRow +=1;
    
    if (isTotalVisible())
        lastRow -=1;
    
    return (lastRow - firstRow + 1);
}


void XLTable::setName(const std::string& tableName)
{
  XMLNode tableNode = m_pXmlData->getXmlDocument()->child("table");
  tableNode.attribute("name").set_value(tableName.c_str());
  tableNode.attribute("displayName").set_value(tableName.c_str());

  // TODO change the formulas  in table.xml
  // TODO change the formulas in the sheet


}


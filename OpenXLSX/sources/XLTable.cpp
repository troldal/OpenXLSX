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
#include "XLTemplates.hpp"


using namespace OpenXLSX;


XLTable::XLTable(XLXmlData* xmlData) 
      : m_pXmlData(xmlData),
        m_sheet(XLWorksheet(xmlData->getParentNode())),
        m_dataBodyRange(getWorksheet()->range()) // Fake range
{
    std::pair<std::string,std::string> p = XLCellRange::topLeftBottomRight(ref());
    XLCellReference topLeft(p.first);
    XLCellReference bottomRight(p.second);

    if (isHeaderVisible())
        topLeft.offset(1);
    
    if (isTotalVisible())
        bottomRight.offset(-1);
    
    m_dataBodyRange.setRangeCoordinates(topLeft, bottomRight);

    // ===== Deal with the columns
    updateColumns();

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
    for (auto& col : m_columns)
        colNames.push_back(col.name());
  
  return colNames;
}

uint16_t XLTable::columnIndex(const std::string& name) const
{
    auto it = std::find_if(m_columns.begin(),
        m_columns.end(), 
        [&](const auto& val){ return val.name() == name; } );
    
    if (it !=m_columns.end())
        return (it - m_columns.begin());
    
    return (uint16_t)(-1);
}


XLTableColumn& XLTable::column(const std::string& name)
{
    uint16_t index = columnIndex(name);
    if (index == (uint16_t)(-1))
        index = 0;
    
    return m_columns[index];
}

const std::vector<XLTableColumn>& XLTable::columns() const
{
    return m_columns;
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
    return XLTableRows(m_pXmlData->getParentNode()->getXmlDocument()->first_child().child("sheetData"),
            m_dataBodyRange.rangeCoordinates().first.row(),
            m_dataBodyRange.rangeCoordinates().second.row(),
            m_dataBodyRange.rangeCoordinates().first.column(),
            m_dataBodyRange.rangeCoordinates().second.column(),
            getWorksheet());
}

XLCellRange& XLTable::dataBodyRange()
{
  return m_dataBodyRange;
}

const XLCellRange& XLTable::dataBodyRange() const 
{
  return m_dataBodyRange;
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
  std::string total = m_pXmlData->getXmlDocument()->child("table").attribute("totalsRowCount").value();
  if (total == "1")
    return true;
  return false; // if missing, not visible
}


void XLTable::setHeaderVisible(bool visible)
{
    if(visible){ // removing the attribute if visible
        m_pXmlData->getXmlDocument()->child("table").remove_attribute("headerRowCount");
    } else {
        auto node = m_pXmlData->getXmlDocument()->child("table").attribute("headerRowCount");
        if (!node)
            node = m_pXmlData->getXmlDocument()->child("table").append_attribute("headerRowCount");
        node.set_value("0"); 
    }

    // TODO remove autofilter

    setHeaderLabels();
    adjustRef();
}


void XLTable::setTotalVisible(bool visible)
{
    if(visible){ 
        m_pXmlData->getXmlDocument()->child("table").remove_attribute("totalsRowShown");

        auto node = m_pXmlData->getXmlDocument()->child("table").attribute("totalsRowCount");
        if (!node)
            node = m_pXmlData->getXmlDocument()->child("table").append_attribute("totalsRowCount");
        node.set_value("1");

    } else { // removing the attribute if not visible
        m_pXmlData->getXmlDocument()->child("table").remove_attribute("totalsRowCount");
        auto node = m_pXmlData->getXmlDocument()->child("table").attribute("totalsRowShown");
        if (!node)
            node = m_pXmlData->getXmlDocument()->child("table").append_attribute("totalsRowShown");
        node.set_value("0");
        //TODO check if required to clear the worksheet when hiding total row
    }

    setTotalFormulas();
    setTotalLabels();
    adjustRef();
    
}

XLAutofilter XLTable::autofilter()
{
    // TODO implement auto filter and adjust header masking accordingly
    return XLAutofilter(m_pXmlData->getXmlDocument()->child("table").child("autoFilter"),m_pXmlData);
}

XLTableStyle XLTable::tableStyle()
{
    return XLTableStyle(m_pXmlData->getXmlDocument()->child("table").child("tableStyleInfo"), this);
}

uint16_t XLTable::columnsCount() const
{
  return (std::stoi(m_pXmlData->getXmlDocument()->child("table")
            .child("tableColumns").attribute("count").value()));
}

uint32_t XLTable::rowsCount() const
{
    auto p = m_dataBodyRange.rangeCoordinates();

    uint32_t firstRow = p.first.row();
    uint32_t lastRow = p.second.row();
    
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

 XLTableColumn& XLTable::insertColumn(const std::string& columnName, uint16_t index)
 {
    
    // If index is not continuous, the column is appended at the last
    if (index > m_columns.size())
        index = m_columns.size();

    // Increment the name of the column if required
    std::string colName = columnName;
    bool notValid = true;
    while (notValid){
        if (std::find_if(m_columns.begin(), m_columns.end(), 
                    [&colName](const XLTableColumn& i){ return i.name() == colName; }) != m_columns.end())
            colName += INCREMENT_STRING;
        else
            notValid = false;
    }

    // Copy the content of the bodyrange of the column, anr re index tablecolumn in table.xml
    for (uint16_t i = m_columns.size() - 1; i >= index; --i) {
        auto& tblCol = m_columns[i];
        
        // Change the index in table.xml
        tblCol.setIndex(i+2); // 1 based index

        // Copy column n in n+1 in the sheet
        XLCellRange src = tblCol.bodyRange();
        XLCellRange dest = src;
        dest.offset(0,1);
        for(uint32_t j = 0; j < src.numRows(); j++){
            dest[j].value() = src[j].value();
            if ( src[j].hasFormula())
                dest[j].formula() = src[j].formula();
            else
                dest[j].formula().clear();
        }
    }

    // add new entry in table.xml
    auto nextNode = m_pXmlData->getXmlDocument()->document_element()
                    .child("tableColumns").find_child_by_attribute("id",std::to_string(index+2).c_str());
    
    XMLNode newNode;
    if(nextNode)
        newNode = m_pXmlData->getXmlDocument()->document_element()
                    .child("tableColumns").insert_child_before("tableColumn", nextNode);
    else
        newNode = m_pXmlData->getXmlDocument()->document_element()
                    .child("tableColumns").append_child("tableColumn");

    
    newNode.append_attribute("id").set_value(std::to_string(index + 1).c_str());
    newNode.append_attribute("name").set_value(colName.c_str());

    // Update the range coordinates
    auto p = m_dataBodyRange.rangeCoordinates();
    m_dataBodyRange.setRangeCoordinates(p.first, p.second.offset(0,1));
 

    updateColumns();    // Update the member variable

    setHeaderLabels();  // relocate the labels
    setTotalLabels();   // relocate the formulas
    setTotalLabels();   // reolcate the lables
    adjustRef();        // update the ref

    //Clear the content of the new column
    for(auto& cell : m_columns[index].bodyRange())
        cell.value().clear();
    
    return m_columns[index];
 }

XLTableColumn& XLTable::appendColumn(const std::string& columnName)
{
    return insertColumn(columnName, m_columns.size());
}

void XLTable::updateColumns()
{
    m_columns.clear();

    XMLNode pTblColumns =  m_pXmlData->getXmlDocument()->child("table").child("tableColumns");

    for (const XMLNode& col : pTblColumns.children())
        m_columns.push_back(XLTableColumn(col, *this));
}


void XLTable::setColumnFormulas() const
{
    // TO be implemented
    uint32_t firstRow = m_dataBodyRange.rangeCoordinates().first.row();
    uint16_t firstCol = m_dataBodyRange.rangeCoordinates().first.column();

    for (auto& col: m_columns){
        std::string colFormula = col.columnFormula();
        if (colFormula.empty())  // if there is nothing, remove formula keep values
            for (auto& cell: col.bodyRange())
                cell.formula().clear();  
        else
            for (auto& cell: col.bodyRange()){
                cell.formula() = colFormula;
                cell.value().clear();
            }
    }
    
}

void XLTable::setTotalFormulas() const
{
    uint32_t totalRow = m_dataBodyRange.rangeCoordinates().second.row() + 1;
    uint16_t totalCol = m_dataBodyRange.rangeCoordinates().first.column();
    
    // Hide everything if the total is hidded
    if (!isTotalVisible()){
        for (auto& col: m_columns){
            m_sheet.cell(totalRow, totalCol).formula().clear();
            totalCol += 1;
        }
        return;
    }
    
    for (auto& col: m_columns){
        std::string colFunc = col.totalsRowFormula();
        if (colFunc.empty()){  // if there is nothing, remove formula and values
            m_sheet.cell(totalRow, totalCol).formula().clear(); 
        }
        else
        {
            // Check if its exist in the list
            auto it = std::find(std::begin(XLTemplate::totalsRowFunctionList), 
                    std::end(XLTemplate::totalsRowFunctionList), colFunc);
            if( it != std::end(XLTemplate::totalsRowFunctionList)){

                std::string function =  XLTemplate::sheetTotalFunctionList[
                            std::distance(std::begin(XLTemplate::totalsRowFunctionList), it)
                            ];
                if (function != "none")
                {
                    // get the equivalent formula in the sheet
                     std::string sheetFunction =  "SUBTOTAL(" + function 
                                + "," + name() 
                                    + "[" + col.name() + "])";
                     m_sheet.cell(totalRow, totalCol).formula() = sheetFunction;
                } else {// remove the formula if function is none 
                    m_sheet.cell(totalRow, totalCol).formula().clear();
                }
               
            }
        } // if col function not empty
        totalCol +=1;
    }
    
}

void XLTable::setTotalLabels() const
{
    uint32_t totalRow = m_dataBodyRange.rangeCoordinates().second.row() + 1;
    uint16_t totalCol = m_dataBodyRange.rangeCoordinates().first.column();
    
    // Hide everything if the total is hidded
    if (!isTotalVisible()){
        for (auto& col: m_columns){
            m_sheet.cell(totalRow, totalCol).value().clear();
            totalCol += 1;
        }
        return;
    }

    for (auto& col: m_columns){
        std::string colLabel = col.totalsRowLabel();
        if (colLabel.empty()){  // if there is nothing, remove values
            m_sheet.cell(totalRow, totalCol).value().clear(); 
        }
        else
            m_sheet.cell(totalRow, totalCol).value() = colLabel;

        totalCol += 1;
    }
    
}

void XLTable::setHeaderLabels() const
{
    uint32_t headerRow = m_dataBodyRange.rangeCoordinates().first.row() - 1;
    uint16_t headerCol = m_dataBodyRange.rangeCoordinates().first.column();
    
    // Hide everything if the total is hidded
    if (!isHeaderVisible()){
        for (auto& col: m_columns){
            m_sheet.cell(headerRow, headerCol).value().clear();
            headerCol += 1;
        }
        return;
    }

    for (auto& col: m_columns){
        auto test = col.name();
        m_sheet.cell(headerRow, headerCol).value() = col.name();

        headerCol += 1;
    }
    
}


void XLTable::adjustRef()
{
    uint32_t firstRow   = m_dataBodyRange.rangeCoordinates().first.row();
    uint32_t lastRow    = m_dataBodyRange.rangeCoordinates().second.row();

    uint16_t firstCol   = m_dataBodyRange.rangeCoordinates().first.column();
    uint16_t lastCol    = m_dataBodyRange.rangeCoordinates().second.column();

    if(isHeaderVisible())
        firstRow -=1;

    if(isTotalVisible())
        lastRow +=1;
    
    std::string ref  = XLCellReference::adressFromCoordinates(firstRow, firstCol, false);
    ref       += ":" + XLCellReference::adressFromCoordinates(lastRow,  lastCol,  false);

    m_pXmlData->getXmlDocument()->child("table")
                    .attribute("ref").set_value(ref.c_str());

    // TODO adjust autofilter Move this elsewhere
}
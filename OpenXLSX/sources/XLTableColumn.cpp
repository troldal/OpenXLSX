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

XLTableColumn::XLTableColumn(const XMLNode& dataNode, const XLTable& table): 
            m_dataNode(std::make_shared<XMLNode>(dataNode)),
            m_table(table),
            m_proxyTotal(XLTableColumnTotalProxy(dataNode,"totalsRowFunction", table)),
            m_proxyColumn(XLTableColumnFormulaProxy(dataNode,"calculatedColumnFormula", table)),
            m_proxyLabel(XLTableColumnTotalLabelProxy(dataNode,"totalsRowLabel", table))
{
    // TODO implement <totalsRowLabel>
    // TODO implement <totalsRowFormula> == custom function
    
    /* TODO implement the followings
    
    
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
        : m_dataNode(other.m_dataNode ? std::make_shared<XMLNode>(*other.m_dataNode) : nullptr),
         m_table(other.m_table),
         m_proxyTotal(XLTableColumnTotalProxy((*other.m_dataNode),"totalsRowFunction", other.m_table)),
         m_proxyColumn(XLTableColumnFormulaProxy((*other.m_dataNode),"calculatedColumnFormula", other.m_table)),
         m_proxyLabel(XLTableColumnTotalLabelProxy(*other.m_dataNode,"totalsRowLabel", other.m_table))
{}

XLTableColumn::XLTableColumn(XLTableColumn&& other) noexcept
        : m_dataNode(std::move(other.m_dataNode)),
          m_table(other.m_table),
          m_proxyTotal(std::move(other.m_proxyTotal)),
          m_proxyColumn( std::move(other.m_proxyColumn)),
          m_proxyLabel( std::move(other.m_proxyLabel))
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
    }
    return *this;
}


std::string XLTableColumn::name() const
{
    return std::string(m_dataNode->attribute("name").value());
}

void XLTableColumn::setName(const std::string& columnName) const
{
    m_dataNode->attribute("name").set_value(columnName.c_str());
    // TODO change the formulas  in table.xml
    // TODO change the formulas in the sheet
}

XLTableColumnProxy& XLTableColumn::totalsRowFormula()
{
    return m_proxyTotal;
}

const XLTableColumnProxy& XLTableColumn::totalsRowFormula() const
{
    return m_proxyTotal;
}

void XLTableColumn::clearTotalsRowFormula()
{
    m_proxyTotal.clear();
}

XLTableColumnFormulaProxy& XLTableColumn::columnFormula()
{
    return m_proxyColumn;
}

const XLTableColumnFormulaProxy& XLTableColumn::columnFormula() const
{
    return m_proxyColumn;
}

void XLTableColumn::clearColumnFormula()
{
    m_proxyColumn.clear();
}

XLTableColumnProxy& XLTableColumn::totalsRowLabel()
{
    return m_proxyLabel;
}

const XLTableColumnProxy& XLTableColumn::totalsRowLabel() const
{
    return m_proxyLabel;
}

void XLTableColumn::clearTotalsRowLabel()
{
    m_proxyLabel.clear();
}

XLCellRange XLTableColumn::bodyRange() const
{
    uint16_t colIndex = m_table.columnIndex(name());
    XLCellRange tableRange = m_table.dataBodyRange();
    auto p = tableRange.rangeCoordinates();
    uint16_t  firstCol = p.first.column();
    XLCellReference tl = p.first;
    XLCellReference br = p.second;

    tl.setColumn(firstCol + colIndex);
    br.setColumn(firstCol + colIndex);

    tableRange.setRangeCoordinates(tl, br);
    return tableRange;
}

/////////////////////////////////////////////////////////////////////////////////////////
//
//                  XLTableColumnProxy
//
/////////////////////////////////////////////////////////////////////////////////////////


/**
 * @details Constructor
 * @pre The attr pointer must not be nullptr and must point to valid objects.
 * @post A valid XLCellValueProxy has been created.
 */
XLTableColumnProxy::XLTableColumnProxy(const XMLNode& node, const std::string& attr, const XLTable& table ) 
                        : m_node(std::make_shared<XMLNode>(node)), 
                          m_attribute(attr),
                          m_table(table)
{
    assert(node);                  // NOLINT
}

/**
 * @details Destructor. Default implementation has been used.
 * @pre
 * @post
 */
XLTableColumnProxy::~XLTableColumnProxy() = default;

/**
 * @details Copy constructor. Default implementation has been used.
 * @pre
 * @post
 */
XLTableColumnProxy::XLTableColumnProxy(const XLTableColumnProxy& other)
                : m_node(other.m_node ? std::make_shared<XMLNode>(*other.m_node) : nullptr),
                  m_attribute(other.m_attribute),
                  m_table(other.m_table)
{}

/**
 * @details Move constructor. Default implementation has been used.
 * @pre
 * @post
 */
XLTableColumnProxy::XLTableColumnProxy(XLTableColumnProxy&& other) noexcept
                            : m_node(std::move(other.m_node)),
                              m_attribute(std::move(other.m_attribute)),
                              m_table(std::move(other.m_table))
{}

/**
 * @details Copy assignment operator. The function is implemented in terms of the templated
 * value assignment operators, i.e. it is the XLCellValue that is that is copied,
 * not the object itself.
 * @pre
 * @post
 */
XLTableColumnProxy& XLTableColumnProxy::operator=(const XLTableColumnProxy& other)
{
    if (&other != this) {
        *this = other.getFormula();
    }

    return *this;
}


/**
 * @details Move assignment operator. Default implementation has been used.
 * @pre
 * @post
 */
XLTableColumnProxy& XLTableColumnProxy::operator=(XLTableColumnProxy&& other) noexcept
{
    if (&other != this) {
        m_node          = std::move(other.m_node);
        m_attribute     = std::move(other.m_attribute);
    }

    return *this;
}



/////////////////////////////////////////////////////////////////////////////////////////
//
//                  XLTableColumnTotalProxy
//
/////////////////////////////////////////////////////////////////////////////////////////

XLTableColumnProxy& XLTableColumnTotalProxy::operator=(const std::string& formula)
{   
    setFormula(formula);
    return *this;
}

/**
 * @details Clear the contents of the cell. This removes all children of the cell node.
 * @pre The m_cellNode must not be null, and must point to a valid XML cell node object.
 * @post The cell node must be valid, but empty.
 */
void XLTableColumnTotalProxy::clear()
{
    // ===== Check that the m_attribute is valid.
    assert(m_node);              // NOLINT
    m_node->remove_attribute(m_attribute.c_str());
    m_table.setTotalFormulas();
    
}

/**
 * @details Set the the formula in the corresponding attribute.
 *  This is private helper function for setting the attr 
 * directly in the underlying XML file.
 * @pre The m_attribute must not be null, and must point to a valid XMLNode object.
 * @post The underlying attribute has been updated correctly, representing a string value.
 */
void XLTableColumnTotalProxy::setFormula(const std::string& formula)
{
    // ===== Check that the m_cellNode is valid.
    assert(m_node);              // NOLINT
    auto node = m_node->attribute(m_attribute.c_str());
    if (!node)
        node = m_node->append_attribute(m_attribute.c_str());
    
     // If empty string, we remove the formula
    if(formula.empty()){
        clear();
        return;
    }

    auto it = std::find(std::begin(XLTemplate::totalsRowFunctionList), 
                    std::end(XLTemplate::totalsRowFunctionList), formula);
    if( it == std::end(XLTemplate::totalsRowFunctionList)){
        XLLogError("The formula \"" + formula + "\" is not available for total Row function");
        return;
    }

    m_node->attribute(m_attribute.c_str()).set_value(formula.c_str());
    m_table.setTotalFormulas();
}

/**
 * @details Get a copy of the XLCellValue object for the cell. This is private helper function for returning an
 * XLCellValue object corresponding to the cell value.
 * @pre The m_cellNode must not be null, and must point to a valid XMLNode object.
 * @post No changes should be made.
 */
std::string XLTableColumnTotalProxy::getFormula() const
{
    // ===== Check that the m_attribute is valid.
    assert(m_node);              // NOLINT
    auto node = m_node->attribute(m_attribute.c_str());
    if (!node)
        return std::string();
    
    return std::string(m_node->attribute(m_attribute.c_str()).value());
}
/////////////////////////////////////////////////////////////////////////////////////////
//
//                  XLTableColumnFormulaProxy
//
/////////////////////////////////////////////////////////////////////////////////////////

XLTableColumnProxy& XLTableColumnFormulaProxy::operator=(const std::string& formula)
{
    setFormula(formula);
    return *this;
}

void XLTableColumnFormulaProxy::clear()
{
    // ===== Check that the m_attribute is valid.
    assert(m_node);              // NOLINT
    m_node->remove_child(m_attribute.c_str());
    m_table.setColumnFormulas();
    
}

void XLTableColumnFormulaProxy::setFormula(const std::string& formula)
{
    // ===== Check that the m_cellNode is valid.
    assert(m_node);              // NOLINT
    auto node = m_node->child(m_attribute.c_str());
    if (!node)
        node = m_node->append_child(m_attribute.c_str());
    
     // If empty string, we remove the formula
    if(formula.empty()){
        clear();
        return;
    }

    m_node->child(m_attribute.c_str()).text().set(formula.c_str());
    m_table.setColumnFormulas();
}

std::string XLTableColumnFormulaProxy::getFormula() const
{
    // ===== Check that the m_attribute is valid.
    assert(m_node);              // NOLINT
    auto node = m_node->child(m_attribute.c_str());
    if (!node)
        return std::string();
    
    return std::string(m_node->child(m_attribute.c_str()).text().get());

}

/////////////////////////////////////////////////////////////////////////////////////////
//
//                  XLTableColumnTotalLabelProxy
//
/////////////////////////////////////////////////////////////////////////////////////////

XLTableColumnProxy& XLTableColumnTotalLabelProxy::operator=(const std::string& formula)
{   
    setFormula(formula);
    return *this;
}

/**
 * @details Clear the contents of the cell. This removes all children of the cell node.
 * @pre The m_cellNode must not be null, and must point to a valid XML cell node object.
 * @post The cell node must be valid, but empty.
 */
void XLTableColumnTotalLabelProxy::clear()
{
    // ===== Check that the m_attribute is valid.
    assert(m_node);              // NOLINT
    m_node->remove_attribute(m_attribute.c_str());
    m_table.setTotalLabels();
    
}

/**
 * @details Set the the formula in the corresponding attribute.
 *  This is private helper function for setting the attr 
 * directly in the underlying XML file.
 * @pre The m_attribute must not be null, and must point to a valid XMLNode object.
 * @post The underlying attribute has been updated correctly, representing a string value.
 */
void XLTableColumnTotalLabelProxy::setFormula(const std::string& formula)
{
    // ===== Check that the m_cellNode is valid.
    assert(m_node);              // NOLINT
    auto node = m_node->attribute(m_attribute.c_str());
    if (!node)
        node = m_node->append_attribute(m_attribute.c_str());
    
     // If empty string, we remove the formula
    if(formula.empty()){
        clear();
        return;
    }

    m_node->attribute(m_attribute.c_str()).set_value(formula.c_str());
    m_table.setTotalLabels();
}

/**
 * @details Get a copy of the XLCellValue object for the cell. This is private helper function for returning an
 * XLCellValue object corresponding to the cell value.
 * @pre The m_cellNode must not be null, and must point to a valid XMLNode object.
 * @post No changes should be made.
 */
std::string XLTableColumnTotalLabelProxy::getFormula() const
{
    // ===== Check that the m_attribute is valid.
    assert(m_node);              // NOLINT
    auto node = m_node->attribute(m_attribute.c_str());
    if (!node)
        return std::string();
    
    return std::string(m_node->attribute(m_attribute.c_str()).value());
}
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
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLCell.hpp"
#include "XLCellRange.hpp"
#include "utilities/XLUtilities.hpp"

using namespace OpenXLSX;

/**
 * @details
 */
XLCell::XLCell()
    : m_cellNode(nullptr),
      m_valueProxy(XLCellValueProxy(this, m_cellNode.get())),
      m_formulaProxy(XLFormulaProxy(this, m_cellNode.get()))
{}

/**
 * @details This constructor creates a XLCell object based on the cell XMLNode input parameter, and is
 * intended for use when the corresponding cell XMLNode already exist.
 * If a cell XMLNode does not exist (i.e., the cell is empty), use the relevant constructor to create an XLCell
 * from a XLCellReference parameter.
 */
XLCell::XLCell(const XMLNode& cellNode, const XLSharedStrings& sharedStrings)
    : m_cellNode(std::make_unique<XMLNode>(cellNode)),
      m_sharedStrings(sharedStrings),
      m_valueProxy(XLCellValueProxy(this, m_cellNode.get())),
      m_formulaProxy(XLFormulaProxy(this, m_cellNode.get()))
{}

/**
 * @details
 */
XLCell::XLCell(const XLCell& other)
    : m_cellNode(other.m_cellNode ? std::make_unique<XMLNode>(*other.m_cellNode) : nullptr),
      m_sharedStrings(other.m_sharedStrings),
      m_valueProxy(XLCellValueProxy(this, m_cellNode.get())),
      m_formulaProxy(XLFormulaProxy(this, m_cellNode.get()))
{}

/**
 * @details
 */
XLCell::XLCell(XLCell&& other) noexcept
    : m_cellNode(std::move(other.m_cellNode)),
      m_sharedStrings(std::move(other.m_sharedStrings)),
      m_valueProxy(XLCellValueProxy(this, m_cellNode.get())),
      m_formulaProxy(XLFormulaProxy(this, m_cellNode.get()))
{}

/**
 * @details
 */
XLCell::~XLCell() = default;

/**
 * @details
 */
XLCell& XLCell::operator=(const XLCell& other)
{
    if (&other != this) {
        XLCell temp = other;
        std::swap(*this, temp);
    }

    return *this;
}

/**
 * @details
 */
XLCell& XLCell::operator=(XLCell&& other) noexcept
{
    if (&other != this) {
        m_cellNode      = std::move(other.m_cellNode);
        m_sharedStrings = other.m_sharedStrings;
        m_valueProxy    = XLCellValueProxy(this, m_cellNode.get());
        m_formulaProxy  = XLFormulaProxy(this, m_cellNode.get());    // pull request #160
    }

    return *this;
}

/**
 * @details
 */
void XLCell::copyFrom(XLCell const& other)
{
    using namespace std::literals::string_literals;
    if (!m_cellNode) {
        // copyFrom invoked by empty XLCell: create a new cell with reference & m_cellNode from other
        m_cellNode      = std::make_unique<XMLNode>(*other.m_cellNode);
        m_sharedStrings = other.m_sharedStrings;
        m_valueProxy    = XLCellValueProxy(this, m_cellNode.get());
        m_formulaProxy  = XLFormulaProxy(this, m_cellNode.get());
        return;
    }

    if ((&other != this) && (*other.m_cellNode == *m_cellNode))    // nothing to do
        return;

    // ===== If m_cellNode points to a different XML node than other
    if ((&other != this) && (*other.m_cellNode != *m_cellNode)) {
        m_cellNode->remove_children();

        // ===== Copy all XML child nodes
        for (XMLNode child = other.m_cellNode->first_child(); not child.empty(); child = child.next_sibling()) m_cellNode->append_copy(child);

        // ===== Delete all XML attributes that are not the cell reference ("r")
        // ===== 2024-07-26 BUGFIX: for-loop was invalidating loop variable with remove_attribute(attr) before advancing to next element
        XMLAttribute attr = m_cellNode->first_attribute();
        while (not attr.empty()) {
            XMLAttribute nextAttr = attr.next_attribute();                         // get a handle on next attribute before potentially removing attr
            if (strcmp(attr.name(), "r") != 0) m_cellNode->remove_attribute(attr); // remove all but the cell reference
            attr = nextAttr;                                                       // advance to previously stored next attribute
        }
        // ===== Copy all XML attributes that are not the cell reference ("r")
        for (auto attr = other.m_cellNode->first_attribute(); not attr.empty(); attr = attr.next_attribute())
            if (strcmp(attr.name(), "r") != 0) m_cellNode->append_copy(attr);
    }
}

/**
 * @details
 */
bool XLCell::empty() const { return (!m_cellNode) || m_cellNode->empty(); }

/**
 * @details
 * @todo 2024-08-10 TBD whether body can be replaced with !empty() (performance?)
 */
XLCell::operator bool() const { return m_cellNode && (not m_cellNode->empty() ); } // ===== 2024-05-28: replaced explicit bool evaluation

/**
 * @details This function returns a const reference to the cellReference property.
 */
XLCellReference XLCell::cellReference() const { return XLCellReference { m_cellNode->attribute("r").value() }; }

/**
 * @details This function returns a const reference to the cell reference by the offset from the current one.
 */
XLCell XLCell::offset(uint16_t rowOffset, uint16_t colOffset) const
{
    const XLCellReference offsetRef(cellReference().row() + rowOffset, cellReference().column() + colOffset);
    const auto            rownode  = getRowNode(m_cellNode->parent().parent(), offsetRef.row());
    const auto            cellnode = getCellNode(rownode, offsetRef.column());
    return XLCell { cellnode, m_sharedStrings };
}

/**
 * @details
 */
bool XLCell::hasFormula() const
{
    return (not m_cellNode->child("f").empty());    // evaluate child XMLNode as boolean
}

/**
 * @details
 */
XLFormulaProxy& XLCell::formula() { return m_formulaProxy; }

/**
* @details get the value of the s attribute of the cell node
*/
size_t XLCell::cellFormat() const { return m_cellNode->attribute("s").as_uint(0); }

/**
* @details set the s attribute of the cell node, pointing to an xl/styles.xml cellXfs index
*          the attribute will be created if not existant, function will fail if attribute creation fails
*/
bool XLCell::setCellFormat(size_t cellFormatIndex)
{
    XMLAttribute attr = m_cellNode->attribute("s");
    if (attr.empty() && not m_cellNode->empty())
        attr = m_cellNode->append_attribute("s");
    attr.set_value(cellFormatIndex); // silently fails on empty attribute, which is intended here
    return attr.empty() == false;
}

/**
 * @details
 */
void XLCell::print(std::basic_ostream<char>& ostr) const { m_cellNode->print(ostr); }

/**
 * @details
 */
XLCellAssignable::XLCellAssignable (XLCell const & other) : XLCell(other) {}

/**
 * @details
 */
XLCellAssignable::XLCellAssignable (XLCell && other) : XLCell(std::move(other)) {}

/**
 * @details
 */
XLCellAssignable& XLCellAssignable::operator=(const XLCell& other)
{
    copyFrom(other);
    return *this;
}

/**
 * @details
 */
XLCellAssignable& XLCellAssignable::operator=(const XLCellAssignable& other)
{
    copyFrom(other);
    return *this;
}

/**
 * @details
 */
XLCellAssignable& XLCellAssignable::operator=(XLCell&& other) noexcept
{
    copyFrom(other);
    return *this;
}

/**
 * @details
 */
XLCellAssignable& XLCellAssignable::operator=(XLCellAssignable&& other) noexcept
{
    copyFrom(other);
    return *this;
}

/**
 * @details
 */
const XLFormulaProxy& XLCell::formula() const { return m_formulaProxy; }

/**
 * @details clear cell contents except for those identified by keep
 */
void  XLCell::clear(uint32_t keep)
{
    // ===== Clear attributes
    XMLAttribute attr = m_cellNode->first_attribute();
    while (not attr.empty()) {
        XMLAttribute nextAttr = attr.next_attribute();
        std::string attrName = attr.name();
        if ((attrName == "r")                                 // if this is cell reference (must always remain untouched)
          ||((keep & XLKeepCellStyle) && attrName == "s")     // or style shall be kept & this is style
          ||((keep & XLKeepCellType ) && attrName == "t"))    // or type shall be kept & this is type
            attr = XMLAttribute{};                                // empty attribute won't get deleted
        // ===== Remove all non-kept attributes
        if (not attr.empty()) m_cellNode->remove_attribute(attr);
        attr = nextAttr; // advance to previously determined next cell node attribute
    }

    // ===== Clear node children
    XMLNode node = m_cellNode->first_child();
    while (not node.empty()) {
        XMLNode nextNode = node.next_sibling();
        // ===== Only preserve non-whitespace nodes
        if (node.type() == pugi::node_element) {
            std::string nodeName = node.name();
            if (((keep & XLKeepCellValue  ) && nodeName == "v")     // if value shall be kept & this is value
              ||((keep & XLKeepCellFormula) && nodeName == "f"))    // or formula shall be kept & this is formula
                node = XMLNode{};                                       // empty node won't get deleted
        }
        // ===== Remove all non-kept cell node children
        if (not node.empty()) m_cellNode->remove_child(node);
        node = nextNode; // advance to previously determined next cell node child
    }
}

/**
 * @pre
 * @post
 */
XLCellValueProxy& XLCell::value() { return m_valueProxy; }

/**
 * @details
 * @pre
 * @post
 */
const XLCellValueProxy& XLCell::value() const { return m_valueProxy; }

/**
 * @details
 * @pre
 * @post
 */
bool XLCell::isEqual(const XLCell& lhs, const XLCell& rhs) { return *lhs.m_cellNode == *rhs.m_cellNode; }

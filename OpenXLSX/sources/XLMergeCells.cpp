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
#include <algorithm>
#include <iostream>
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLMergeCells.hpp"
#include "XLCellReference.hpp"

#include <XLException.hpp>

using namespace OpenXLSX;

/**
 * @details Constructs a new XLMergeCells object. Invoked by XLWorksheet::mergeCells / ::unmergeCells
 * @note Unfortunately, there is no easy way to persist the reference cache, this could be optimized - however, references access shouldn't
 *       be much of a performance issue
 */
XLMergeCells::XLMergeCells(const XMLNode& node) : m_mergeCellsNode(std::make_unique<XMLNode>(node))
{
    if (m_mergeCellsNode->empty())
        throw XLInternalError("XLMergeCells constructor: can not construct with an empty XML node");

    XMLNode mergeNode = m_mergeCellsNode->first_child_of_type(pugi::node_element);
    while (not mergeNode.empty()) {
        bool invalidNode = true;

        // ===== For valid mergeCell nodes, add the reference to the reference cache
        if (std::string(mergeNode.name()) == "mergeCell") {
            std::string ref = mergeNode.attribute("ref").value();
            if (ref.length() > 0) {
                m_referenceCache.emplace_back(ref);
                invalidNode = false;
            }
        }

        // ===== Determine next element mergeNode
        XMLNode nextNode = mergeNode.next_sibling_of_type(pugi::node_element);

        // ===== In case of an invalid XML element: print an error and remove it from the XML, including whitespaces to the next sibling
        if (invalidNode) { // if mergeNode is not named mergeCell or does not have a valid ref attribute: remove it from the XML
            std::cerr << "XLMergeCells constructor: invalid child element, either name is not mergeCell or reference is invalid:" << std::endl;
            mergeNode.print(std::cerr);
            if (not nextNode.empty()) {
                // delete whitespaces between mergeNode and nextNode
                while (mergeNode.next_sibling() != nextNode) m_mergeCellsNode->remove_child(mergeNode.next_sibling());
            }
            m_mergeCellsNode->remove_child(mergeNode);
        }

        // ===== Advance to next element mergeNode
        mergeNode = nextNode;
    }
}

/**
 * @details
 */
XLMergeCells::~XLMergeCells() = default;

namespace { // anonymous namespace: do not export any symbols from here
    /**
     * @brief Test if (range) reference overlaps with the cell window defined by topRow, firstCol, bottomRow, lastCol
     * @return true in case of overlap, false if no overlap
     */
    bool XLReferenceOverlaps(std::string reference, uint32_t topRow, uint16_t firstCol, uint32_t bottomRow, uint16_t lastCol)
    {
        using namespace std::literals::string_literals;

        size_t pos = reference.find_first_of(':'); // find split mark between top left and bottom right cell
        if (pos < 2 || pos + 2 >= reference.length()) // range reference must have at least 2 characters before and after the colon
            throw XLInputError("XLMergeCells::"s + __func__ + ": not a valid range reference: \""s + reference + "\""s);
        XLCellReference refTL(reference.substr(0, pos));  // get top left cell reference
        XLCellReference refBR(reference.substr(pos + 1)); // get bottom right cell reference

        uint32_t refTopRow    = refTL.row();
        uint16_t refFirstCol  = refTL.column();
        uint32_t refBottomRow = refBR.row();
        uint16_t refLastCol   = refBR.column();
        if (refBottomRow < refTopRow || refLastCol < refFirstCol || (refBottomRow == refTopRow && refLastCol == refFirstCol))
            throw XLInputError("XLMergeCells::"s + __func__ + ": not a valid range reference: \""s + reference + "\""s);

// std::cout << __func__ << ":" << " reference is " << reference
//    << " refTopRow is " << refTopRow << " refBottomRow is " << refBottomRow << " refFirstCol is " << refFirstCol << " refLastCol is " << refLastCol
//    << " topRow is " << topRow << " bottomRow is " << bottomRow << " firstCol is " << firstCol << " lastCol is " << lastCol
//    << std::endl;

        // overlap
        if (refTopRow <= bottomRow && refBottomRow >= topRow        // vertical overlap
            && refFirstCol <= lastCol && refLastCol >= firstCol)    // horizontal overlap
            return true;
        return false; // otherwise: no overlap
    }
} // anonymous namespace

/**
 * @details Look up a merge index by the reference. If the reference does not exist, the returned index is -1.
 */
int32_t XLMergeCells::findMerge(const std::string& reference) const
{
    const auto iter = std::find_if(m_referenceCache.begin(), m_referenceCache.end(), [&](const std::string& ref) { return reference == ref; });

    return iter == m_referenceCache.end() ? -1 : static_cast<int32_t>(std::distance(m_referenceCache.begin(), iter));
}

/**
 * @details
 */
bool XLMergeCells::mergeExists(const std::string& reference) const { return findMerge(reference) >= 0; }

/**
 * @details Find the index of the merge of which cellRef is a part. If no such merge exists, the returned index is -1.
 */
int32_t XLMergeCells::findMergeByCell(const std::string& cellRef) const { return findMergeByCell(XLCellReference(cellRef)); }
int32_t XLMergeCells::findMergeByCell(XLCellReference cellRef) const
{
    const auto iter = std::find_if(m_referenceCache.begin(), m_referenceCache.end(),
    /**/                           [&](const std::string& ref) { // use XLReferenceOverlaps with a "range" that only contains cellRef
    /**/                               return XLReferenceOverlaps( ref, cellRef.row(), cellRef.column(), cellRef.row(), cellRef.column());
    /**/                           });

    return iter == m_referenceCache.end() ? -1 : static_cast<int32_t>(std::distance(m_referenceCache.begin(), iter));
}

/**
 * @details
 */
size_t XLMergeCells::count() const { return m_referenceCache.size(); }

/**
 * @details
 */
const char* XLMergeCells::merge(int32_t index) const
{
    if (index < 0 || static_cast<uint32_t>(index) >= m_referenceCache.size()) {
        using namespace std::literals::string_literals;
        throw XLInputError("XLMergeCells::"s + __func__ + ": index "s + std::to_string(index) + " is out of range"s);
    }
    return m_referenceCache[index].c_str();
}

/**
 * @details Append a mergeCell by creating a new node in the XML file and adding the string to it. The index to the
 * appended merge is returned
 * Before appending a mergeCell entry with reference, check that reference does not overlap with any existing references
 */
int32_t XLMergeCells::appendMerge(const std::string& reference)
{
    using namespace std::literals::string_literals;

    size_t referenceCacheSize = m_referenceCache.size();
    if (referenceCacheSize >= XLMaxMergeCells)
        throw XLInputError("XLMergeCells::"s + __func__ + ": exceeded max merge cells count "s + std::to_string(XLMaxMergeCells));

    size_t pos = reference.find_first_of(':'); // find split mark between top left and bottom right cell
    if (pos < 2 || pos + 2 >= reference.length()) // range reference must have at least 2 characters before and after the colon
        throw XLInputError("XLMergeCells::"s + __func__ + ": not a valid range reference: \""s + reference + "\""s);
    XLCellReference refTL(reference.substr(0, pos));  // get top left cell reference
    XLCellReference refBR(reference.substr(pos + 1)); // get bottom right cell reference

    uint32_t refTopRow    = refTL.row();
    uint16_t refFirstCol  = refTL.column();
    uint32_t refBottomRow = refBR.row();
    uint16_t refLastCol   = refBR.column();
    if (refBottomRow < refTopRow || refLastCol < refFirstCol || (refBottomRow == refTopRow && refLastCol == refFirstCol))
        throw XLInputError("XLMergeCells::"s + __func__ + ": not a valid range reference: \""s + reference + "\""s);

    for (std::string ref : m_referenceCache) {
        if (XLReferenceOverlaps(ref, refTopRow, refFirstCol, refBottomRow, refLastCol))
            throw XLInputError("XLMergeCells::"s + __func__ + ": reference \""s + reference
            /**/                   + "\" overlaps with existing reference \""s + ref + "\""s);
    }
    // if execution gets here: no overlaps

    // append new mergeCell element and set attribute ref
    XMLNode insertAfter = m_mergeCellsNode->last_child_of_type(pugi::node_element);
    XMLNode newMerge{};
    if (insertAfter.empty()) newMerge = m_mergeCellsNode->prepend_child("mergeCell");
    else                     newMerge = m_mergeCellsNode->insert_child_after("mergeCell", insertAfter);
    if (newMerge.empty())
        throw XLInternalError("XLMergeCells::"s + __func__ + ": failed to insert reference: \""s + reference + "\""s);
    newMerge.append_attribute("ref").set_value(reference.c_str());

    m_referenceCache.emplace_back(newMerge.attribute("ref").value()); // index of this element = previous referenceCacheSize

    // ===== Update the array count attribute
    XMLAttribute attr = m_mergeCellsNode->attribute("count");
    if (attr.empty()) attr = m_mergeCellsNode->append_attribute("count");
    attr.set_value(m_referenceCache.size());

    return static_cast<int32_t>(referenceCacheSize);
}

/**
 * @details Delete the merge at the given index
 */
void XLMergeCells::deleteMerge(int32_t index)
{
    using namespace std::literals::string_literals;

    if (index < 0 || static_cast<uint32_t>(index) >= m_referenceCache.size())
        throw XLInputError("XLMergeCells::"s + __func__ + ": index "s + std::to_string(index) + " is out of range"s);

    int32_t curIndex = 0;
    XMLNode node = m_mergeCellsNode->first_child_of_type(pugi::node_element);
    while(curIndex < index && not node.empty()) {
        node = node.next_sibling_of_type(pugi::node_element);
		  ++curIndex;
    }
    if (node.empty())
        throw XLInternalError("XLMergeCells::"s + __func__ + ": mismatch between size of mergeCells XML node and internal reference cache"s);

    // ===== node was found: delete preceeding whitespace nodes and the node itself
    while (node.previous_sibling().type() == pugi::node_pcdata) m_mergeCellsNode->remove_child(node.previous_sibling());
    m_mergeCellsNode->remove_child(node);

    m_referenceCache.erase(m_referenceCache.begin() + curIndex);

    // ===== Update the array count attribute
    XMLAttribute attr = m_mergeCellsNode->attribute("count");
    if (attr.empty()) attr = m_mergeCellsNode->append_attribute("count");
    attr.set_value(m_referenceCache.size()); // update the array count attribute
}

/**
 * @details Print the underlying XML using pugixml::xml_node::print
 */
void XLMergeCells::print(std::basic_ostream<char>& ostr) const { m_mergeCellsNode->print(ostr); }

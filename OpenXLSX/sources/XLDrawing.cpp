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
// #include <algorithm>
#include <cctype>       // std::isdigit (issue #330)
#include <string>       // std::string
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"               // pugi_parse_settings
#include "XLDrawing.hpp"
#include "utilities/XLUtilities.hpp"    // OpenXLSX::ignore, appendAndGetNode

using namespace OpenXLSX;

namespace OpenXLSX
{
    const std::string ShapeNodeName = "v:shape";
    const std::string ShapeTypeNodeName = "v:shapetype";

    // utility functions

    /**
     * @brief test if v:shapetype attribute id already exists at begin of file
     * @param rootNode the parent that holds all v:shapetype nodes to be compared
     * @param shapeTypeNode the node to check for a duplicate id
     * @return true if node would be a duplicate (and can be deleted), false if node is unique (so far)
     */
    bool wouldBeDuplicateShapeType(XMLNode const& rootNode, XMLNode const& shapeTypeNode)
    {
        std::string id = shapeTypeNode.attribute("id").value();
        XMLNode node = rootNode.first_child_of_type(pugi::node_element);
        using namespace std::literals::string_literals;
        while (node != shapeTypeNode && node.raw_name() == ShapeTypeNodeName) { // do not compare shapeTypeNode with itself, and abort on first non-shapetype node
           if (node.attribute("id").value() == id) return true;                // found an existing shapetype with the same id
           node = node.next_sibling_of_type(pugi::node_element);
        }
        return false;
    }

    /**
     * @brief move an xml node to a new position within the same root node
     * @param rootNode the parent that can perform inserts / deletes
     * @param node the node to move
     * @param insertAfter the reference for the new location, if insertAfter.empty(): move node to first position in rootNode
     * @param withWhitespaces if true, move all node_pcdata nodes preceeding node, together with the node
     * @return the node as inserted at the destination - empty node if insertion can't be performed (rootNode or node are empty)
     */
    XMLNode moveNode(XMLNode& rootNode, XMLNode& node, XMLNode const& insertAfter, bool withWhitespaces = true)
    {
        if (rootNode.empty() || node.empty()) return XMLNode{}; // can't perform move
        if (node == insertAfter || node == insertAfter.next_sibling()) return node; // nothing to do

        XMLNode newNode{};
        if (insertAfter.empty())
            newNode = rootNode.prepend_copy(node);
        else
            newNode = rootNode.insert_copy_after(node, insertAfter);

        if (withWhitespaces) {
            copyLeadingWhitespaces(rootNode, node, newNode);            // copy leading whitespaces
            while (node.previous_sibling().type() == pugi::node_pcdata) // then remove whitespaces that were just copied to newNode
                rootNode.remove_child(node.previous_sibling());
        }
        rootNode.remove_child(node);                                    // remove node that was just copied to new location

        return newNode;
    }

    XLShapeTextVAlign XLShapeTextVAlignFromString(std::string const& vAlignString)
    {
        if (vAlignString == "Center") return XLShapeTextVAlign::Center;
        if (vAlignString == "Top"   ) return XLShapeTextVAlign::Top;
        if (vAlignString == "Bottom") return XLShapeTextVAlign::Bottom;
        std::cerr << __func__ << ": invalid XLShapeTextVAlign setting" << vAlignString << std::endl;
        return XLShapeTextVAlign::Invalid;

    }
    std::string XLShapeTextVAlignToString(XLShapeTextVAlign vAlign)
    {
        switch (vAlign) {
            case XLShapeTextVAlign::Center:  return "Center";
            case XLShapeTextVAlign::Top:     return "Top";
            case XLShapeTextVAlign::Bottom:  return "Bottom";
            case XLShapeTextVAlign::Invalid: [[fallthrough]];
            default:                                return "(invalid)";
        }
    }
    XLShapeTextHAlign XLShapeTextHAlignFromString(std::string const& hAlignString)
    {
        if (hAlignString == "Left"  ) return XLShapeTextHAlign::Left;
        if (hAlignString == "Right" ) return XLShapeTextHAlign::Right;
        if (hAlignString == "Center") return XLShapeTextHAlign::Center;
        std::cerr << __func__ << ": invalid XLShapeTextHAlign setting" << hAlignString << std::endl;
        return XLShapeTextHAlign::Invalid;
    }
    std::string XLShapeTextHAlignToString(XLShapeTextHAlign hAlign)
    {
        switch (hAlign) {
            case XLShapeTextHAlign::Left:    return "Left";
            case XLShapeTextHAlign::Right:   return "Right";
            case XLShapeTextHAlign::Center:  return "Center";
            case XLShapeTextHAlign::Invalid: [[fallthrough]];
            default:                                return "(invalid)";
        }
    }

}    // namespace OpenXLSX


// ========== XLShapeClientData Member Functions

/**
 * @details default constructor
 */
XLShapeClientData::XLShapeClientData()
 : m_clientDataNode(std::make_unique<XMLNode>())
{}

/**
* @details XMLNode constructor
*/
XLShapeClientData::XLShapeClientData(const XMLNode& node)
 : m_clientDataNode(std::make_unique<XMLNode>(node))
{}


/**
 * @details XLShapeClientData getter functions
 */
bool elementTextAsBool(XMLNode const& node) {
    if (node.empty()) return false;          // node doesn't exist: false
    if (node.text().empty()) return true;    // node exists but has no setting: true
    char c = node.text().get()[0];           // node exists and has a setting:
    if (c == 't' || c == 'T') return true;   //   check true setting on first letter only
    return false;                            //   otherwise return false
}

std::string       XLShapeClientData::objectType()    const { return appendAndGetAttribute(*m_clientDataNode, "ObjectType", "Note"    ).value()  ; }
bool              XLShapeClientData::moveWithCells() const { return elementTextAsBool    ( m_clientDataNode->child("x:MoveWithCells"))          ; }
bool              XLShapeClientData::sizeWithCells() const { return elementTextAsBool    ( m_clientDataNode->child("x:SizeWithCells"))          ; }
std::string       XLShapeClientData::anchor()        const { return                        m_clientDataNode->child("x:Anchor").value()          ; }
bool              XLShapeClientData::autoFill()      const { return elementTextAsBool    ( m_clientDataNode->child("x:AutoFill"))               ; }
XLShapeTextVAlign XLShapeClientData::textVAlign()    const {
   return XLShapeTextVAlignFromString(m_clientDataNode->child("x:TextVAlign").text().get());
   // XMLNode vAlign = m_clientDataNode->child("x:TextVAlign");
   // if (vAlign.empty()) return XLDefaultShapeTextVAlign;
   // return XLShapeTextVAlignFromString(vAlign.text().get());
}
XLShapeTextHAlign XLShapeClientData::textHAlign()    const {
   return XLShapeTextHAlignFromString(m_clientDataNode->child("x:TextHAlign").text().get());
   // XMLNode hAlign = m_clientDataNode->child("x:TextHAlign");
   // if (hAlign.empty()) return XLDefaultShapeTextHAlign;
   // return XLShapeTextHAlignFromString(hAlign.text().get());
}
uint32_t XLShapeClientData::row()                    const {
    return appendAndGetNode(*m_clientDataNode, "x:Row",    m_nodeOrder, XLForceNamespace).text().as_uint(0);
}
uint16_t XLShapeClientData::column()                 const {
    return appendAndGetNode(*m_clientDataNode, "x:Column", m_nodeOrder, XLForceNamespace).text().as_uint(0);
}

/**
 * @details XLShapeClientData setter functions
 */
bool XLShapeClientData::setObjectType(std::string newObjectType) { return appendAndSetAttribute(*m_clientDataNode, "ObjectType", newObjectType ).empty() == false; }
bool XLShapeClientData::setMoveWithCells(bool set)
 { return appendAndGetNode(*m_clientDataNode, "x:MoveWithCells", m_nodeOrder, XLForceNamespace).text().set(set ? "" : "False");
    // XMLNode moveWithCellsNode = appendAndGetNode(*m_clientDataNode, "x:MoveWithCells", m_nodeOrder, XLForceNamespace);
    // if (moveWithCellsNode.empty()) return false;
    // return moveWithCellsNode.text().set(set ? "True" : "False");
}
bool XLShapeClientData::setSizeWithCells(bool set)
 { return appendAndGetNode(*m_clientDataNode, "x:SizeWithCells", m_nodeOrder, XLForceNamespace).text().set(set ? "" : "False"); }

bool XLShapeClientData::setAnchor       (std::string newAnchor)
 { return appendAndGetNode(*m_clientDataNode, "x:Anchor",        m_nodeOrder, XLForceNamespace).text().set(newAnchor.c_str()); }

bool XLShapeClientData::setAutoFill     (bool set)
 { return appendAndGetNode(*m_clientDataNode, "x:AutoFill",      m_nodeOrder, XLForceNamespace).text().set(set ? "True" : "False"); }

bool XLShapeClientData::setTextVAlign   (XLShapeTextVAlign newTextVAlign)
 { return appendAndGetNode(*m_clientDataNode, "x:TextVAlign",    m_nodeOrder, XLForceNamespace).text().set(XLShapeTextVAlignToString(newTextVAlign).c_str()); }

bool XLShapeClientData::setTextHAlign   (XLShapeTextHAlign newTextHAlign)
 { return appendAndGetNode(*m_clientDataNode, "x:TextHAlign",    m_nodeOrder, XLForceNamespace).text().set(XLShapeTextHAlignToString(newTextHAlign).c_str()); }

bool XLShapeClientData::setRow          (uint32_t newRow)
 { return appendAndGetNode(*m_clientDataNode, "x:Row",           m_nodeOrder, XLForceNamespace).text().set(newRow); }

bool XLShapeClientData::setColumn       (uint16_t newColumn)
 { return appendAndGetNode(*m_clientDataNode, "x:Column",        m_nodeOrder, XLForceNamespace).text().set(newColumn); }


// ========== XLShape Member Functions
/**
 * @details
 */
XLShapeStyle::XLShapeStyle()
 : m_style(""),
   m_styleAttribute(std::make_unique<XMLAttribute>())
{
    setPosition("absolute");
    setMarginLeft(100);
    setMarginTop(0);
    setWidth(80);
    setHeight(50);
    setMsoWrapStyle("none");
    setVTextAnchor("middle");
    hide();
}

/**
 * @details
 */
XLShapeStyle::XLShapeStyle(const std::string& styleAttribute)
 : m_style(styleAttribute),
   m_styleAttribute(std::make_unique<XMLAttribute>())
{}

/**
 * @details
 */
XLShapeStyle::XLShapeStyle(const XMLAttribute& styleAttribute)
 : m_style(""),
   m_styleAttribute(std::make_unique<XMLAttribute>(styleAttribute))
{}

/**
 * @details find attributeName in m_nodeOrder, then return index
 */
int16_t XLShapeStyle::attributeOrderIndex(std::string const& attributeName) const
{
// std::string m_style{"position:absolute;margin-left:100pt;margin-top:0pt;width:50pt;height:50.0pt;mso-wrap-style:none;v-text-anchor:middle;visibility:hidden"};
    auto attributeIterator = std::find(m_nodeOrder.begin(), m_nodeOrder.end(), attributeName);
    if (attributeIterator == m_nodeOrder.end())
        return -1;
    return attributeIterator - m_nodeOrder.begin();
}

/**
 * @details
 */
XLShapeStyleAttribute XLShapeStyle::getAttribute(std::string const& attributeName, std::string const& valIfNotFound) const
{
    if (attributeOrderIndex(attributeName) == -1) {
        using namespace std::literals::string_literals;
        throw XLInternalError("XLShapeStyle.getAttribute: attribute "s + attributeName + " is not defined in class"s);
    }

    // if attribute is linked, re-read m_style each time in case the underlying XML has changed
    if (not m_styleAttribute->empty()) m_style = std::string(m_styleAttribute->value());

    size_t lastPos = 0;
    XLShapeStyleAttribute result;
    result.name = "";   // indicates "not found"
    result.value = "";  // default in case attribute name is found but has no value

    do {
        size_t pos = m_style.find(';', lastPos);
        std::string attrPair = m_style.substr(lastPos, pos - lastPos);
        if (attrPair.length() > 0) {
            size_t separatorPos = attrPair.find(':');
            if (attributeName == attrPair.substr(0, separatorPos)) { // found!
                result.name = attributeName;
                if( separatorPos != std::string::npos )
                    result.value = attrPair.substr(separatorPos + 1);
    break;
            }
        }
        lastPos = pos+1;
    } while (lastPos < m_style.length());
    if (lastPos >= m_style.length())   // attribute was not found
        result.value = valIfNotFound;     // -> return default value
    return result;
}

/**
 * @details
 */
bool XLShapeStyle::setAttribute(std::string const& attributeName, std::string const& attributeValue)
{
    int16_t attrIndex = attributeOrderIndex(attributeName);
    if (attrIndex == -1) {
        using namespace std::literals::string_literals;
        throw XLInternalError("XLShapeStyle.setAttribute: attribute "s + attributeName + " is not defined in class"s);
    }

    // if attribute is linked, re-read m_style each time in case the underlying XML has changed
    if (not m_styleAttribute->empty()) m_style = std::string(m_styleAttribute->value());

    size_t lastPos = 0;
    size_t appendPos = 0;
    do {
        size_t pos = m_style.find(';', lastPos);

        std::string attrPair = m_style.substr(lastPos, pos - lastPos);
        if (attrPair.length() > 0) {
            size_t separatorPos = attrPair.find(':');
            int16_t thisAttrIndex = attributeOrderIndex(attrPair.substr(0, separatorPos));
            if (thisAttrIndex >= attrIndex) { // can insert or update
                appendPos = (thisAttrIndex == attrIndex ? pos : lastPos); // if match: append from following attribute, if not found, append from current (insert)
    break;
            }
        }
        if(pos == std::string::npos)
            lastPos = pos;
        else
            lastPos = pos + 1;
    } while (lastPos < m_style.length());
    if (lastPos >= m_style.length() || appendPos > m_style.length()) { // if attribute needs to be appended or was found at last position
        appendPos = m_style.length();                                     // then nothing remains from m_style to follow (appended) attribute

        // for semi-colon logic, lastPos needs to be capped at string length (so that it can be 0 for empty string)
        if (lastPos > m_style.length()) lastPos = m_style.length();
    }

    using namespace std::literals::string_literals;
    m_style = m_style.substr(0, lastPos) + ((lastPos != 0 && appendPos == m_style.length()) ? ";"s : ""s)   // prepend ';' if attribute is appended to non-empty string
    /**/      + attributeName + ":"s + attributeValue                                                       // insert attribute:value pair
    /**/      + (appendPos < m_style.length() ? ";"s : ""s) + m_style.substr(appendPos);                    // append ';' if attribute is inserted before other data

    // if attribute is linked, update it with the new style value
    if (not m_styleAttribute->empty()) m_styleAttribute->set_value(m_style.c_str());

    return true;
}

/**
 * @details XLShapeStyle getter functions
 */
std::string XLShapeStyle::position    () const { return                                 getAttribute("position"      ).value  ; }
uint16_t    XLShapeStyle::marginLeft  () const { return static_cast<uint16_t>(std::stoi(getAttribute("margin-left"   ).value)); }
uint16_t    XLShapeStyle::marginTop   () const { return static_cast<uint16_t>(std::stoi(getAttribute("margin-top"    ).value)); }
uint16_t    XLShapeStyle::width       () const { return static_cast<uint16_t>(std::stoi(getAttribute("width"         ).value)); }
uint16_t    XLShapeStyle::height      () const { return static_cast<uint16_t>(std::stoi(getAttribute("height"        ).value)); }
std::string XLShapeStyle::msoWrapStyle() const { return                                 getAttribute("mso-wrap-style").value  ; }
std::string XLShapeStyle::vTextAnchor () const { return                                 getAttribute("v-text-anchor" ).value  ; }
bool        XLShapeStyle::hidden      () const { return                    ("hidden" == getAttribute("visibility"    ).value ); }
bool        XLShapeStyle::visible     () const { return !hidden(); }

/**
 * @details XLShapeStyle setter functions
 */
bool XLShapeStyle::setPosition    (std::string newPosition)     { return setAttribute("position",                      newPosition                       ); }
bool XLShapeStyle::setMarginLeft  (uint16_t newMarginLeft)      { return setAttribute("margin-left",    std::to_string(newMarginLeft) + std::string("pt")); }
bool XLShapeStyle::setMarginTop   (uint16_t newMarginTop)       { return setAttribute("margin-top",     std::to_string(newMarginTop)  + std::string("pt")); }
bool XLShapeStyle::setWidth       (uint16_t newWidth)           { return setAttribute("width",          std::to_string(newWidth)      + std::string("pt")); }
bool XLShapeStyle::setHeight      (uint16_t newHeight)          { return setAttribute("height",         std::to_string(newHeight)     + std::string("pt")); }
bool XLShapeStyle::setMsoWrapStyle(std::string newMsoWrapStyle) { return setAttribute("mso-wrap-style",                newMsoWrapStyle                   ); }
bool XLShapeStyle::setVTextAnchor (std::string newVTextAnchor)  { return setAttribute("v-text-anchor",                 newVTextAnchor                    ); }
bool XLShapeStyle::hide           ()                            { return setAttribute("visibility",                    "hidden"                          ); }
bool XLShapeStyle::show           ()                            { return setAttribute("visibility",                    "visible"                         ); }

// ========== XLShape Member Functions
XLShape::XLShape() : m_shapeNode(std::make_unique<XMLNode>(XMLNode())) {}

XLShape::XLShape(const XMLNode& node) : m_shapeNode(std::make_unique<XMLNode>(node)) {}

/**
 * @details getter functions: return the shape's attributes
 */
std::string  XLShape::shapeId()     const { return                                     m_shapeNode->attribute("id").value()               ; } // const for user, managed by parent
std::string  XLShape::fillColor()   const { return              appendAndGetAttribute(*m_shapeNode, "fillcolor",     "#ffffc0").value()   ; }
bool         XLShape::stroked()     const { return              appendAndGetAttribute(*m_shapeNode, "stroked",       "t"      ).as_bool() ; }
std::string  XLShape::type()        const { return              appendAndGetAttribute(*m_shapeNode, "type",          ""       ).value()   ; }
bool         XLShape::allowInCell() const { return              appendAndGetAttribute(*m_shapeNode, "o:allowincell", "f"      ).as_bool() ; }
XLShapeStyle XLShape::style()             { return XLShapeStyle(appendAndGetAttribute(*m_shapeNode, "style",         ""       )          ); }

XLShapeClientData XLShape::clientData() {
    return XLShapeClientData(appendAndGetNode(*m_shapeNode, "x:ClientData", m_nodeOrder, XLForceNamespace));
}

/**
 * @details setter functions: assign the shape's attributes
 */
bool XLShape::setFillColor  (std::string const& newFillColor) { return appendAndSetAttribute(*m_shapeNode, "fillcolor",     newFillColor     ).empty() == false; }
bool XLShape::setStroked    (bool set)                        { return appendAndSetAttribute(*m_shapeNode, "stroked",       (set ? "t" : "f")).empty() == false; }
bool XLShape::setType       (std::string const& newType)      { return appendAndSetAttribute(*m_shapeNode, "type",          newType          ).empty() == false; }
bool XLShape::setAllowInCell(bool set)                        { return appendAndSetAttribute(*m_shapeNode, "o:allowincell", (set ? "t" : "f")).empty() == false; }
bool XLShape::setStyle      (std::string const& newStyle)     { return appendAndSetAttribute(*m_shapeNode, "style",         newStyle         ).empty() == false; }
bool XLShape::setStyle      (XLShapeStyle const& newStyle)    { return setStyle( newStyle.raw() ); }


// ========== XLVmlDrawing Member Functions

/**
 * @details The constructor creates an instance of the superclass, XLXmlFile
 */
XLVmlDrawing::XLVmlDrawing(XLXmlData* xmlData)
 : XLXmlFile(xmlData)
{
    if (xmlData->getXmlType() != XLContentType::VMLDrawing)
        throw XLInternalError("XLVmlDrawing constructor: Invalid XML data.");

    XMLDocument & doc = xmlDocument();
    if (doc.document_element().empty())   // handle a bad (no document element) drawing XML file
        doc.load_string(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
            "<xml"
            " xmlns:v=\"urn:schemas-microsoft-com:vml\""
            " xmlns:o=\"urn:schemas-microsoft-com:office:office\""
            " xmlns:x=\"urn:schemas-microsoft-com:office:excel\""
            " xmlns:w10=\"urn:schemas-microsoft-com:office:word\""
            ">"
            "\n</xml>",
            pugi_parse_settings
        );

    // ===== Re-sort the document: move all v:shapetype nodes to the beginning of the XML document element and eliminate duplicates
    // ===== Also: determine highest used shape id, regardless of basename (pattern [^0-9]*[0-9]*) and m_shapeCount
    using namespace std::literals::string_literals;

    XMLNode rootNode = doc.document_element();
    XMLNode node = rootNode.first_child_of_type(pugi::node_element);
    XMLNode lastShapeTypeNode{};
    while (not node.empty()) {
       XMLNode nextNode = node.next_sibling_of_type(pugi::node_element); // determine next node early because node may be invalidated by moveNode
       if (node.raw_name() == ShapeTypeNodeName) {
           if (wouldBeDuplicateShapeType(rootNode, node)) { // if shapetype attribute id already exists at begin of file
               while(node.previous_sibling().type() == pugi::node_pcdata) // delete preceeding whitespaces
                   rootNode.remove_child(node.previous_sibling());        //  ...
               rootNode.remove_child(node);                               // and the v:shapetype node, as it can not be referenced for lack of a unique id
           }
           else
               lastShapeTypeNode = moveNode(rootNode, node, lastShapeTypeNode); // move node to end of continuous list of shapetype nodes
       }
       else if (node.raw_name() == ShapeNodeName) {
          ++m_shapeCount;
          std::string strNodeId = node.attribute("id").value();
          size_t pos = strNodeId.length();
          uint32_t nodeId = 0;
          while (pos > 0 && std::isdigit(strNodeId[--pos])) // determine any trailing nodeId
              nodeId = 10*nodeId + strNodeId[pos] - '0';
          m_lastAssignedShapeId = std::max(m_lastAssignedShapeId, nodeId);
       }
       node = nextNode;
    }
    // Henceforth: assume that it is safe to consider shape nodes a continuous list (well - unless there are other node types as well)

    XMLNode shapeTypeNode{};
    if (not lastShapeTypeNode.empty()) {
        shapeTypeNode = rootNode.first_child_of_type(pugi::node_element);
        while (not shapeTypeNode.empty() && shapeTypeNode.raw_name() != ShapeTypeNodeName)
            shapeTypeNode = shapeTypeNode.next_sibling_of_type(pugi::node_element);
    }
    if (shapeTypeNode.empty()) {
        shapeTypeNode = rootNode.prepend_child(ShapeTypeNodeName.c_str(), XLForceNamespace);
        rootNode.prepend_child(pugi::node_pcdata).set_value("\n\t");
    }
    if (shapeTypeNode.first_child().empty())
        shapeTypeNode.append_child(pugi::node_pcdata).set_value("\n\t"); // insert indentation if node was empty

    // ===== Ensure that attributes exist for first shapetype node, default-initialize if needed:
    m_defaultShapeTypeId = appendAndGetAttribute(shapeTypeNode, "id",        "_x0000_t202").value();
    appendAndGetAttribute(shapeTypeNode, "coordsize", "21600,21600");
    appendAndGetAttribute(shapeTypeNode, "o:spt",     "202");
    appendAndGetAttribute(shapeTypeNode, "path",      "m,l,21600l21600,21600l21600,xe");

    XMLNode strokeNode = shapeTypeNode.child("v:stroke");
    if (strokeNode.empty()) {
        strokeNode = shapeTypeNode.prepend_child("v:stroke", XLForceNamespace);
        shapeTypeNode.prepend_child(pugi::node_pcdata).set_value("\n\t\t");
    }
    appendAndGetAttribute(strokeNode, "joinstyle", "miter");

    XMLNode pathNode = shapeTypeNode.child("v:path");
    if (pathNode.empty()) {
        pathNode = shapeTypeNode.insert_child_after("v:path", strokeNode, XLForceNamespace);
        copyLeadingWhitespaces(shapeTypeNode, strokeNode, pathNode);
    }
    appendAndGetAttribute(pathNode, "gradientshapeok", "t");
    appendAndGetAttribute(pathNode, "o:connecttype", "rect");
}


/**
 * @details get first shape node in document_element
 * @return first shape node, empty node if none found
 */
XMLNode XLVmlDrawing::firstShapeNode() const
{
    using namespace std::literals::string_literals;

    XMLNode node = xmlDocument().document_element().first_child_of_type(pugi::node_element);
    while (not node.empty() && node.raw_name() != ShapeNodeName)   // skip non shape nodes
        node = node.next_sibling_of_type(pugi::node_element);
    return node;
}

/**
 * @details get last shape node in document_element
 * @return last shape node, empty node if none found
 */
XMLNode XLVmlDrawing::lastShapeNode() const
{
    using namespace std::literals::string_literals;

    XMLNode node = xmlDocument().document_element().last_child_of_type(pugi::node_element);
    while (not node.empty() && node.raw_name() != ShapeNodeName)
        node = node.previous_sibling_of_type(pugi::node_element);
    return node;
}

/**
 * @details get last shape node at index in document_element
 * @return shape node at index - throws if index is out of bounds
 */
XMLNode XLVmlDrawing::shapeNode(uint32_t index) const
{
    using namespace std::literals::string_literals;

    XMLNode node{}; // scope declaration, ensures node.empty() when index >= m_shapeCount
    if (index < m_shapeCount) {
        uint16_t i = 0;
        node = firstShapeNode();
        while (i != index && not node.empty() && node.raw_name() == ShapeNodeName) {   // follow shape index
            ++i;
            node = node.next_sibling_of_type(pugi::node_element);
        }
    }
    if (node.empty() || node.raw_name() != ShapeNodeName)
        throw XLException("XLVmlDrawing: shape index "s + std::to_string(index) + " is out of bounds"s);

    return node;
}

/**
 * @details
 */
XMLNode XLVmlDrawing::shapeNode(std::string const& cellRef) const
{
    XLCellReference destRef(cellRef);
    uint32_t destRow = destRef.row() - 1;    // for accessing a shape: x:Row and x:Column are zero-indexed
    uint16_t destCol = destRef.column() - 1; // ..

    XMLNode node = firstShapeNode();
    while (not node.empty()) {
        if ((destRow == node.child("x:ClientData").child("x:Row").text().as_uint())
          &&(destCol == node.child("x:ClientData").child("x:Column").text().as_uint()))
    break; // found shape for cellRef

        do { // locate next shape node
            node = node.next_sibling_of_type(pugi::node_element);
        } while (not node.empty() && node.name() != ShapeNodeName);
    }
    return node;
}

/**
 * @details TODO: write doxygen headers for functions in this module
 */
uint32_t XLVmlDrawing::shapeCount() const { return m_shapeCount; }

/**
 * @details TODO: write doxygen headers for functions in this module
 */
XLShape XLVmlDrawing::shape(uint32_t index) const { return XLShape(shapeNode(index)); }

/**
 * @details TODO: write doxygen headers for functions in this module
 */
bool XLVmlDrawing::deleteShape(uint32_t index)
{
    XMLNode rootNode = xmlDocument().document_element();
    XMLNode node = shapeNode(index);   // returns a valid node or throws
    --m_shapeCount;                    // if shapeNode(index) did not throw: decrement shape count
    while (node.previous_sibling().type() == pugi::node_pcdata) // remove leading whitespaces
        rootNode.remove_child(node.previous_sibling());
    rootNode.remove_child(node);                                // then remove shape node itself

    return true;
}

/**
 * @details TODO: write doxygen headers for functions in this module
 */
bool XLVmlDrawing::deleteShape(std::string const& cellRef)
{
    XMLNode rootNode = xmlDocument().document_element();
    XMLNode node = shapeNode(cellRef);
    if (node.empty()) return false;    // nothing found to delete

    --m_shapeCount;                    // if shapeNode(cellRef) returned a non-empty node: decrement shape count
    while (node.previous_sibling().type() == pugi::node_pcdata) // remove leading whitespaces
        rootNode.remove_child(node.previous_sibling());
    rootNode.remove_child(node);                                // then remove shape node itself

    return true;
}

/**
 * @details insert shape and return index
 */
XLShape XLVmlDrawing::createShape([[maybe_unused]] const XLShape& shapeTemplate )
{

    XMLNode rootNode = xmlDocument().document_element();
    XMLNode node = lastShapeNode();
    if (node.empty()) {
        node = rootNode.last_child_of_type(pugi::node_element);
    }
    if (not node.empty()) {                                                      // default case: a previous element node exists
        node = rootNode.insert_child_after(ShapeNodeName.c_str(), node, XLForceNamespace);   // insert the node after the last shape node if any
        copyLeadingWhitespaces(rootNode, node.previous_sibling(), node);                     // copy whitespaces from node after which this one was just inserted
    }
    else {                                                                       // else: shouldn't happen - but just in case
        node = rootNode.prepend_child(ShapeNodeName.c_str(), XLForceNamespace);              // insert node prior to trailing whitespaces
        rootNode.prepend_child(pugi::node_pcdata).set_value("\n\t");                         // prefix new node with whitespaces
    }

    // ===== Assign a new shape id & account for it in m_lastAssignedShapeId
    using namespace std::literals::string_literals;
    node.prepend_attribute("id").set_value(("shape_"s + std::to_string( m_lastAssignedShapeId++ )).c_str());
    node.append_attribute("type").set_value(("#"s + m_defaultShapeTypeId).c_str());

    m_shapeCount++;
    return XLShape(node); // return object to manipulate new shape
}


/**
 * @details Print the underlying XML using pugixml::xml_node::print
 */
void XLVmlDrawing::print(std::basic_ostream<char>& ostr) const { xmlDocument().document_element().print( ostr ); }

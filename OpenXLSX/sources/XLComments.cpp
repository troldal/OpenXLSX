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
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"               // pugi_parse_settings
#include "XLComments.hpp"
#include "utilities/XLUtilities.hpp"    // OpenXLSX::ignore, appendAndGetNode

using namespace OpenXLSX;

namespace {
    // module-local utility functions
    /**
     * @details TODO: write doxygen headers for functions in this module
     */
    std::string getCommentString(XMLNode const & commentNode)
    {
        std::string result{};

        using namespace std::literals::string_literals;
        XMLNode textElement = commentNode.child("text").first_child_of_type(pugi::node_element);
        while (not textElement.empty()) {
            if (textElement.name() == "t"s) {
                result += textElement.first_child().value();
            }
            else if (textElement.name() == "r"s) { // rich text
                XMLNode richTextSubnode = textElement.first_child_of_type(pugi::node_element);
                while (not richTextSubnode.empty()) {
                    if (richTextSubnode.name() == "t"s) {
                        result += richTextSubnode.first_child().value();
                    }
                    else if (richTextSubnode.name() == "rPr"s) {} // ignore rich text formatting info
                    else {} // ignore other nodes
                    richTextSubnode = richTextSubnode.next_sibling_of_type(pugi::node_element);
                }
            }
            else {} // ignore other elements (for now)
            textElement = textElement.next_sibling_of_type(pugi::node_element);
        }
        return result;
    }
}    // namespace


// ========== XLComment Member Functions

/**
 * @details
 */
// XLComment::XLComment()
//  : m_commentNode(std::make_unique<XMLNode>())
//  {}

/**
* @details
*/
XLComment::XLComment(const XMLNode& node)
 : m_commentNode(std::make_unique<XMLNode>(node))
 {}

/**
 * @details
 * @note Function body moved to cpp module as it uses "not" keyword for readability, which MSVC sabotages with non-CPP compatibility.
 * @note For the library it is reasonable to expect users to compile it with MSCV /permissive- flag, but for the user's own projects the header files shall "just work"
 */
bool XLComment::valid() const { return m_commentNode != nullptr &&(not m_commentNode->empty()); }

/**
 * @brief Getter functions
 */
std::string XLComment::ref() const { return m_commentNode->attribute("ref").value(); }
std::string XLComment::text() const { return getCommentString(*m_commentNode); }
uint16_t XLComment::authorId() const { return static_cast<uint16_t>(m_commentNode->attribute("authorId").as_uint()); }

/**
 * @brief Setter functions
 */
bool XLComment::setText(std::string newText)
{
    m_commentNode->remove_children(); // clear previous text
    XMLNode tNode = m_commentNode->prepend_child("text").prepend_child("t");   // insert <text><t/></text> nodes
    tNode.append_attribute("xml:space").set_value("preserve");                // set <t> node attribute xml:space
    return tNode.prepend_child(pugi::node_pcdata).set_value(newText.c_str()); // finally, insert <t> node_pcdata value
}
bool XLComment::setAuthorId(uint16_t newAuthorId) { return appendAndSetAttribute(*m_commentNode, "authorId", std::to_string(newAuthorId)).empty() == false; }


// ========== XLComments Member Functions

XLComments::XLComments()
 : XLXmlFile(nullptr),
   m_vmlDrawing(std::make_unique<XLVmlDrawing>())
 {}

/**
 * @details The constructor creates an instance of the superclass, XLXmlFile
 */
XLComments::XLComments(XLXmlData* xmlData)
 : XLXmlFile(xmlData),
   m_vmlDrawing(std::make_unique<XLVmlDrawing>())
{
    if (xmlData->getXmlType() != XLContentType::Comments)
        throw XLInternalError("XLComments constructor: Invalid XML data.");

    XMLDocument & doc = xmlDocument();
    if (doc.document_element().empty())    // handle a bad (no document element) comments XML file
        doc.load_string(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
                "<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
                " xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">\n"
                "</comments>",
                pugi_parse_settings
        );

    XMLNode rootNode = doc.document_element();
    bool docNew = rootNode.first_child_of_type(pugi::node_element).empty();   // check and store status: was document empty?
    m_authors = appendAndGetNode (rootNode, "authors", m_nodeOrder);          // (insert and) get authors node
    if (docNew) rootNode.prepend_child(pugi::node_pcdata).set_value("\n\t");  // if document was empty: prefix the newly inserted authors node with a single tab
    m_commentList = appendAndGetNode (rootNode, "commentList", m_nodeOrder);  // (insert and) get commentList node -> this should now copy the whitespace prefix of m_authors

    // whitespace formatting / indentation for closing tags:
    if (    m_authors.first_child().empty())     m_authors.append_child(pugi::node_pcdata).set_value("\n\t");
    if (m_commentList.first_child().empty()) m_commentList.append_child(pugi::node_pcdata).set_value("\n\t");
}

/**
 * @details copy-construct an XLComments object
 */
XLComments::XLComments(const XLComments& other)
    : XLXmlFile(other),
      m_authors(other.m_authors),
      m_commentList(other.m_commentList),
      m_vmlDrawing(std::make_unique<XLVmlDrawing>(*other.m_vmlDrawing)),
      // m_vmlDrawing(std::make_unique<XLVmlDrawing>(other.m_vmlDrawing ? *other.m_vmlDrawing : XLVmlDrawing())) // this can be used if other.m_vmlDrawing can be uninitialized
      m_hintNode(other.m_hintNode),
      m_hintIndex(other.m_hintIndex)
{}

/**
 * @details move-construct an XLComments object
 */
XLComments::XLComments(XLComments&& other) noexcept
    : XLXmlFile(other),
      m_authors(std::move(other.m_authors)),
      m_commentList(std::move(other.m_commentList)),
      m_vmlDrawing(std::move(other.m_vmlDrawing)),
      m_hintNode(other.m_hintNode),
      m_hintIndex(other.m_hintIndex)
{}

/**
 * @details move-assign an XLComments object
 */
XLComments& XLComments::operator=(XLComments&& other) noexcept
{
    if (&other != this) {
        XLXmlFile::operator=(std::move(other));
        m_authors          = std::move(other.m_authors);
        m_commentList      = std::move(other.m_commentList);
        m_vmlDrawing       = std::move(other.m_vmlDrawing);
        m_hintNode         = std::move(other.m_hintNode);
        m_hintIndex        = other.m_hintIndex;
    }
    return *this;
}

/**
 * @details copy-assign an XLComments object
 */
XLComments& XLComments::operator=(const XLComments& other)
{
    if (&other != this) {
        XLComments temp = other; // copy-construct
        *this = std::move(temp); // move-assign & invalidate temp
    }
    return *this;
}


bool XLComments::setVmlDrawing(XLVmlDrawing &vmlDrawing)
{
    m_vmlDrawing = std::make_unique<XLVmlDrawing>(vmlDrawing);
    return true;
}

/**
 * @details TODO: write doxygen headers for functions in this module
 */
XMLNode XLComments::authorNode(uint16_t index) const
{
    XMLNode auth = m_authors.first_child_of_type(pugi::node_element);
    uint16_t i = 0;
    while (not auth.empty() && i != index) {
        ++i;
        auth = auth.next_sibling_of_type(pugi::node_element);
    }
    return auth; // auth.empty() will be true if not found
}

/**
 * @brief find a comment XML node by index (sequence within source XML)
 * @param index the position (0-based) of the comment node to return
 * @throws XLException if index is out of bounds vs. XLComments::count()
 */
XMLNode XLComments::commentNode(size_t index) const
{
    if( m_hintNode.empty() || m_hintIndex > index ) { // check if m_hintNode can be used - otherwise initialize it
        m_hintNode = m_commentList.first_child_of_type(pugi::node_element);
        m_hintIndex = 0;
    }

    while (not m_hintNode.empty() && m_hintIndex < index) {
        ++m_hintIndex;
        m_hintNode = m_hintNode.next_sibling_of_type(pugi::node_element);
    }
    if (m_hintNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLComments::commentNode: index "s + std::to_string(index) + " is out of bounds"s);
    }
    return m_hintNode; // can be empty XMLNode if index is >= count
}

/**
 * @details TODO: write doxygen headers for functions in this module
 */
XMLNode XLComments::commentNode(const std::string& cellRef) const
{
    return m_commentList.find_child_by_attribute("comment", "ref", cellRef.c_str());
}

/**
 * @details TODO: write doxygen headers for functions in this module
 */
uint16_t XLComments::authorCount() const
{
    XMLNode auth = m_authors.first_child_of_type(pugi::node_element);
    uint16_t count = 0;
    while (not auth.empty()) {
        ++count;
        auth = auth.next_sibling_of_type(pugi::node_element);
    }
    return count;
}

/**
 * @details TODO: write doxygen headers for functions in this module
 */
std::string XLComments::author(uint16_t index) const
{
    XMLNode auth = authorNode(index);
    if (auth.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLComments::author: index "s + std::to_string(index) + " is out of bounds"s);
    }
    return auth.first_child().value(); // author name is stored as a node_pcdata within the author node
}

/**
 * @details TODO: write doxygen headers for functions in this module
 */
bool XLComments::deleteAuthor(uint16_t index)
{
    XMLNode auth = authorNode(index);
    if (auth.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLComments::deleteAuthor: index "s + std::to_string(index) + " is out of bounds"s);
    }
    else {
        while (auth.previous_sibling().type() == pugi::node_pcdata) // remove leading whitespaces
            m_authors.remove_child(auth.previous_sibling());
        m_authors.remove_child(auth);                             // then remove author node itself
    }
    return true;
}

/**
 * @details insert author and return index
 */
uint16_t XLComments::addAuthor(const std::string& authorName)
{
    XMLNode auth = m_authors.first_child_of_type(pugi::node_element);
    uint16_t index = 0;
    while (not auth.next_sibling_of_type(pugi::node_element).empty()) {
        ++index;
        auth = auth.next_sibling_of_type(pugi::node_element);
    }
    if (auth.empty()) { // if this is the first entry
        auth = m_authors.prepend_child("author"); // insert new node
        m_authors.prepend_child(pugi::node_pcdata).set_value("\n\t\t");  // prefix first author with second level indentation
    }
    else { // found the last author node at index
        auth = m_authors.insert_child_after("author", auth);               // append a new author
        copyLeadingWhitespaces(m_authors, auth.previous_sibling(), auth ); // copy whitespaces prefix from previous author
        ++index;                                                           // increment index
    }
    auth.prepend_child(pugi::node_pcdata).set_value(authorName.c_str());
    return index;
}

/**
 * @details TODO: write doxygen headers for functions in this module
 */
size_t XLComments::count() const
{
    XMLNode comment = m_commentList.first_child_of_type(pugi::node_element);
    size_t count = 0;
    while (not comment.empty()) {
        // if (comment.name() == "comment") // TBD: safe-guard against potential rogue node
        ++count;
        comment = comment.next_sibling_of_type(pugi::node_element);
    }
    return count;
}

/**
 * @details TODO: write doxygen headers for functions in this module
 */
uint16_t XLComments::authorId(const std::string& cellRef) const
{
    XMLNode comment = commentNode(cellRef);
    return static_cast<uint16_t>(comment.attribute("authorId").as_uint());
}

/**
 * @details TODO: write doxygen headers for functions in this module
 */
bool XLComments::deleteComment(const std::string& cellRef)
{
    XMLNode comment = commentNode(cellRef);
    if (comment.empty()) return false;
    else {
        m_commentList.remove_child(comment);
        m_hintNode = XMLNode{}; // reset hint after modification of comment list
        m_hintIndex = 0;
    }
    // ===== Delete the shape associated with the comment.
    OpenXLSX::ignore(m_vmlDrawing->deleteShape(cellRef)); // disregard if deleteShape fails
    return true;
}

// comment entries:
//  attribute ref -> cell
//  attribute authorId -> author index in array
//  node text
//     subnode t -> regular text
//        attribute xml:space="preserve" - seems useful to always apply
//        subnode pc_data -> the comment text
//     subnode r -> rich text, repetition of:
//        subnode rPr -> rich text formatting
//           subnode sz -> font size (int)
//           subnode rFont -> font name (string)
//           subnode family -> TBC: font family (int)
//        subnode t -> regular text like above

/**
 * @details TODO: write doxygen headers for functions in this module
 */
XLComment XLComments::get(size_t index) const { return XLComment(commentNode(index)); }
std::string XLComments::get(const std::string& cellRef) const { return getCommentString(commentNode(cellRef)); }

/**
 * @details TODO: write doxygen headers for functions in this module
 */
bool XLComments::set(std::string const& cellRef, std::string const& commentText, uint16_t authorId_)
{
    XLCellReference destRef(cellRef);
    uint32_t destRow = destRef.row();
    uint16_t destCol = destRef.column();
    bool newCommentCreated = false; // if false, try to find an existing shape before creating one

    using namespace std::literals::string_literals;
    XMLNode comment = m_commentList.first_child_of_type(pugi::node_element);
    while (not comment.empty()) {
        if (comment.name() == "comment"s) { // safeguard against rogue nodes
            XLCellReference ref(comment.attribute("ref").value());
            if (ref.row() > destRow ||(ref.row() == destRow && ref.column() >= destCol)) // abort when node or a node behind it is found
    break;
        }
        comment = comment.next_sibling_of_type(pugi::node_element);
    }
    if(comment.empty()) {                                                     // no comments yet or this will be the last comment
        comment = m_commentList.last_child_of_type(pugi::node_element);
        if (comment.empty()) {                                                   // if this is the only comment so far
            comment = m_commentList.prepend_child("comment");                                   // prepend new comment
            m_commentList.insert_child_before(pugi::node_pcdata, comment).set_value("\n\t\t");  // insert double indent before comment
        }
        else {
            comment = m_commentList.insert_child_after("comment", comment);                     // insert new comment at end of list
            copyLeadingWhitespaces(m_commentList, comment.previous_sibling(), comment);         // and copy whitespaces prefix from previous comment
        }
        newCommentCreated = true;
    }
    else {
        XLCellReference ref(comment.attribute("ref").value());
        if(ref.row() != destRow || ref.column() != destCol) {                 // if node has to be inserted *before* this one
            comment = m_commentList.insert_child_before("comment", comment);        // insert new comment
            copyLeadingWhitespaces(m_commentList, comment, comment.next_sibling()); // and copy whitespaces prefix from next node
            newCommentCreated = true;
        }
        else // node exists / was found
            comment.remove_children();    // clear node content
    }

    // ===== If the list of nodes was modified, re-set m_hintNode that is used to access nodes by index
    if (newCommentCreated) {
        m_hintNode = XMLNode{}; // reset hint after modification of comment list
        m_hintIndex = 0;
    }

    // now that we have a valid comment node: update attributes and content
    if (comment.attribute("ref").empty())                                  // if ref has to be created
        comment.append_attribute("ref").set_value(destRef.address().c_str()); // then do so - otherwise it can remain untouched
    appendAndSetAttribute(comment, "authorId", std::to_string(authorId_)); // update authorId
    XMLNode tNode = comment.prepend_child("text").prepend_child("t");      // insert <text><t/></text> nodes
    tNode.append_attribute("xml:space").set_value("preserve");             // set <t> node attribute xml:space
    tNode.prepend_child(pugi::node_pcdata).set_value(commentText.c_str()); // finally, insert <t> node_pcdata value

    if (m_vmlDrawing->valid()) {
        XLShape cShape{};
        bool newShapeNeeded = newCommentCreated; // on new comments: create new shape
        if (!newCommentCreated) {
           try {
               cShape = shape(cellRef); // for existing comments, try to access existing shape
           }
           catch (XLException const& e) {
               newShapeNeeded = true; // not found: create fresh
           }
        }
        if (newShapeNeeded)
            cShape = m_vmlDrawing->createShape();

        cShape.setFillColor("#ffffc0");
        cShape.setStroked(true);
        // setType: already done by XLVmlDrawing::createShape
        cShape.setAllowInCell(false);
        {
            // XLShapeStyle shapeStyle("position:absolute;margin-left:100pt;margin-top:0pt;width:50pt;height:50.0pt;mso-wrap-style:none;v-text-anchor:middle;visibility:hidden");
            XLShapeStyle shapeStyle{}; // default construct with meaningful values
            cShape.setStyle(shapeStyle);
        }

        XLShapeClientData clientData = cShape.clientData();
        clientData.setObjectType("Note");
        clientData.setMoveWithCells();
        clientData.setSizeWithCells();

        {
            constexpr const uint16_t leftColOffset = 1;
            constexpr const uint16_t widthCols = 2;
            constexpr const uint16_t topRowOffset = 1;
            constexpr const uint16_t heightRows = 2;

            uint16_t anchorLeftCol, anchorRightCol;
            if( OpenXLSX::MAX_COLS - destCol > leftColOffset + widthCols ) {
                anchorLeftCol  = (destCol - 1) + leftColOffset;
                anchorRightCol = (destCol - 1) + leftColOffset + widthCols;
            }
            else { // if anchor would overflow MAX_COLS: move column anchor to the left of destCol
                anchorLeftCol  = (destCol - 1) - leftColOffset - widthCols;
                anchorRightCol = (destCol - 1) - leftColOffset;
            }

            uint32_t anchorTopRow, anchorBottomRow;
            if( OpenXLSX::MAX_ROWS - destRow > topRowOffset + heightRows ) {
                anchorTopRow    = (destRow - 1) + topRowOffset;
                anchorBottomRow = (destRow - 1) + topRowOffset + heightRows;
            }
            else { // if anchor would overflow MAX_ROWS: move row anchor to the top of destCol
                anchorTopRow    = (destRow - 1) - topRowOffset - heightRows;
                anchorBottomRow = (destRow - 1) - topRowOffset;
            }
            if (anchorRightCol > MAX_SHAPE_ANCHOR_COLUMN)
                std::cout << "XLComments::set WARNING: anchoring comment shapes beyond column "s
                /**/          + XLCellReference::columnAsString(MAX_SHAPE_ANCHOR_COLUMN) + " may not get displayed correctly (LO Calc, TBD in Excel)"s << std::endl;
            if (anchorBottomRow > MAX_SHAPE_ANCHOR_ROW)
                std::cout << "XLComments::set WARNING: anchoring comment shapes beyond row "s
                /**/          +                  std::to_string(MAX_SHAPE_ANCHOR_ROW   ) + " may not get displayed correctly (LO Calc, TBD in Excel)"s << std::endl;

            uint16_t anchorLeftOffsetInCell = 10, anchorRightOffsetInCell = 10;
            uint16_t anchorTopOffsetInCell = 5, anchorBottomOffsetInCell = 5;

            // clientData.setAnchor("3, 23, 0, 0, 4, 25, 3, 5");
            using namespace std::literals::string_literals;
            clientData.setAnchor(
                  std::to_string(anchorLeftCol)           + ","s
                + std::to_string(anchorLeftOffsetInCell)  + ","s
                + std::to_string(anchorTopRow)            + ","s
                + std::to_string(anchorTopOffsetInCell)   + ","s
                + std::to_string(anchorRightCol)          + ","s
                + std::to_string(anchorRightOffsetInCell) + ","s
                + std::to_string(anchorBottomRow)         + ","s
                + std::to_string(anchorBottomOffsetInCell)
            );
        }
        clientData.setAutoFill(false);
        clientData.setTextVAlign(XLShapeTextVAlign::Top);
        clientData.setTextHAlign(XLShapeTextHAlign::Left);
        clientData.setRow(destRow - 1);    // row and column are zero-indexed in XLShapeClientData
        clientData.setColumn(destCol - 1); // ..

	// 	<v:shadow on="t" obscured="t" color="black"/>
	// 	<v:fill o:detectmouseclick="t" type="solid" color2="#00003f"/>
	// 	<v:stroke color="#3465a4" startarrow="block" startarrowwidth="medium" startarrowlength="medium" joinstyle="round" endcap="flat"/>
	// 	<x:ClientData ObjectType="Note">
	// 		<x:MoveWithCells/>
	// 		<x:SizeWithCells/>
	// 		<x:Anchor>3, 23, 0, 0, 4, 25, 3, 5</x:Anchor>
	// 		<x:AutoFill>False</x:AutoFill>
	// 		<x:TextVAlign>Top</x:TextVAlign>
	// 		<x:TextHAlign>Left</x:TextHAlign>
	// 		<x:Row>0</x:Row>
	// 		<x:Column>2</x:Column>
	// 	</x:ClientData>
	// </v:shape>
    }
    else
        throw XLException("XLComments::set: can not set (format) any comments when VML Drawing object is invalid");

    return true;
}

/**
 * @details
 */
XLShape XLComments::shape(std::string const& cellRef)
{
    if (!m_vmlDrawing->valid())
        throw XLException("XLComments::shape: can not access any shapes when VML Drawing object is invalid");

    XMLNode shape = m_vmlDrawing->shapeNode(cellRef);
    if (shape.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLComments::shape: not found for cell "s + cellRef + " - was XLComment::set invoked first?"s);
    }
    return XLShape(shape);
}

/**
 * @details Print the underlying XML using pugixml::xml_node::print
 */
void XLComments::print(std::basic_ostream<char>& ostr) const { xmlDocument().document_element().print( ostr ); }

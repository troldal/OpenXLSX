#include <cctype>       // std::isdigit (issue #330)
#include <string>       // std::string
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"               // pugi_parse_settings
//#include "XLDrawing.hpp"
#include "utilities/XLUtilities.hpp"    // OpenXLSX::ignore, appendAndGetNode

#include "XLDrawing1.hpp"

using namespace OpenXLSX;

const std::string ShapeNodeNameDr0 = "xdr:absoluteAnchor";
const std::string ShapeNodeNameDr1 = "xdr:oneCellAnchor";
const std::string ShapeNodeNameDr2 = "xdr:twoCellAnchor";

static bool nodeDr(const std::string s)
{
	if(s==ShapeNodeNameDr1)return true;
	if(s==ShapeNodeNameDr2)return true;
	if(s==ShapeNodeNameDr0)return true;
	return false;
}


XLDrawing1::XLDrawing1(XLXmlData* xmlData) : XLXmlFile(xmlData)
{
	if (xmlData->getXmlType() != XLContentType::Drawing)
		throw XLInternalError("XLDrawing constructor: Invalid XML data.");
	XMLDocument& doc = xmlDocument();
	if (doc.document_element().empty()) {  // handle a bad (no document element) drawing XML file
		std::string s1 = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
		std::string s2 = "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"</xdr:wsDr>";
		doc.load_string((s1 + "\n" + s2).data(), pugi_parse_settings);
	}
	XMLNode rootNode = doc.document_element();
	XMLNode node = rootNode.first_child_of_type(pugi::node_element);
	m_shapeCount=0;
	while (not node.empty() && nodeDr(node.raw_name())) {
		XMLNode nextNode = node.next_sibling_of_type(pugi::node_element); // determine next node early because node may be invalidated by moveNode
		++m_shapeCount;
		node = nextNode;
	}

}

XMLNode XLDrawing1::rootNode() const {
	return xmlDocument().document_element();
}

XMLNode XLDrawing1::firstShapeNode() const
{
	XMLNode node = xmlDocument().document_element().first_child_of_type(pugi::node_element);
	while (not node.empty() && !nodeDr(node.raw_name()))   // skip non shape nodes
		node = node.next_sibling_of_type(pugi::node_element);
	return node;
}

XMLNode XLDrawing1::lastShapeNode() const
{
	XMLNode node = xmlDocument().document_element().last_child_of_type(pugi::node_element);
	while (not node.empty() && !nodeDr(node.raw_name()))
		node = node.previous_sibling_of_type(pugi::node_element);
	return node;

}

XMLNode XLDrawing1::shapeNode(uint32_t index) const
{
	using namespace std::literals::string_literals;

	XMLNode node{}; // scope declaration, ensures node.empty() when index >= m_shapeCount
	if (index < shapeCount()) {
		uint16_t i = 0;
		node = firstShapeNode();
		while (i != index && nodeDr(node.raw_name())) {
			++i;
			node = node.next_sibling_of_type(pugi::node_element);
		}
	}
	if (node.empty() || !nodeDr(node.raw_name()))
		throw XLException("XLDrawing: shape index "s + std::to_string(index) + " is out of bounds"s);

	return node;
}

XMLNode XLDrawing1::shapeNode(std::string const& cellRef) const
{
	XLCellReference destRef(cellRef);
	uint32_t destRow = destRef.row() - 1;    // for accessing a shape: x:Row and x:Column are zero-indexed
	uint16_t destCol = destRef.column() - 1; // ..

	XMLNode node = firstShapeNode();
	while (not node.empty() && nodeDr(node.raw_name())) {
		if ((destRow == node.child("x:ClientData").child("x:Row").text().as_uint())
			&& (destCol == node.child("x:ClientData").child("x:Column").text().as_uint()))
			break; // found shape for cellRef

		do { // locate next shape node
			node = node.next_sibling_of_type(pugi::node_element);
		} while (not node.empty() && !nodeDr(node.name()));
	}
	return node;
}

uint32_t XLDrawing1::shapeCount() const { return m_shapeCount; }

XLShape1 XLDrawing1::createShape(int32_t n)
{
	XMLNode rootNode = xmlDocument().document_element();
	XMLNode node;
	switch (n) {
		case 0:node = rootNode.append_child(ShapeNodeNameDr0.c_str()); break;
		case 1:node = rootNode.append_child(ShapeNodeNameDr1.c_str()); break;
		case 2:node = rootNode.append_child(ShapeNodeNameDr2.c_str()); break;
	}
	m_shapeCount++;
	return XLShape1(node);
}

XLShape1 XLDrawing1::shape(uint32_t index) const { return XLShape1(shapeNode(index)); }

XLShape1::XLShape1() : m_shapeNode(std::make_unique<XMLNode>(XMLNode())) {}

XLShape1::XLShape1(const XMLNode& node) : m_shapeNode(std::make_unique<XMLNode>(node)) {}


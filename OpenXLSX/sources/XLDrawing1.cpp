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

void setAttribute(XMLNode n, char* path, char* attribute, char* value)
{
	char* s, * s0; int i, att = 0, val = 0,skip=0; XMLNode f;
	std::string ss, sa, sv;
	s = s0 = path;
	while (1) {
		if (*s == '/' || *s == '@' || *s == '=' || *s == ',' || *s == '|' || !*s) {
			i = s - s0;
			if (i)ss = std::string((const char*)s0, (size_t)i);
			s0 = s + 1;
			if (att && i) {
				if (*s == '=') {
					sa = ss;
				}
				else {
					if (val) {
						sv = ss;
					}
				}
			}
			if (att && val && i) {
				const pugi::char_t* a = (const pugi::char_t*)sa.data();
				if (n.attribute(a).empty())n.append_attribute(a) = sv.data();
				else n.attribute(a).set_value(sv.data());
				att = 0;
				val = 0;
			}
			else {
				if (!att && !val && i && !skip) {
					f = n.first_child_of_type(pugi::node_element);
					while (!f.empty()) {
						if (f.raw_name() == ss)break;
						f = f.next_sibling_of_type(pugi::node_element);
					}
					if (f.empty()) {
						if (*s != ',' && *s != '|') {
							f = n.append_child(ss.c_str());
							n = f;
						}
					}
					else {
						if (*s == ',' || *s == '|')skip = 1;
						n = f;
					}
				}
			}
			if (!*s)break;
			if (*s == '@')att = 1;
			if (*s == '=')val = 1;
			if (*s == '/')skip = 0;
			if (*s == '/' || *s == ',' || *s == '|') {
				s0 = s + 1;
			}
		}
		s++;
	}
	if (attribute && *attribute) {
		const pugi::char_t* a = (const pugi::char_t*)attribute;
		if (n.attribute(a).empty())n.append_attribute(a) = value;
		else n.attribute(a).set_value(value);
	}
	else {
		if (value && *value)
			n.text() = value;
	}

}

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
	uint32_t index = m_shapeCount;
	m_shapeCount++;
	return XLShape1(node);
}

XLShape1 XLDrawing1::shape(uint32_t index) const { return XLShape1(shapeNode(index)); }

XLShape1::XLShape1() : m_shapeNode(std::make_unique<XMLNode>(XMLNode())) {}

XLShape1::XLShape1(const XMLNode& node) : m_shapeNode(std::make_unique<XMLNode>(node)) {}

static XMLNode setNode(XMLNode root, std::string name)
{
	XMLNode f;
	f = root.first_child_of_type(pugi::node_element);
	while (!f.empty()) {
		if (f.raw_name() == name)break;
		f = f.next_sibling_of_type(pugi::node_element);
	}
	if (f.empty()) {
		f = root.append_child(name.c_str());
	}
	return f;
}

static XMLNode getNode(XMLNode root, std::string name)
{
	XMLNode f;
	f = root.first_child_of_type(pugi::node_element);
	while (!f.empty()) {
		if (f.raw_name() == name)break;
		f = f.next_sibling_of_type(pugi::node_element);
	}
	return f;
}

static XMLNode setNodeBefore(XMLNode root, std::string name,std::string before)
{
	XMLNode f; XMLNode nodebefore = nullptr;
	f = root.first_child_of_type(pugi::node_element);
	while (!f.empty()) {
		if (f.raw_name() == name)break;
		if (f.raw_name() == before) {
			nodebefore = f;
		}
		f = f.next_sibling_of_type(pugi::node_element);
	}
	if (f.empty()) {
		if (nodebefore.empty()) f = root.append_child(name.c_str());
		else f = root.prepend_child(name.c_str());
	}
	return f;
}


static void setAtt(XMLNode node, char* att, char* val)
{
	if (node.attribute(att).empty())node.append_attribute(att) = val;
	else node.attribute(att).set_value(val);
}

XLTextFrame1 XLShape1::textFrame()
{
	XMLNode *n = xmlnode();
	XMLNode node=setNode(*n,"xdr:sp");
	node = setNode(node,"xdr:txBody");
	setNode(node,"a:bodyPr");
	setNode(node,"a:p");
	return XLTextFrame1(node);
}
XLFill2 XLShape1::fill()
{
	XMLNode* n = xmlnode();
	XMLNode node = setNode(*n, "xdr:blipFill");
	return XLFill2(node);
}

XLPictureFormat1 XLShape1::pictureFormat()
{
	XMLNode* n = xmlnode();
	XMLNode node = getNode(*n, "xdr:pic");
	if (node.empty())node = getNode(*n, "xdr:cxnSp");
	if (node.empty())node = getNode(*n, "xdr:sp");
	node = setNode(node, "xdr:spPr");
	return XLPictureFormat1(node);
}

void XLDrawing1::setShapeAttribute(uint32_t shapeIndex,char* path, char* attribute, char* value)
{
	XMLNode n = shapeNode(shapeIndex);
	setAttribute(n,path,attribute,value);
}

XLTextFrame1::XLTextFrame1() : m_node(std::make_unique<XMLNode>(XMLNode())) {}

XLTextFrame1::XLTextFrame1(const XMLNode& node) : m_node(std::make_unique<XMLNode>(node)) {}

void XLTextFrame1::setVerticalAlignment(std::string val)
{
	std::string v = "";
	XMLNode* n = xmlnode();
	if (val == "bottom")v = "b";
	if (val == "center")v = "ctr";
	if (val == "distributed")v = "dist";
	if (val == "justify")v = "just";
	if (val == "top")v = "t";
	if (!v.length())return;
	setAttribute(*n, (char *)"a:bodyPr", (char *)"anchor", v.data());
}
void XLTextFrame1::setHorizontalAlignment(std::string val)
{
	std::string v = "0";
	XMLNode* n = xmlnode();
	if (val == "center")v = "1";
	setAttribute(*n, (char*)"a:bodyPr", (char*)"anchorCtr", v.data());
}

XLPictureFormat1::XLPictureFormat1() : m_node(std::make_unique<XMLNode>(XMLNode())) {}

XLPictureFormat1::XLPictureFormat1(const XMLNode& node) : m_node(std::make_unique<XMLNode>(node)) {}

XMLNode XLPictureFormat1::setBlipFill()
{
	XMLNode* n = xmlnode();
	return setNode(*n, "a:blipFill");
}

XMLNode XLPictureFormat1::setPrstGeom(char *prst)
{
	XMLNode* n = xmlnode();
	XMLNode node=setNode(*n, "a:prstGeom");
	node.append_attribute("prst") = prst;
	return node;
}

XMLNode XLPictureFormat1::setSolidFill()
{
	XMLNode* n = xmlnode();
	return setNode(*n, "a:solidFill");

}

XMLNode XLPictureFormat1::setXfrm()
{
	XMLNode* n = xmlnode();
	return setNode(*n, "a:xfrm");
}

XMLNode XLPictureFormat1::setLn()
{
	XMLNode* n = xmlnode();
	return setNode(*n, "a:ln");
}

XLFill2::XLFill2() : m_node(std::make_unique<XMLNode>(XMLNode())) {}

XLFill2::XLFill2(const XMLNode& node) : m_node(std::make_unique<XMLNode>(node)) {}

XLCharacters2 XLTextFrame1::Characters()
{
	XMLNode* n = xmlnode();
	XMLNode node=setNode(*n, "a:p");
	node = setNode(node, "a:r");
	return XLCharacters2(node);
}

XLCharacters2::XLCharacters2() : m_node(std::make_unique<XMLNode>(XMLNode())) {}

XLCharacters2::XLCharacters2(const XMLNode& node) : m_node(std::make_unique<XMLNode>(node)) {}

void XLCharacters2::setText(std::string text)
{
	XMLNode *n=xmlnode();
	XMLNode t = setNode(*n, "a:t");
	t.text() = text.c_str();
}

XLFont2 XLCharacters2::font()
{
	XMLNode* n = xmlnode();
	XMLNode node = setNodeBefore(*n, "a:rPr","a:t");
	return XLFont2(node);
}

XLFont2::XLFont2() : m_node(std::make_unique<XMLNode>(XMLNode())) {}

XLFont2::XLFont2(const XMLNode& node) : m_node(std::make_unique<XMLNode>(node)) {}

void XLFont2::setBold(bool val)
{
	char* b =(char *)"0";
	XMLNode* n = xmlnode();
	if (val)b = (char *)"1";
	setAtt(*n, (char *)"b", b);
}

void XLFont2::setItalic(bool val)
{
	char* b = (char*)"0";
	XMLNode* n = xmlnode();
	if (val)b = (char*)"1";
	setAtt(*n, (char*)"i", b);
}
void XLFont2::setStrikethrough(bool val)
{
	char* b = (char*)"noStrike";
	XMLNode* n = xmlnode();
	if (val)b = (char*)"sngStrike";
	setAtt(*n, (char*)"strike", b);
}

void XLFont2::setSize(int32_t val)
{
	std::string s=std::to_string(100*val);
	XMLNode* n = xmlnode();
	setAtt(*n, (char*)"sz", s.data());
}

void XLFont2::setUnderline(int32_t val)
{
	char* s = (char *)"sng";
	if (val == 2)s = (char *)"dbl";
	XMLNode* n = xmlnode();
	setAtt(*n, (char*)"u", s);
}

void XLShape1::setName(std::string name)
{
	XMLNode node = sp();
	if (node.empty()) {
		node = cxnSp();
		if (node.empty()) {
			node = pic();
			node = setNode(node,"xdr:nvPicPr");
		}
		else {
			node = setNode(node,"xdr:nvCxnSpPr");
		}
	}
	else {
		node = setNode(node, "xdr:nvSpPr");
	}
	node = setNode(node, "xdr:cNvPr");
	setAtt(node,(char*) "name", name.data());
}

void XLShape1::setId(int32_t id)
{
	XMLNode node = sp();
	if (node.empty()) {
		node = cxnSp();
		if (node.empty()) {
			node = pic();
			node = setNode(node, "xdr:nvPicPr");
		}
		else {
			node = setNode(node, "xdr:nvCxnSpPr");
		}
	}
	else {
		node = setNode(node, "xdr:nvSpPr");
	}
	node = setNode(node, "xdr:cNvPr");
	setAtt(node, (char*)"id", std::to_string(id).data());
}

void XLShape1::setRotation(float rot)
{
	XMLNode *node = xmlnode();
	int irot = (int)(60000 * rot);
	setAttribute(*node,(char *)"xdr:sp|xdr:pic/xdr:spPr/a:xfrm", (char *)"rot", std::to_string(irot).data());
}
void XLShape1::setLeft(float v)
{
}
void XLShape1::setTop(float v)
{
}
void XLShape1::setWidth(float v)
{
}
void XLShape1::setHeight(float v)
{
}
std::string name()
{
	return "";
}
float XLShape1::rotation()
{
	return 0;
}
float XLShape1::left()
{
	return 0;
}
float XLShape1::top()
{
	return 0;
}
float XLShape1::width()
{
	return 0;
}

float XLShape1::height()
{
	return 0;
}

XMLNode XLShape1::cxnSp()
{
	XMLNode* n = xmlnode();
	return getNode(*n, "xdr:cxnSp");
}

XMLNode XLShape1::sp()
{
	XMLNode* n = xmlnode();
	return getNode(*n, "xdr:sp");
}

XMLNode XLShape1::pic()
{
	XMLNode* n = xmlnode();
	return getNode(*n, "xdr:pic");
}

XMLNode XLShape1::setCxnSp(char *macro)
{
	XMLNode* n = xmlnode();
	XMLNode node=setNode(*n,"xdr:cxnSp");
	if (macro && *macro)setAtt(node, (char*)"macro", macro);
	return node;
}

XMLNode XLShape1::setSp(char *macro,char *textlink)
{
	XMLNode* n = xmlnode();
	XMLNode node=setNode(*n, "xdr:sp");
	if(macro && *macro)setAtt(node, (char *)"macro", macro);
	if(textlink && *textlink)setAtt(node, (char*)"textlink", textlink);
	return node;
}

XMLNode XLShape1::setPic(char* macro)
{
	XMLNode* n = xmlnode();
	XMLNode node = setNode(*n, "xdr:pic");
	if(macro && *macro)setAtt(node, (char*)"macro", macro);
	return node;
}

XMLNode XLShape1::setClientData()
{
	XMLNode* n = xmlnode();
	return setNode(*n, "xdr:clientData");
}

void XLShape1::from(int row, int col, int rowoff, int coloff)
{
	XMLNode* n = xmlnode();
	XMLNode node = setNode(*n,"xdr:from");
	XMLNode nn = setNode(node, "xdr:col");
	nn.text() = std::to_string(col).data();
	nn = setNode(node, "xdr:colOff");
	nn.text() = std::to_string(coloff).data();
	nn = setNode(node, "xdr:row");
	nn.text() = std::to_string(row).data();
	nn = setNode(node, "xdr:rowOff");
	nn.text() = std::to_string(coloff).data();
}

void XLShape1::to(int row, int col, int rowoff, int coloff)
{
	XMLNode* n = xmlnode();
	XMLNode node=setNode(*n, "xdr:to");
	XMLNode nn=setNode(node, "xdr:col");
	nn.text() = std::to_string(col).data();
	nn = setNode(node, "xdr:colOff");
	nn.text() = std::to_string(coloff).data();
	nn = setNode(node, "xdr:row");
	nn.text() = std::to_string(row).data();
	nn = setNode(node, "xdr:rowOff");
	nn.text() = std::to_string(coloff).data();
}

void XLShape1::ext(int cx, int cy)
{
	XMLNode* n = xmlnode();
	XMLNode ext = setNode(*n,"xdr:ext");
	setAtt(ext,(char *)"cx",std::to_string(cx).data());
	setAtt(ext,(char *)"cy",std::to_string(cy).data());
}


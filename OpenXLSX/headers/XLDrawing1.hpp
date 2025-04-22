#ifndef OPENXLSX_XLDRAWING1_HPP
#define OPENXLSX_XLDRAWING1_HPP

// ===== External Includes ===== //
#include <cstdint>      // uint8_t, uint16_t, uint32_t
#include <ostream>      // std::basic_ostream

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLException.hpp"
#include "XLXmlData.hpp"
#include "XLXmlFile.hpp"
#include "XLDrawing.hpp"

using namespace OpenXLSX;

void setAttribute(XMLNode n, char* path, char* attribute, char* value);

class XLPictureFormat1;
class XLTextFrame1;
class XLShape1;
class XLCharacters2;
class XLFill2;
class XLFont2;

class XLDrawing1 : public XLXmlFile
{
    friend class XLShape1;
	friend class XLWorksheet;
public:
	XLDrawing1() : XLXmlFile(nullptr) {};
	XLDrawing1(XLXmlData* xmlData);
	XLDrawing1(const XLDrawing1& other) = default;
	XLDrawing1(XLDrawing1&& other) noexcept = default;
	~XLDrawing1() = default;

	XLDrawing1& operator=(const XLDrawing1&) = default;
	XLDrawing1& operator=(XLDrawing1&& other) noexcept = default;

	XMLNode shapeNode(std::string const& cellRef) const;
	XMLNode shapeNode(uint32_t index) const;

	XLShape1 shape(uint32_t index) const;
	XLShape1 createShape(int32_t n);

	uint32_t shapeCount() const;

	XMLNode rootNode() const;
    void setShapeAttribute(uint32_t shapeIndex,char* path, char* attribute, char* value);

private:
	XMLNode firstShapeNode() const;
	XMLNode lastShapeNode() const;

private:
	uint32_t m_shapeCount{ 0 };
	uint32_t m_lastAssignedShapeId{ 0 };
	std::string m_defaultShapeTypeId{};
};

class XLFill2 : public XMLNode
{
public:
    XLFill2();
    explicit XLFill2(const XMLNode& node);
    ~XLFill2() = default;
    XMLNode* xmlnode() { return m_node.get(); };
    void reset() { m_node.reset(); };
private:
    std::unique_ptr<XMLNode> m_node;
};

class XLPictureFormat1 : public XMLNode
{
public:
    XLPictureFormat1();
    explicit XLPictureFormat1(const XMLNode& node);
    ~XLPictureFormat1() = default;
    XMLNode* xmlnode() { return m_node.get(); };
    void reset() { m_node.reset(); };
    XMLNode setBlipFill();
    XMLNode setPrstGeom(char *prst);
    XMLNode setSolidFill();
    XMLNode setXfrm();
    XMLNode setLn();
private:
    std::unique_ptr<XMLNode> m_node;
};

class XLTextFrame1 : public XMLNode
{
public:
    XLTextFrame1();
    explicit XLTextFrame1(const XMLNode& node);
    ~XLTextFrame1() = default;
    XLShape1* s1() { return m_s1; };
    XLCharacters2 Characters();
    XMLNode *xmlnode() { return m_node.get(); };
    void reset() { m_node.reset(); };
    void setVerticalAlignment(std::string val);
    void setHorizontalAlignment(std::string val);
private:
    XLShape1* m_s1;
    std::unique_ptr<XMLNode> m_node;
};

class XLCharacters2
{
public :
    XLCharacters2();
    explicit XLCharacters2(const XMLNode& node);
    ~XLCharacters2() = default;
    XMLNode* xmlnode() { return m_node.get(); };
    void reset(){m_node.reset();};
    void setText(std::string text);
    XLFont2 font();
private :
    std::unique_ptr<XMLNode> m_node;
};

class XLFont2
{
public:
    XLFont2();
    explicit XLFont2(const XMLNode& node);
    ~XLFont2() = default;
    XMLNode* xmlnode() { return m_node.get(); };
    void reset() { m_node.reset(); };
    void setBold(bool val);
    void setItalic(bool val);
    void setStrikethrough(bool val);
    void setUnderline(int32_t val);
    void setSize(int32_t val);
private:
    std::unique_ptr<XMLNode> m_node;
};

class XLShape1 {
        friend class XLDrawing1;
    public:
        XLShape1();
        explicit XLShape1(const XMLNode& node);
	    XLShape1(const XLShape1& other) = default;
        XLShape1(XLShape1&& other) noexcept = default;
       ~XLShape1() = default;
        XLShape1& operator=(const XLShape1& other);
        XLShape1& operator=(XLShape1&& other) noexcept = default;
        XMLNode *xmlnode() { return m_shapeNode.get();};
        void reset() { m_shapeNode.reset(); };
        XLTextFrame1 textFrame();
        XLPictureFormat1 pictureFormat();
        XLFill2 fill();
        void from(int row, int col, int rowoff, int coloff);
        void to(int row, int col, int rowoff, int coloff);
        void ext(int cx, int cy);
        XMLNode cxnSp();
        XMLNode sp();
        XMLNode pic();
        XMLNode setCxnSp(char *macro);
        XMLNode setSp(char *macro,char *textlink);
        XMLNode setPic(char *macro);
        void setName(std::string name);
        void setId(int32_t id);
        void setRotation(float rot);
        void setLeft(float left);
        void setTop(float top);
        void setWidth(float width);
        void setHeight(float height);
        std::string name();
        float rotation();
        float left();
        float top();
        float width();
        float height();
        XMLNode setClientData();
    private:
        std::unique_ptr<XMLNode> m_shapeNode;
        inline static const std::vector< std::string_view > m_nodeOrder = {
            "xdr:ClientData"
        };
    }; 

#endif

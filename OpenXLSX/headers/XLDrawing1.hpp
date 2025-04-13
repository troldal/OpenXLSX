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

class XLShape1;

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

private:
	XMLNode firstShapeNode() const;
	XMLNode lastShapeNode() const;

private:
	uint32_t m_shapeCount{ 0 };
	uint32_t m_lastAssignedShapeId{ 0 };
	std::string m_defaultShapeTypeId{};
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

        std::string shapeId() const;   // v:shape attribute id - shape_# - can't be set by the user
        std::string fillColor() const; // v:shape attribute fillcolor, #<3 byte hex code>, e.g. #ffffc0
        bool stroked() const;          // v:shape attribute stroked "t" ("f"?)
        std::string type() const;      // v:shape attribute type, link to v:shapetype attribute id
        bool allowInCell() const;      // v:shape attribute o:allowincell "f"
//        XLShapeStyle style();          // v:shape attribute style, but constructed from the XMLAttribute

        // XLShapeShadow& shadow();          // v:shape subnode v:shadow
        // XLShapeFill& fill();              // v:shape subnode v:fill
        // XLShapeStroke& stroke();          // v:shape subnode v:stroke
        // XLShapePath& path();              // v:shape subnode v:path
        // XLShapeTextbox& textbox();        // v:shape subnode v:textbox
//        XLShapeClientData clientData();  // v:shape subnode x:ClientData

        // NOTE: setShapeId is not available because shape id is managed by the parent class in createShape
 
       bool setFillColor(std::string const& newFillColor);
        bool setStroked(bool set);
        bool setType(std::string const& newType);
        bool setAllowInCell(bool set);
        bool setStyle(std::string const& newStyle);
  //      bool setStyle(XLShapeStyle const& newStyle);

    private:
        std::unique_ptr<XMLNode> m_shapeNode;        /**< An XMLNode object with the v:shape item */
        inline static const std::vector< std::string_view > m_nodeOrder = {
            "v:shadow",
            "v:fill",
            "v:stroke",
            "v:path",
            "v:textbox",
            "x:ClientData"
        };
    }; 

#endif

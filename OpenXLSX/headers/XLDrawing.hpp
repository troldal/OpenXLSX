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

#ifndef OPENXLSX_XLDRAWING_HPP
#define OPENXLSX_XLDRAWING_HPP

// ===== External Includes ===== //
#include <cstdint>      // uint8_t, uint16_t, uint32_t
#include <ostream>      // std::basic_ostream

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
// #include "XLDocument.hpp"
#include "XLException.hpp"
#include "XLXmlData.hpp"
#include "XLXmlFile.hpp"


namespace OpenXLSX
{
		// <v:fill o:detectmouseclick="t" type="solid" color2="#00003f"/>
		// <v:shadow on="t" obscured="t" color="black"/>
		// <v:stroke color="#3465a4" startarrow="block" startarrowwidth="medium" startarrowlength="medium" joinstyle="round" endcap="flat"/>
		// <v:path o:connecttype="none"/>
		// <v:textbox style="mso-direction-alt:auto;mso-fit-shape-to-text:t;">
		// 	<div style="text-align:left;"/>
		// </v:textbox>

    extern const std::string ShapeNodeName;     // = "v:shape"
    extern const std::string ShapeTypeNodeName; // = "v:shapetype"

    // NOTE: numerical values of XLShapeTextVAlign and XLShapeTextHAlign are shared with the same alignments from XLAlignmentStyle (XLStyles.hpp)
    enum class XLShapeTextVAlign : uint8_t {
        Center           =   3, // value="center",           both
        Top              =   4, // value="top",              vertical only
        Bottom           =   5, // value="bottom",           vertical only
        Invalid          = 255  // all other values
    };
    constexpr const XLShapeTextVAlign XLDefaultShapeTextVAlign = XLShapeTextVAlign::Top;

    enum class XLShapeTextHAlign : uint8_t {
        Left             =   1, // value="left",             horizontal only
        Right            =   2, // value="right",            horizontal only
        Center           =   3, // value="center",           both
        Invalid          = 255  // all other values
    };
    constexpr const XLShapeTextHAlign XLDefaultShapeTextHAlign = XLShapeTextHAlign::Left;


    /**
     * @brief An encapsulation of a shape client data element x:ClientData
     */
    class OPENXLSX_EXPORT XLShapeClientData
    {
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLShapeClientData();

        /**
         * @brief Constructor. New items should only be created through an XLShape object.
         * @param node An XMLNode object with the x:ClientData XMLNode. If no input is provided, a null node is used.
         */
        explicit XLShapeClientData(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLShapeClientData(const XLShapeClientData& other) = default;

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLShapeClientData(XLShapeClientData&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLShapeClientData() = default;

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLShapeClientData& operator=(const XLShapeClientData& other) = default;

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLShapeClientData& operator=(XLShapeClientData&& other) noexcept = default;

        /**
         * @brief Getter functions
         */
        std::string objectType() const; // attribute ObjectType, value "Note"
        bool moveWithCells() const;     // element x:MoveWithCells - true = present or lowercase node_pcdata "true", false = not present or lowercase node_pcdata "false"
        bool sizeWithCells() const;     // element x:SizeWithCells - logic as in MoveWithCells
        std::string anchor() const;     // element x:Anchor - Example node_pcdata: "3, 23, 0, 0, 4, 25, 3, 5" - no idea what any number means - TBD
        bool autoFill() const;          // element x:AutoFill - logic as in MoveWithCells
        XLShapeTextVAlign textVAlign() const;   // element x:TextVAlign - Top, ???
        XLShapeTextHAlign textHAlign() const;   // element x:TextHAlign - Left, ???
        uint32_t row() const;           // element x:Row, 0-indexed row of cell to which this shape is linked
        uint16_t column() const;        // element x:Column, 0-indexed column of cell to which this shape is linked

        /**
         * @brief Setter functions
         */
        bool setObjectType(std::string newObjectType);
        bool setMoveWithCells(bool set = true);
        bool setSizeWithCells(bool set = true);
        bool setAnchor(std::string newAnchor);
        bool setAutoFill(bool set = true);
        bool setTextVAlign(XLShapeTextVAlign newTextVAlign);
        bool setTextHAlign(XLShapeTextHAlign newTextHAlign);
        bool setRow(uint32_t newRow);
        bool setColumn(uint16_t newColumn);

        // /**
        //  * @brief Return a string summary of the x:ClientData properties
        //  * @return string with info about the x:ClientData object
        //  */
        // std::string summary() const;

    private:                                         // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_clientDataNode;   /**< An XMLNode object with the x:ClientData item */
        inline static const std::vector< std::string_view > m_nodeOrder = {
            "x:MoveWithCells",
            "x:SizeWithCells",
            "x:Anchor",
            "x:AutoFill",
            "x:TextVAlign",
            "x:TextHAlign",
            "x:Row",
            "x:Column"
        };
     };


    struct XLShapeStyleAttribute {
        std::string name;
        std::string value;
    };

    class OPENXLSX_EXPORT XLShapeStyle {
    public:
        /**
         * @brief
         */
        XLShapeStyle();

        /**
         * @brief Constructor. Init XLShapeStyle properties from styleAttribute
         * @param styleAttribute a string with the value of the style attribute of a v:shape element
         */
        explicit XLShapeStyle(const std::string& styleAttribute);

        /**
         * @brief Constructor. Init XLShapeStyle properties from styleAttribute and link to the attribute so that setter functions directly modify it
         * @param styleAttribute an XMLAttribute constructed with the style attribute of a v:shape element
         */
        explicit XLShapeStyle(const XMLAttribute& styleAttribute);

    private:
        /**
         * @brief get index of an attribute name within m_nodeOrder
         * @return index of attribute in m_nodeOrder
         * @return -1 if not found
         */
        int16_t attributeOrderIndex(std::string const& attributeName) const;
        /**
         * @brief XLShapeStyle internal generic getter & setter functions
         */
        XLShapeStyleAttribute getAttribute(std::string const& attributeName, std::string const& valIfNotFound = "") const;
        bool setAttribute(std::string const& attributeName, std::string const& attributeValue);
        // bool setAttribute(XLShapeStyleAttribute const& attribute);

    public:
        /**
         * @brief XLShapeStyle getter functions
         */
        std::string position() const;
        uint16_t marginLeft() const;
        uint16_t marginTop() const;
        uint16_t width() const;
        uint16_t height() const;
        std::string msoWrapStyle() const;
        std::string vTextAnchor() const;
        bool hidden() const;
        bool visible() const;

        std::string raw() const { return m_style; }

        /**
         * @brief XLShapeStyle setter functions
         */
        bool setPosition(std::string newPosition);
        bool setMarginLeft(uint16_t newMarginLeft);
        bool setMarginTop(uint16_t newMarginTop);
        bool setWidth(uint16_t newWidth);
        bool setHeight(uint16_t newHeight);
        bool setMsoWrapStyle(std::string newMsoWrapStyle);
        bool setVTextAnchor(std::string newVTextAnchor);
        bool hide(); // set visibility:hidden
        bool show(); // set visibility:visible

        bool setRaw(std::string newStyle) { m_style = newStyle; return true; }

    private:
        mutable std::string m_style; // mutable so getter functions can update it from m_styleAttribute if the latter is not empty
        std::unique_ptr<XMLAttribute> m_styleAttribute;
        inline static const std::vector< std::string_view > m_nodeOrder = {
            "position",
            "margin-left",
            "margin-top",
            "width",
            "height",
            "mso-wrap-style",
            "v-text-anchor",
            "visibility"
        };
    };


    class OPENXLSX_EXPORT XLShape {
        friend class XLVmlDrawing;    // for access to m_shapeNode in XLVmlDrawing::addShape
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLShape();

        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the styles item. If no input is provided, a null node is used.
         */
        explicit XLShape(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLShape(const XLShape& other) = default;

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLShape(XLShape&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLShape() = default;

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLShape& operator=(const XLShape& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLShape& operator=(XLShape&& other) noexcept = default;

        /**
         * @brief Getter functions
         */
        std::string shapeId() const;   // v:shape attribute id - shape_# - can't be set by the user
        std::string fillColor() const; // v:shape attribute fillcolor, #<3 byte hex code>, e.g. #ffffc0
        bool stroked() const;          // v:shape attribute stroked "t" ("f"?)
        std::string type() const;      // v:shape attribute type, link to v:shapetype attribute id
        bool allowInCell() const;      // v:shape attribute o:allowincell "f"
        XLShapeStyle style();          // v:shape attribute style, but constructed from the XMLAttribute

        // XLShapeShadow& shadow();          // v:shape subnode v:shadow
        // XLShapeFill& fill();              // v:shape subnode v:fill
        // XLShapeStroke& stroke();          // v:shape subnode v:stroke
        // XLShapePath& path();              // v:shape subnode v:path
        // XLShapeTextbox& textbox();        // v:shape subnode v:textbox
        XLShapeClientData clientData();  // v:shape subnode x:ClientData

        /**
         * @brief Setter functions
         * @param value that shall be set
         * @return true for success, false for failure
         */
        // NOTE: setShapeId is not available because shape id is managed by the parent class in createShape
        bool setFillColor(std::string const& newFillColor);
        bool setStroked(bool set);
        bool setType(std::string const& newType);
        bool setAllowInCell(bool set);
        bool setStyle(std::string const& newStyle);
        bool setStyle(XLShapeStyle const& newStyle);

    private:                                         // ---------- Private Member Variables ---------- //
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

    /**
     * @brief The XLVmlDrawing class is the base class for worksheet comments
     */
    class OPENXLSX_EXPORT XLVmlDrawing : public XLXmlFile
    {
        friend class XLWorksheet;   // for access to XLXmlFile::getXmlPath
        friend class XLComments;    // for access to firstShapeNode
    public:
        /**
         * @brief Constructor
         */
        XLVmlDrawing() : XLXmlFile(nullptr) {};

        /**
         * @brief The constructor.
         * @param xmlData the source XML of the comments file
         */
        XLVmlDrawing(XLXmlData* xmlData);

        /**
         * @brief The copy constructor.
         * @param other The object to be copied.
         * @note The default copy constructor is used, i.e. only shallow copying of pointer data members.
         */
        XLVmlDrawing(const XLVmlDrawing& other) = default;

        /**
         * @brief
         * @param other
         */
        XLVmlDrawing(XLVmlDrawing&& other) noexcept = default;

        /**
         * @brief The destructor
         * @note The default destructor is used, since cleanup of pointer data members is not required.
         */
        ~XLVmlDrawing() = default;

        /**
         * @brief Assignment operator
         * @return A reference to the new object.
         * @note The default assignment operator is used, i.e. only shallow copying of pointer data members.
         */
        XLVmlDrawing& operator=(const XLVmlDrawing&) = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLVmlDrawing& operator=(XLVmlDrawing&& other) noexcept = default;

    private: // helper functions with repeating code
        XMLNode firstShapeNode() const;
        XMLNode lastShapeNode() const;
        XMLNode shapeNode(uint32_t index) const;

    public:

        /**
         * @brief Get the shape XML node that is associated with the cell indicated by cellRef
         * @param cellRef the reference to the cell for which a shape shall be found
         * @return the XMLNode that contains the desired shape, or an empty XMLNode if not found
         */
        XMLNode shapeNode(std::string const& cellRef) const;

        uint32_t shapeCount() const;

        XLShape shape(uint32_t index) const;

        bool deleteShape(uint32_t index);
        bool deleteShape(std::string const& cellRef);

        XLShape createShape(const XLShape& shapeTemplate = XLShape());

        /**
         * @brief Print the XML contents of this XLVmlDrawing instance using the underlying XMLNode print function
         */
        void print(std::basic_ostream<char>& ostr) const;

    private:
        uint32_t m_shapeCount{0};
        uint32_t m_lastAssignedShapeId{0};
        std::string m_defaultShapeTypeId{};
    };
}    // namespace OpenXLSX

#endif    // OPENXLSX_XLDRAWING_HPP

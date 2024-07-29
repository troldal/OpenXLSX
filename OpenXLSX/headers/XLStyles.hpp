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

#ifndef OPENXLSX_XLSTYLES_HPP
#define OPENXLSX_XLSTYLES_HPP

#pragma warning(push)
#pragma warning(disable : 4251)
#pragma warning(disable : 4275)

// ===== External Includes ===== //
#include <cstdint>   // uint32_t etc
#include <string>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLColor.hpp"
#include "XLXmlFile.hpp"
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    using XLStyleIndex = size_t; // custom data type for XLStyleIndex

    constexpr const uint32_t XLInvalidUInt16 = 0xffff;     // used to signal "value not defined" for uint16_t return types
    constexpr const uint32_t XLInvalidUInt32 = 0xffffffff; // used to signal "value not defined" for uint32_t return types

    constexpr const bool XLPermitXfID      = true;         // use with XLCellFormat constructor to enable xfId() getter and setXfId() setter

    constexpr const bool XLCreateIfMissing = true;         // use with XLCellFormat::alignment(XLCreateIfMissing)
    constexpr const bool XLDoNotCreate     = false;        // use with XLCellFormat::alignment(XLDoNotCreate)

    constexpr const char * XLDefaultStylesPrefix       = "\n\t";   // indentation to use for newly created root level style node tags
    constexpr const char * XLDefaultStyleEntriesPrefix = "\n\t\t"; // indentation to use for newly created style entry nodes

    constexpr const XLStyleIndex XLDefaultCellFormat = 0;          // default cell format index in xl/styles.xml:<styleSheet>:<cellXfs>

    // ===== As pugixml attributes are not guaranteed to support value range of XLStyleIndex, use 32 bit unsigned int
    constexpr const XLStyleIndex XLInvalidStyleIndex = XLInvalidUInt32;    // as a function return value, indicates no valid index

    constexpr const uint32_t XLDefaultFontSize       = 12;         //
    constexpr const char *   XLDefaultFontColor      = "ff000000"; // default font color
    constexpr const char *   XLDefaultFontColorTheme = "";         // TBD what this means / how it is used
    constexpr const char *   XLDefaultFontName       = "Arial";    //
    constexpr const uint32_t XLDefaultFontFamily     = 0;          // TBD what this means / how it is used
    constexpr const uint32_t XLDefaultFontCharset    = 1;          // TBD what this means / how it is used

    constexpr const char * XLDefaultFillType           = "patternFill"; // node name for the pattern description
    constexpr const char * XLDefaultFillPatternType    = "none";        // attribute patternType value
    constexpr const char * XLDefaultFillPatternFgColor = "ffffffff";    // child node fgcolor attribute rgb value
    constexpr const char * XLDefaultFillPatternBgColor = "ff000000";    // child node bgcolor attribute rgb value

    constexpr const char * XLDefaultLineStyle = ""; // empty string = line not set

    // forward declarations of all classes in this header
    class XLNumberFormat;
    class XLNumberFormats;
    class XLFont;
    class XLFonts;
    class XLFill;
    class XLFills;
    class XLLine;
    class XLBorder;
    class XLBorders;
    class XLAlignment;
    class XLCellFormat;
    class XLCellFormats;
    class XLCellStyle;
    class XLCellStyles;
    class XLStyles;

    // TODO: implement fill pattern constants & use thereof
    enum XLFillPattern: uint8_t {
        XLFillPatternNone = 0,      // "none"
        XLFillPatternSolid = 1      // "solid"
    };

    enum XLLineType: uint8_t {
        XLLineLeft     = 0,
        XLLineRight    = 1,
        XLLineTop      = 2,
        XLLineBottom   = 3,
        XLLineDiagonal = 4,
        XLLineInvalid  = 255
    };

    enum XLLineStyle : uint8_t {
        XLLineStyleNone          = 0,
        XLLineStyleThin          = 1 // TBD: all valid line styles
    };

    enum XLAlignmentStyle : uint8_t {
        XLAlignGeneral       = 0,
        XLAlignLeft          = 1,
        XLAlignRight         = 2,
        XLAlignCenter        = 3,
        XLAlignTop           = 4,
        XLAlignBottom        = 5,
        XLAlignUnknown       = 255 // TBD: all valid alignment styles
    };

    enum XLUnderlineStyle : uint8_t {
        XLUnderlineNone    = 0,
        XLUnderlineSingle  = 1,
        XLUnderlineDouble  = 2,
        XLUnderlineInvalid = 255
    };

    // ================================================================================
    // XLNumberFormats Class
    // ================================================================================
    
    /**
     * @brief An encapsulation of a number format (numFmt) item
     */
    class OPENXLSX_EXPORT XLNumberFormat
    {
        friend class XLNumberFormats;    // for access to m_numberFormatNode in XLNumberFormats::create
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLNumberFormat();

        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the styles item. If no input is provided, a null node is used.
         */
        explicit XLNumberFormat(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLNumberFormat(const XLNumberFormat& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLNumberFormat(XLNumberFormat&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLNumberFormat();

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLNumberFormat& operator=(const XLNumberFormat& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLNumberFormat& operator=(XLNumberFormat&& other) noexcept = default;

        /**
         * @brief Get the id of the number format
         * @return The id for this number format
         */
        uint32_t numberFormatId() const;

        /**
         * @brief Get the code of the number format
         * @return The format code for this number format
         */
        std::string formatCode() const;

        /**
         * @brief Setter functions for style parameters
         * @param value that shall be set
         * @return true for success, false for failure
         */
        bool setNumberFormatId(uint32_t newNumberFormatId);
        bool setFormatCode(std::string newFormatCode);

        /**
         * @brief Return a string summary of the number format
         * @return string with info about the number format object
         */
        std::string summary() const;

    private:                                         // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_numberFormatNode; /**< An XMLNode object with the number format item */
    };


    /**
     * @brief An encapsulation of the XLSX number formats (numFmts)
     */
    class OPENXLSX_EXPORT XLNumberFormats
    {
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLNumberFormats();

        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the styles item. If no input is provided, a null node is used.
         */
        explicit XLNumberFormats(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLNumberFormats(const XLNumberFormats& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLNumberFormats(XLNumberFormats&& other);

        /**
         * @brief
         */
        ~XLNumberFormats();

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLNumberFormats& operator=(const XLNumberFormats& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLNumberFormats& operator=(XLNumberFormats&& other) noexcept = default;

        /**
         * @brief Get the count of number formats
         * @return The amount of entries in the number formats
         */
        size_t count() const;

        /**
         * @brief Get the number format identified by index
         * @param index an array index within XLStyles::numberFormats()
         * @return An XLNumberFormat object
         * @throw XLException when index is out of m_numberFormats range
         */
        XLNumberFormat numberFormatByIndex(XLStyleIndex index) const;

        /**
         * @brief Operator overload: allow [] as shortcut access to numberFormatByIndex
         */
        XLNumberFormat operator[](XLStyleIndex index) const { return numberFormatByIndex(index); }

        /**
         * @brief Get the number format identified by numberFormatId
         * @param numberFormatId a numFmtId reference to a number format
         * @return An XLNumberFormat object
         * @throw XLException if no match for numberFormatId is found within m_numberFormats
         */
        XLNumberFormat numberFormatById(uint32_t numberFormatId) const;

        /**
         * @brief Get the numFmtId from the number format identified by index
         * @param index an array index within XLStyles::numberFormats()
         * @return the uint32_t numFmtId corresponding to index
         * @throw XLException when index is out of m_numberFormats range
         */
        uint32_t numberFormatIdFromIndex(XLStyleIndex index) const;

        /**
         * @brief Append a new XLNumberFormat, based on copyFrom, and return its index in numFmts node
         * @param copyFrom Can provide an XLNumberFormat to use as template for the new style
         * @param styleEntriesPrefix Prefix the newly created cell style XMLNode with this pugi::node_pcdata text
         * @returns The index of the new style as used by operator[]
         * @todo: TBD assign a unique, non-reserved uint32_t numFmtId. Alternatively, the user should configure it manually via setNumberFormatId
         * @todo: TBD implement a "getFreeNumberFormatId()" method that skips reserved identifiers and iterates over m_numberFormats to avoid
         *         all existing number format Ids.
         */
        XLStyleIndex create(XLNumberFormat copyFrom = XLNumberFormat{}, std::string styleEntriesPrefix = XLDefaultStyleEntriesPrefix);


    private:                                         // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_numberFormatsNode; /**< An XMLNode object with the number formats item */
        std::vector<XLNumberFormat> m_numberFormats;
    };


    // ================================================================================
    // XLFonts Class
    // ================================================================================
    
    /**
     * @brief An encapsulation of a font item
     */
    class OPENXLSX_EXPORT XLFont
    {
        friend class XLFonts;    // for access to m_fontNode in XLFonts::create
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLFont();

        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the fonts XMLNode. If no input is provided, a null node is used.
         */
        explicit XLFont(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLFont(const XLFont& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLFont(XLFont&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLFont();

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLFont& operator=(const XLFont& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLFont& operator=(XLFont&& other) noexcept = default;

        /**
         * @brief Get the font name
         * @return The font name
         */
        std::string fontName() const;

        /**
         * @brief Get the font size
         * @return The font size
         */
        size_t fontSize() const;

        /**
         * @brief Get the font bold status
         * @return true = bold, false = not bold
         */
        bool bold() const;

        /**
         * @brief Get the font italic (cursive) status
         * @return true = italic, false = not italice
         */
        bool italic() const;

        /**
         * @brief Get the font underline status
         * @return An XLUnderlineStyle value
         */
        XLUnderlineStyle underline() const;

        /**
         * @brief Get the font strikethrough status
         * @return true = strikethrough, false = no strikethrough
         */
        bool strikethrough() const;

        /**
         * @brief Get the font color
         * @return The font color
         */
        XLColor fontColor() const;

        /**
         * @brief Get the font family
         * @return The font family id
         */
        size_t fontFamily() const;

        /**
         * @brief Get the font charset
         * @return The font charset id
         */
        size_t fontCharset() const;

        /**
         * @brief Setter functions for style parameters
         * @param value that shall be set
         * @return true for success, false for failure
         */
        bool setFontName(std::string newName);
        bool setFontSize(size_t newSize);
        bool setBold(bool set = true);
        bool setItalic(bool set = true);
        bool setUnderline(XLUnderlineStyle style = XLUnderlineSingle);
        bool setStrikethrough(bool set = true);
        bool setFontColor(XLColor newColor);
        bool setFontFamily(size_t newFamily);
        bool setFontCharset(size_t newCharset);

        /**
         * @brief Return a string summary of the font properties
         * @return string with info about the font object
         */
        std::string summary() const;

    private:                                         // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_fontNode;         /**< An XMLNode object with the font item */
    };


    /**
     * @brief An encapsulation of the XLSX fonts
     */
    class OPENXLSX_EXPORT XLFonts
    {
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLFonts();

        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the styles item. If no input is provided, a null node is used.
         */
        explicit XLFonts(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLFonts(const XLFonts& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLFonts(XLFonts&& other);

        /**
         * @brief
         */
        ~XLFonts();

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLFonts& operator=(const XLFonts& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLFonts& operator=(XLFonts&& other) noexcept = default;

        /**
         * @brief Get the count of fonts
         * @return The amount of font entries
         */
        size_t count() const;

        /**
         * @brief Get the font identified by index
         * @param index The index within the XML sequence
         * @return An XLFont object
         */
        XLFont fontByIndex(XLStyleIndex index) const;

        /**
         * @brief Operator overload: allow [] as shortcut access to fontByIndex
         * @param index The index within the XML sequence
         * @return An XLFont object
         */
        XLFont operator[](XLStyleIndex index) const { return fontByIndex(index); }

        /**
         * @brief Append a new XLFont, based on copyFrom, and return its index in fonts node
         * @param copyFrom Can provide an XLFont to use as template for the new style
         * @param styleEntriesPrefix Prefix the newly created cell style XMLNode with this pugi::node_pcdata text
         * @returns The index of the new style as used by operator[]
         */
        XLStyleIndex create(XLFont copyFrom = XLFont{}, std::string styleEntriesPrefix = XLDefaultStyleEntriesPrefix);

    private:                                         // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_fontsNode;        /**< An XMLNode object with the fonts item */
        std::vector<XLFont> m_fonts;
    };


    // ================================================================================
    // XLFills Class
    // ================================================================================
    
    /**
     * @brief An encapsulation of a fill item
     */
    class OPENXLSX_EXPORT XLFill
    {
        friend class XLFills;    // for access to m_fillNode in XLFills::create
    public:                                          // ---------- Public Member Functions ----------- //
        /**
         * @brief
         */
        XLFill();

        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the fonts XMLNode. If no input is provided, a null node is used.
         */
        explicit XLFill(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLFill(const XLFill& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLFill(XLFill&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLFill();

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLFill& operator=(const XLFill& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLFill& operator=(XLFill&& other) noexcept = default;

        /**
         * @brief Get the name of the element describing a fill
         * @return The name of the first child element of the fill node
         */
        std::string fillType() const;

        /**
         * @brief Get the pattern type
         * @return The pattern type string
         */
        std::string patternType() const;

        /**
         * @brief Get the foreground color
         * @return The foreground color
         */
        XLColor color();

        /**
         * @brief Get the background color
         * @return The background color
         */
        XLColor backgroundColor();

    private:                                         // ----- switch to private for one function ----- //
        /**
         * @brief Fetch a valid first element child of m_fillNode. Create with default if needed
         * @return An XMLNode containing a fill description
         */
        XMLNode getValidFillDescription();
    public:                                          // ----- switch back to public functions -------- //
        /**
         * @brief Setter functions for style parameters
         * @param value that shall be set
         * @return true for success, false for failure
         */
        bool setFillType(std::string newFillType);
        bool setPatternType(std::string newPatternType);
        bool setColor(XLColor newColor);
        bool setBackgroundColor(XLColor newBgColor);

        /**
         * @brief Return a string summary of the fill properties
         * @return string with info about the fill object
         */
        std::string summary();

    private:                                         // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_fillNode;         /**< An XMLNode object with the fill item */
    };


    /**
     * @brief An encapsulation of the XLSX fills
     */
    class OPENXLSX_EXPORT XLFills
    {
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLFills();

        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the styles item. If no input is provided, a null node is used.
         */
        explicit XLFills(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLFills(const XLFills& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLFills(XLFills&& other);

        /**
         * @brief
         */
        ~XLFills();

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLFills& operator=(const XLFills& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLFills& operator=(XLFills&& other) noexcept = default;

        /**
         * @brief Get the count of fills
         * @return The amount of fill entries
         */
        size_t count() const;

        /**
         * @brief Get the fill entry identified by index
         * @param index The index within the XML sequence
         * @return An XLFill object
         */
        XLFill fillByIndex(XLStyleIndex index) const;

        /**
         * @brief Operator overload: allow [] as shortcut access to fillByIndex
         * @param index The index within the XML sequence
         * @return An XLFill object
         */
        XLFill operator[](XLStyleIndex index) const { return fillByIndex(index); }

        /**
         * @brief Append a new XLFill, based on copyFrom, and return its index in fills node
         * @param copyFrom Can provide an XLFill to use as template for the new style
         * @param styleEntriesPrefix Prefix the newly created cell style XMLNode with this pugi::node_pcdata text
         * @returns The index of the new style as used by operator[]
         */
        XLStyleIndex create(XLFill copyFrom = XLFill{}, std::string styleEntriesPrefix = XLDefaultStyleEntriesPrefix);

    private:                                         // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_fillsNode;        /**< An XMLNode object with the fills item */
        std::vector<XLFill> m_fills;
    };


    // ================================================================================
    // XLBorders Class
    // ================================================================================
    
    /**
     * @brief An encapsulation of a line item
     */
    class OPENXLSX_EXPORT XLLine
    {
        // friend class TBD: XLBorder or XLBorders;    // for access to m_lineNode in TBD
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLLine();
  
        /**
         * @brief Constructor. New items should only be created through an XLBorder object.
         * @param node An XMLNode object with the fonts XMLNode. If no input is provided, a null node is used.
         */
        explicit XLLine(const XMLNode& node);
  
        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLLine(const XLLine& other);
  
        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLLine(XLLine&& other) noexcept = default;
  
        /**
         * @brief
         */
        ~XLLine();
  
        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLLine& operator=(const XLLine& other);
  
        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLLine& operator=(XLLine&& other) noexcept = default;
  
        /**
         * @brief Get the line style
         * @return An XLLineStyle enum
         */
        XLLineStyle style() const;
  
        /**
         * @brief Evaluate XLLine as bool
         * @return true if line is set, false if not
         */
        explicit operator bool() const;
  
        /**
         * @brief Get the line color
         * @return An XLColor object
         */
        XLColor color() const;
  
        /**
         * @brief Setter functions for style parameters
         * @note: Please refer to XLBorder setLine / setLeft / setRight / setTop / setBottom / setDiagonal
         * @param value that shall be set
         * @return true for success, false for failure
         */
        // bool setStyle(XLLineStyle newStyle);
        // bool setColor(XLColor newColor);

        /**
         * @brief Return a string summary of the line properties
         * @return string with info about the line object
         */
        std::string summary() const;
  
    private:                                         // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_lineNode;         /**< An XMLNode object with the line item */
     };
  
  
    /**
     * @brief An encapsulation of a border item
     */
    class OPENXLSX_EXPORT XLBorder
    {
        friend class XLBorders;    // for access to m_borderNode in XLBorders::create
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLBorder();
  
        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the fonts XMLNode. If no input is provided, a null node is used.
         */
        explicit XLBorder(const XMLNode& node);
  
        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLBorder(const XLBorder& other);
  
        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLBorder(XLBorder&& other) noexcept = default;
  
        /**
         * @brief
         */
        ~XLBorder();
  
        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLBorder& operator=(const XLBorder& other);
  
        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLBorder& operator=(XLBorder&& other) noexcept = default;
  
        /**
         * @brief Get the diagonal up property
         * @return true if set, otherwise false
         */
        bool diagonalUp() const;
  
        /**
         * @brief Get the diagonal down property
         * @return true if set, otherwise false
         */
        bool diagonalDown() const;
  
        /**
         * @brief Get the left line property
         * @return An XLLine object
         */
        XLLine left() const;
  
        /**
         * @brief Get the left line property
         * @return An XLLine object
         */
        XLLine right() const;
  
        /**
         * @brief Get the left line property
         * @return An XLLine object
         */
        XLLine top() const;
  
        /**
         * @brief Get the bottom line property
         * @return An XLLine object
         */
        XLLine bottom() const;
  
        /**
         * @brief Get the diagonal line property
         * @return An XLLine object
         */
        XLLine diagonal() const;
  
        /**
         * @brief Setter functions for style parameters
         * @param value that shall be set
         * @param value2 (optional) that shall be set
         * @return true for success, false for failure
         */
        bool setDiagonalUp  (bool set);
        bool setDiagonalDown(bool set);
        bool setLine        (XLLineType lineType, XLLineStyle lineStyle, XLColor lineColor);
        bool setLeft        (                     XLLineStyle lineStyle, XLColor lineColor);
        bool setRight       (                     XLLineStyle lineStyle, XLColor lineColor);
        bool setTop         (                     XLLineStyle lineStyle, XLColor lineColor);
        bool setBottom      (                     XLLineStyle lineStyle, XLColor lineColor);
        bool setDiagonal    (                     XLLineStyle lineStyle, XLColor lineColor);

        /**
         * @brief Return a string summary of the font properties
         * @return string with info about the font object
         */
        std::string summary() const;
  
    private:                                         // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_borderNode;       /**< An XMLNode object with the font item */
    };
  
  
    /**
     * @brief An encapsulation of the XLSX borders
     */
    class OPENXLSX_EXPORT XLBorders
    {
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLBorders();
  
        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the styles item. If no input is provided, a null node is used.
         */
        explicit XLBorders(const XMLNode& node);
  
        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLBorders(const XLBorders& other);
  
        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLBorders(XLBorders&& other);
  
        /**
         * @brief
         */
        ~XLBorders();
  
        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLBorders& operator=(const XLBorders& other);
  
        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLBorders& operator=(XLBorders&& other) noexcept = default;
  
        /**
         * @brief Get the count of border descriptions
         * @return The amount of border description entries
         */
        size_t count() const;
  
        /**
         * @brief Get the border description identified by index
         * @param index The index within the XML sequence
         * @return An XLBorder object
         */
        XLBorder borderByIndex(XLStyleIndex index) const;
  
        /**
         * @brief Operator overload: allow [] as shortcut access to borderByIndex
         * @param index The index within the XML sequence
         * @return An XLBorder object
         */
        XLBorder operator[](XLStyleIndex index) const { return borderByIndex(index); }
  
        /**
         * @brief Append a new XLBorder, based on copyFrom, and return its index in borders node
         * @param copyFrom Can provide an XLBorder to use as template for the new style
         * @param styleEntriesPrefix Prefix the newly created cell style XMLNode with this pugi::node_pcdata text
         * @returns The index of the new style as used by operator[]
         */
        XLStyleIndex create(XLBorder copyFrom = XLBorder{}, std::string styleEntriesPrefix = XLDefaultStyleEntriesPrefix);

    private:                                         // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_bordersNode;      /**< An XMLNode object with the borders item */
        std::vector<XLBorder> m_borders;
    };


    // ================================================================================
    // XLCellFormats Class
    // ================================================================================

    /**
     * @brief An encapsulation of an alignment item
     */
    class OPENXLSX_EXPORT XLAlignment
    {
        // friend class TBD: XLCellFormat or XLCellFormats;    // for access to m_alignmentNode in TBD
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLAlignment();

        /**
         * @brief Constructor. New items should only be created through an XLBorder object.
         * @param node An XMLNode object with the fonts XMLNode. If no input is provided, a null node is used.
         */
        explicit XLAlignment(const XMLNode& node);
  
        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLAlignment(const XLAlignment& other);
  
        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLAlignment(XLAlignment&& other) noexcept = default;
  
        /**
         * @brief
         */
        ~XLAlignment();
  
        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLAlignment& operator=(const XLAlignment& other);
  
        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLAlignment& operator=(XLAlignment&& other) noexcept = default;
  
        /**
         * @brief Get the horizontal alignment
         * @return An XLAlignmentStyle
         */
        XLAlignmentStyle horizontal() const;

        /**
         * @brief Get the vertical alignment
         * @return An XLAlignmentStyle
         */
        XLAlignmentStyle vertical() const;

        /**
         * @brief Get the text rotation
         * @return A value in degrees (TBD: clockwise? counter-clockwise?)
         */
        uint16_t textRotation() const;

        /**
         * @brief Check whether text wrapping is enabled
         * @return true if enabled, false otherwise
         */
        bool wrapText() const;

        /**
         * @brief Get the indent setting
         * @return TBD: what is the indent / what's the unit?
         */
        uint32_t indent() const;

        /**
         * @brief Check whether shrink to fit is enabled
         * @return true if enabled, false otherwise
         */
        bool shrinkToFit() const;

        /**
         * @brief Setter functions for style parameters
         * @param value that shall be set
         * @return true for success, false for failure
         */
        bool setHorizontal(XLAlignmentStyle newStyle);
        bool setVertical(XLAlignmentStyle newStyle);
        bool setTextRotation(uint16_t newRotation);
        bool setWrapText(bool set);
        bool setIndent(uint32_t newIndent);
        bool setShrinkToFit(bool set);

        /**
         * @brief Return a string summary of the alignment properties
         * @return string with info about the alignment object
         */
        std::string summary() const;
  
    private:                                         // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_alignmentNode;    /**< An XMLNode object with the alignment item */
    };


    /**
     * @brief An encapsulation of a cell format item
     */
    class OPENXLSX_EXPORT XLCellFormat
    {
        friend class XLCellFormats;    // for access to m_cellFormatNode in XLCellFormats::create

    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLCellFormat();
  
        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the fonts XMLNode. If no input is provided, a null node is used.
         * @param permitXfId true (XLPermitXfID) -> getter xfId and setter setXfId are enabled, otherwise will throw XLException if invoked
         */
        explicit XLCellFormat(const XMLNode& node, bool permitXfId);
  
        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLCellFormat(const XLCellFormat& other);
  
        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLCellFormat(XLCellFormat&& other) noexcept = default;
  
        /**
         * @brief
         */
        ~XLCellFormat();
  
        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLCellFormat& operator=(const XLCellFormat& other);
  
        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLCellFormat& operator=(XLCellFormat&& other) noexcept = default;
  
        /**
         * @brief Get the number format id
         * @return The identifier of a number format, built-in (predefined by office) or defind in XLNumberFormats
         */
        uint32_t numberFormatId() const;
  
        /**
         * @brief Get the font index
         * @return The index(!) of a font as defined in XLFonts
         */
        XLStyleIndex fontIndex() const;

        /**
         * @brief Get the fill index
         * @return The index(!) of a fill as defined in XLFills
         */
        XLStyleIndex fillIndex() const;

        /**
         * @brief Get the border index
         * @return The index(!) of a border as defined in XLBorders
         */
        XLStyleIndex borderIndex() const;

        /**
         * @brief Report whether font is applied
         * @return true for a setting enabled, or false if disabled
         */
        bool applyFont() const;

        /**
         * @brief Report whether fill is applied
         * @return true for a setting enabled, or false if disabled
         */
        bool applyFill() const;

        /**
         * @brief Report whether border is applied
         * @return true for a setting enabled, or false if disabled
         */
        bool applyBorder() const;

        /**
         * @brief Report whether alignment is applied
         * @return true for a setting enabled, or false if disabled
         */
        bool applyAlignment() const;

        /**
         * @brief Report whether protection is applied
         * @return true for a setting enabled, or false if disabled
         */
        bool applyProtection() const;

        /**
         * @brief Report whether protection locked is applied
         * @return true for a setting enabled, or false if disabled
         */
        bool locked() const;

        /**
         * @brief Report whether protection hidden is applied
         * @return true for a setting enabled, or false if disabled
         */
        bool hidden() const;

        /**
         * @brief Return a reference to applicable alignment
         * @param createIfMissing triggers creation of alignment node - should be used with setter functions of XLAlignment
         * @return An XLAlignment object reference
         */
        XLAlignment alignment(bool createIfMissing = false) const;

        /**
         * @brief Get the id of a referred <xf> entry
         * @return The id referring to an index in cell style formats (cellStyleXfs)
         * @throw XLException when invoked from cellStyleFormats
         * @note - only permitted for cellFormats
         */
        XLStyleIndex xfId() const;

        /**
         * @brief Setter functions for style parameters
         * @param value that shall be set
         * @return true for success, false for failure
         */
        bool setNumberFormatId  (uint32_t newNumFmtId);
        bool setFontIndex       (XLStyleIndex newFontIndex);
        bool setFillIndex       (XLStyleIndex newFillIndex);
        bool setBorderIndex     (XLStyleIndex newBorderIndex);
        bool setApplyFont       (bool set);
        bool setApplyFill       (bool set);
        bool setApplyBorder     (bool set);
        bool setApplyAlignment  (bool set);
        bool setApplyProtection (bool set);
        bool setLocked          (bool set);
        bool setHidden          (bool set);
        bool setXfId            (XLStyleIndex newXfId); // NOTE: throws when invoked from cellStyleFormats

        /**
         * @brief Return a string summary of the cell format properties
         * @return string with info about the cell format object
         */
        std::string summary() const;
  
    private:                                         // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_cellFormatNode;   /**< An XMLNode object with the cell format (xf) item */
        bool m_permitXfId{false};
    };


    /**
     * @brief An encapsulation of the XLSX cell style formats
     */
    class OPENXLSX_EXPORT XLCellFormats
    {
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLCellFormats();
  
        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the styles item. If no input is provided, a null node is used.
         * @param permitXfId Pass-through to XLCellFormat constructor: true (XLPermitXfID) -> setter setXfId is enabled, otherwise throws
         */
        explicit XLCellFormats(const XMLNode& node, bool permitXfId = false);
  
        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLCellFormats(const XLCellFormats& other);
  
        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLCellFormats(XLCellFormats&& other);
  
        /**
         * @brief
         */
        ~XLCellFormats();
  
        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLCellFormats& operator=(const XLCellFormats& other);
  
        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLCellFormats& operator=(XLCellFormats&& other) noexcept = default;
  
        /**
         * @brief Get the count of cell style format descriptions
         * @return The amount of cell style format description entries
         */
        size_t count() const;
  
        /**
         * @brief Get the cell style format description identified by index
         * @param index The index within the XML sequence
         * @return An XLCellFormat object
         */
        XLCellFormat cellFormatByIndex(XLStyleIndex index) const;
  
        /**
         * @brief Operator overload: allow [] as shortcut access to cellFormatByIndex
         * @param index The index within the XML sequence
         * @return An XLCellFormat object
         */
        XLCellFormat operator[](XLStyleIndex index) const { return cellFormatByIndex(index); }
  
        /**
         * @brief Append a new XLCellFormat, based on copyFrom, and return its index
         *       in cellXfs (for XLStyles::cellFormats) or cellStyleXfs (for XLStyles::cellStyleFormats)
         * @param copyFrom Can provide an XLCellFormat to use as template for the new style
         * @param styleEntriesPrefix Prefix the newly created cell style XMLNode with this pugi::node_pcdata text
         * @returns The index of the new style as used by operator[]
         */
        XLStyleIndex create(XLCellFormat copyFrom = XLCellFormat{}, std::string styleEntriesPrefix = XLDefaultStyleEntriesPrefix);

    private:                                         // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_cellFormatsNode;  /**< An XMLNode object with the cell formats item */
        std::vector<XLCellFormat> m_cellFormats;
        bool m_permitXfId{false};
    };


    // ================================================================================
    // XLCellStyles Class
    // ================================================================================
    
    /**
     * @brief An encapsulation of a cell style item
     */
    class OPENXLSX_EXPORT XLCellStyle
    {
        friend class XLCellStyles;    // for access to m_cellStyleNode in XLCellStyles::create

    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLCellStyle();

        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the styles item. If no input is provided, a null node is used.
         */
        explicit XLCellStyle(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLCellStyle(const XLCellStyle& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLCellStyle(XLCellStyle&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLCellStyle();

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLCellStyle& operator=(const XLCellStyle& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLCellStyle& operator=(XLCellStyle&& other) noexcept = default;

        /**
         * @brief Test if this is an empty node
         * @return true if underlying XMLNode is empty
         */
        bool empty() const;

        /**
         * @brief Get the name of the cell style
         * @return The name for this cell style entry
         */
        std::string name() const;

        /**
         * @brief Get the id of the cell style format
         * @return The id referring to an index in cell style formats (cellStyleXfs) - TBD to be confirmed
         */
        XLStyleIndex xfId() const;

        /**
         * @brief Get the built-in id of the cell style
         * @return The built-in id of the cell style - TBD what this means
         */
        uint32_t builtinId() const;

        /**
         * @brief Setter functions for style parameters
         * @param value that shall be set
         * @return true for success, false for failure
         */
        bool setName     (std::string newName);
        bool setXfId     (XLStyleIndex newXfId);
        bool setBuiltinId(uint32_t newBuiltinId);

        /**
         * @brief Return a string summary of the cell style properties
         * @return string with info about the cell style object
         */
        std::string summary() const;
  
    private:                                         // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_cellStyleNode;    /**< An XMLNode object with the cell style item */
    };


    /**
     * @brief An encapsulation of the XLSX cell styles
     */
    class OPENXLSX_EXPORT XLCellStyles
    {
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLCellStyles();

        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the styles item. If no input is provided, a null node is used.
         */
        explicit XLCellStyles(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLCellStyles(const XLCellStyles& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLCellStyles(XLCellStyles&& other);

        /**
         * @brief
         */
        ~XLCellStyles();

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLCellStyles& operator=(const XLCellStyles& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLCellStyles& operator=(XLCellStyles&& other) noexcept = default;

        /**
         * @brief Get the count of cell styles
         * @return The amount of entries in the cell styles
         */
        size_t count() const;

        /**
         * @brief Get the cell style identified by index
         * @return An XLCellStyle object
         */
        XLCellStyle cellStyleByIndex(XLStyleIndex index) const;

        /**
         * @brief Operator overload: allow [] as shortcut access to cellStyleByIndex
         * @param index The index within the XML sequence
         * @return An XLCellStyle object
         */
        XLCellStyle operator[](XLStyleIndex index) const { return cellStyleByIndex(index); }

        /**
         * @brief Append a new XLCellStyle, based on copyFrom, and return its index in cellStyles node
         * @param copyFrom Can provide an XLCellStyle to use as template for the new style
         * @param styleEntriesPrefix Prefix the newly created cell style XMLNode with this pugi::node_pcdata text
         * @returns The index of the new style as used by operator[]
         */
        XLStyleIndex create(XLCellStyle copyFrom = XLCellStyle{}, std::string styleEntriesPrefix = XLDefaultStyleEntriesPrefix);

    private:                                         // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_cellStylesNode;   /**< An XMLNode object with the cell styles item */
        std::vector<XLCellStyle> m_cellStyles;
    };


    // ================================================================================
    // XLStyles Class
    // ================================================================================
    
    /**
     * @brief An encapsulation of the styles file (xl/styles.xml) in an Excel document package.
     */
    class OPENXLSX_EXPORT XLStyles : public XLXmlFile
    {
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLStyles() = default;
    
        /**
         * @brief
         * @param xmlData
         * @param stylesPrefix Prefix any newly created root style nodes with this text as pugi::node_pcdata
         */
        explicit XLStyles(XLXmlData* xmlData, std::string stylesPrefix = XLDefaultStylesPrefix);
    
        /**
         * @brief Destructor
         */
        ~XLStyles();
    
        /**
         * @brief
         * @param other
         */
        XLStyles(const XLStyles& other) = default;
    
        /**
         * @brief
         * @param other
         */
        XLStyles(XLStyles&& other) noexcept = default;
    
        /**
         * @brief
         * @param other
         * @return
         */
        XLStyles& operator=(const XLStyles& other) = default;
    
        /**
         * @brief
         * @param other
         * @return
         */
        XLStyles& operator=(XLStyles&& other) noexcept = default;
    
        /**
         * @brief Get the number formats object
         * @return An XLNumberFormats object
         */
        XLNumberFormats& numberFormats() const;

        /**
         * @brief Get the fonts object
         * @return An XLFonts object
         */
        XLFonts& fonts() const;

        /**
         * @brief Get the fills object
         * @return An XLFills object
         */
        XLFills& fills() const;

        /**
         * @brief Get the borders object
         * @return An XLBorders object
         */
        XLBorders& borders() const;

        /**
         * @brief Get the cell style formats object
         * @return An XLCellFormats object
         */
        XLCellFormats& cellStyleFormats() const;

        /**
         * @brief Get the cell formats object
         * @return An XLCellFormats object
         */
        XLCellFormats& cellFormats() const;

        /**
         * @brief Get the cell styles object
         * @return An XLCellStyles object
         */
        XLCellStyles& cellStyles() const;

        // ---------- Protected Member Functions ---------- //
    private:
        std::unique_ptr<XLNumberFormats>    m_numberFormats;    // handle to the underlying number formats
        std::unique_ptr<XLFonts>            m_fonts;            // handle to the underlying fonts
        std::unique_ptr<XLFills>            m_fills;            // handle to the underlying fills
        std::unique_ptr<XLBorders>          m_borders;          // handle to the underlying border descriptions
        std::unique_ptr<XLCellFormats>      m_cellStyleFormats; // handle to the underlying cell style formats descriptions
        std::unique_ptr<XLCellFormats>      m_cellFormats;      // handle to the underlying cell formats descriptions
        std::unique_ptr<XLCellStyles>       m_cellStyles;       // handle to the underlying cell styles
    };
}    // namespace OpenXLSX

#pragma warning(pop)
#endif    // OPENXLSX_XLSTYLES_HPP

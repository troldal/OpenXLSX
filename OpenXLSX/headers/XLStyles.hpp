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

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#   pragma warning(push)
#   pragma warning(disable : 4251)
#   pragma warning(disable : 4275)
#endif // _MSC_VER

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

    constexpr const bool XLForceFillType   = true;

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

    constexpr const char * XLDefaultLineStyle = ""; // empty string = line not set

    // forward declarations of all classes in this header
    class XLUnsupportedElement; // used for stub getter / setter functions for unsupported style options
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

    enum XLUnderlineStyle : uint8_t {
        XLUnderlineNone    = 0,
        XLUnderlineSingle  = 1,
        XLUnderlineDouble  = 2,
        XLUnderlineInvalid = 255
    };

    enum XLFontSchemeStyle : uint8_t {
        XLFontSchemeNone    =   0, // <scheme val="none"/>
        XLFontSchemeMajor   =   1, // <scheme val="major"/>
        XLFontSchemeMinor   =   2, // <scheme val="minor"/>
        XLFontSchemeInvalid = 255  // all other values
    };

    enum XLVerticalAlignRunStyle : uint8_t {
        XLBaseline                =   0, // <vertAlign val="baseline"/>
        XLSubscript               =   1, // <vertAlign val="subscript"/>
        XLSuperscript             =   2, // <vertAlign val="superscript"/>
        XLVerticalAlignRunInvalid = 255
    };

    enum XLFillType : uint8_t {
        XLGradientFill     =   0,    // <gradientFill />
        XLPatternFill      =   1,    // <patternFill />
        XLFillTypeInvalid  = 255,    // any child of <fill> that is not one of the above
    };

    enum XLGradientType : uint8_t {
        XLGradientLinear      =   0,
        XLGradientPath        =   1,
        XLGradientTypeInvalid = 255
    };

    enum XLPatternType: uint8_t {
        XLPatternNone            =   0,   // "none"
        XLPatternSolid           =   1,   // "solid"
        XLPatternMediumGray      =   2,   // "mediumGray"
        XLPatternDarkGray        =   3,   // "darkGray"
        XLPatternLightGray       =   4,   // "lightGray"
        XLPatternDarkHorizontal  =   5,   // "darkHorizontal"
        XLPatternDarkVertical    =   6,   // "darkVertical"
        XLPatternDarkDown        =   7,   // "darkDown"
        XLPatternDarkUp          =   8,   // "darkUp"
        XLPatternDarkGrid        =   9,   // "darkGrid"
        XLPatternDarkTrellis     =  10,   // "darkTrellis"
        XLPatternLightHorizontal =  11,   // "lightHorizontal"
        XLPatternLightVertical   =  12,   // "lightVertical"
        XLPatternLightDown       =  13,   // "lightDown"
        XLPatternLightUp         =  14,   // "lightUp"
        XLPatternLightGrid       =  15,   // "lightGrid"
        XLPatternLightTrellis    =  16,   // "lightTrellis"
        XLPatternGray125         =  17,   // "gray125"
        XLPatternGray0625        =  18,   // "gray0625"
        XLPatternTypeInvalid     = 255    // any patternType that is not one of the above
    };
    constexpr const XLFillType    XLDefaultFillType       = XLPatternFill; // node name for the pattern description is derived from this
    constexpr const XLPatternType XLDefaultPatternType    = XLPatternNone; // attribute patternType default value: no fill
    constexpr const char*         XLDefaultPatternFgColor = "ffffffff";    // child node fgcolor attribute rgb value
    constexpr const char*         XLDefaultPatternBgColor = "ff000000";    // child node bgcolor attribute rgb value


    enum XLLineType: uint8_t {
        XLLineLeft       =   0,
        XLLineRight      =   1,
        XLLineTop        =   2,
        XLLineBottom     =   3,
        XLLineDiagonal   =   4,
        XLLineVertical   =   5,
        XLLineHorizontal =   6,
        XLLineInvalid    = 255
    };

    enum XLLineStyle : uint8_t {
        XLLineStyleNone             =   0,
        XLLineStyleThin             =   1,
        XLLineStyleMedium           =   2,
        XLLineStyleDashed           =   3,
        XLLineStyleDotted           =   4,
        XLLineStyleThick            =   5,
        XLLineStyleDouble           =   6,
        XLLineStyleHair             =   7,
        XLLineStyleMediumDashed     =   8,
        XLLineStyleDashDot          =   9,
        XLLineStyleMediumDashDot    =  10,
        XLLineStyleDashDotDot       =  11,
        XLLineStyleMediumDashDotDot =  12,
        XLLineStyleSlantDashDot     =  13,
        XLLineStyleInvalid          = 255
    };

    enum XLAlignmentStyle : uint8_t {
        XLAlignGeneral          =   0, // value="general",          horizontal only
        XLAlignLeft             =   1, // value="left",             horizontal only
        XLAlignRight            =   2, // value="right",            horizontal only
        XLAlignCenter           =   3, // value="center",           both
        XLAlignTop              =   4, // value="top",              vertical only
        XLAlignBottom           =   5, // value="bottom",           vertical only
        XLAlignFill             =   6, // value="fill",             horizontal only
        XLAlignJustify          =   7, // value="justify",          both
        XLAlignCenterContinuous =   8, // value="centerContinuous", horizontal only
        XLAlignDistributed      =   9, // value="distributed",      both
        XLAlignInvalid          = 255  // all other values
    };

    enum XLReadingOrder : uint32_t {
        XLReadingOrderContextual  = 0,
        XLReadingOrderLeftToRight = 1,
        XLReadingOrderRightToLeft = 2
    };

    // ================================================================================
    // XLUnsupportedElement Class
    // ================================================================================
    class OPENXLSX_EXPORT XLUnsupportedElement
    {
    public:
        XLUnsupportedElement() {}
        bool empty() { return true; }
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
         * @brief Get the font charset
         * @return The font charset id
         */
        size_t fontCharset() const;

        /**
         * @brief Get the font family
         * @return The font family id
         */
        size_t fontFamily() const;

        /**
         * @brief Get the font size
         * @return The font size
         */
        size_t fontSize() const;

        /**
         * @brief Get the font color
         * @return The font color
         */
        XLColor fontColor() const;

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
         * @brief Get the font strikethrough status
         * @return true = strikethrough, false = no strikethrough
         */
        bool strikethrough() const;

        /**
         * @brief Get the font underline status
         * @return An XLUnderlineStyle value
         */
        XLUnderlineStyle underline() const;

        /**
         * @brief get the font scheme: none, major or minor
         * @return An XLFontSchemeStyle
         */
        XLFontSchemeStyle scheme() const;

        /**
         * @brief get the font vertical alignment run style: baseline, subscript or superscript
         * @return An XLVerticalAlignRunStyle
         */
        XLVerticalAlignRunStyle vertAlign() const;

        /**
         * @brief Get the outline status
         * @return a TBD bool
         * @todo need to find a use case for this
         */
        bool outline() const;

        /**
         * @brief Get the shadow status
         * @return a TBD bool
         * @todo need to find a use case for this
         */
        bool shadow() const;

        /**
         * @brief Get the condense status
         * @return a TBD bool
         * @todo need to find a use case for this
         */
        bool condense() const;

        /**
         * @brief Get the extend status
         * @return a TBD bool
         * @todo need to find a use case for this
         */
        bool extend() const;


        /**
         * @brief Setter functions for style parameters
         * @param value that shall be set
         * @return true for success, false for failure
         */
        bool setFontName(std::string newName);
        bool setFontCharset(size_t newCharset);
        bool setFontFamily(size_t newFamily);
        bool setFontSize(size_t newSize);
        bool setFontColor(XLColor newColor);
        bool setBold(bool set = true);
        bool setItalic(bool set = true);
        bool setStrikethrough(bool set = true);
        bool setUnderline(XLUnderlineStyle style = XLUnderlineSingle);
        bool setScheme(XLFontSchemeStyle newScheme);
        bool setVertAlign(XLVerticalAlignRunStyle newVertAlign);
        bool setOutline(bool set = true);
        bool setShadow(bool set = true);
        bool setCondense(bool set = true);
        bool setExtend(bool set = true);

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
    // XLDataBarColor Class
    // ================================================================================

    /**
     * @brief An encapsulation of an XLSX Data Bar Color (CT_Color) item
     */
    class OPENXLSX_EXPORT XLDataBarColor
    {
    public:
        /**
         * @brief
         */
        XLDataBarColor();

        /**
         * @brief Constructor. New items should only be created through an XLGradientStop or XLLine object.
         * @param node An XMLNode object with a data bar color XMLNode. If no input is provided, a null node is used.
         */
        explicit XLDataBarColor(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLDataBarColor(const XLDataBarColor& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLDataBarColor(XLDataBarColor&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLDataBarColor() = default;

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLDataBarColor& operator=(const XLDataBarColor& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLDataBarColor& operator=(XLDataBarColor&& other) noexcept = default;

        /**
         * @brief Get the line color from the rgb attribute
         * @return An XLColor object
         */
        XLColor rgb() const;
  
        /**
         * @brief Get the line color tint
         * @return A double value as stored in the "tint" attribute (should be between [-1.0;+1.0]), 0.0 if attribute does not exist
         */
        double tint() const;

        /**
         * @brief currently unsupported getter stubs
         */
        bool     automatic() const; // <color auto="true" />
        uint32_t indexed()   const; // <color indexed="1" />
        uint32_t theme()     const; // <color theme="1" />

        /**
         * @brief Setter functions for data bar color parameters
         * @param value that shall be set
         * @return true for success, false for failure
         */
        bool setRgb      (XLColor newColor);
        bool set         (XLColor newColor) { return setRgb(newColor); } // alias for setRgb
        bool setTint     (double newTint);
        bool setAutomatic(bool set = true);
        bool setIndexed  (uint32_t newIndex);
        bool setTheme    (uint32_t newTheme);

        /**
         * @brief Return a string summary of the color properties
         * @return string with info about the color object
         */
        std::string summary() const;

    private:                                         // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_colorNode;        /**< An XMLNode object with the color item */
    };


    // ================================================================================
    // XLFills Class
    // ================================================================================

    /**
     * @brief An encapsulation of an fill::gradientFill::stop item
     */
    class OPENXLSX_EXPORT XLGradientStop
    {
        friend class XLGradientStops;    // for access to m_stopNode in XLGradientStops::create
    public:                                          // ---------- Public Member Functions ----------- //
        /**
         * @brief
         */
        XLGradientStop();

        /**
         * @brief Constructor. New items should only be created through an XLGradientStops object.
         * @param node An XMLNode object with the gradient stop XMLNode. If no input is provided, a null node is used.
         */
        explicit XLGradientStop(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLGradientStop(const XLGradientStop& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLGradientStop(XLGradientStop&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLGradientStop() = default;

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLGradientStop& operator=(const XLGradientStop& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLGradientStop& operator=(XLGradientStop&& other) noexcept = default;

        /**
         * @brief Getter functions
         * @return The requested gradient stop property
         */
        XLDataBarColor color() const; // <stop><color /></stop>
        double position()      const; // <stop position="1.2" />

        /**
         * @brief Setter functions
         * @param value that shall be set
         * @return true for success, false for failure
         * @note for color setters, use the color() getter with the XLDataBarColor setter functions
         */
        bool setPosition(double newPosition);

        /**
         * @brief Return a string summary of the stop properties
         * @return string with info about the stop object
         */
        std::string summary() const;

    private:                                         // ---------- Private Member Variables ---------- //
        std::unique_ptr<XMLNode> m_stopNode;         /**< An XMLNode object with the stop item */
    };

    /**
     * @brief An encapsulation of an array of fill::gradientFill::stop items
     */
    class OPENXLSX_EXPORT XLGradientStops
    {
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
        XLGradientStops();

        /**
         * @brief Constructor. New items should only be created through an XLFill object.
         * @param node An XMLNode object with the gradientFill item. If no input is provided, a null node is used.
         */
        explicit XLGradientStops(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLGradientStops(const XLGradientStops& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLGradientStops(XLGradientStops&& other);

        /**
         * @brief
         */
        ~XLGradientStops();

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLGradientStops& operator=(const XLGradientStops& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLGradientStops& operator=(XLGradientStops&& other) noexcept = default;

        /**
         * @brief Get the count of gradient stops
         * @return The amount of stop entries
         */
        size_t count() const;

        /**
         * @brief Get the gradient stop entry identified by index
         * @param index The index within the XML sequence
         * @return An XLGradientStop object
         */
        XLGradientStop stopByIndex(XLStyleIndex index) const;

        /**
         * @brief Operator overload: allow [] as shortcut access to stopByIndex
         * @param index The index within the XML sequence
         * @return An XLGradientStop object
         */
        XLGradientStop operator[](XLStyleIndex index) const { return stopByIndex(index); }

        /**
         * @brief Append a new XLGradientStop, based on copyFrom, and return its index in fills node
         * @param copyFrom Can provide an XLGradientStop to use as template for the new style
         * @param stopEntriesPrefix Prefix the newly created stop XMLNode with this pugi::node_pcdata text
         * @returns The index of the new style as used by operator[]
         */
        XLStyleIndex create(XLGradientStop copyFrom = XLGradientStop{}, std::string styleEntriesPrefix = "");

        std::string summary() const;

    private:
        std::unique_ptr<XMLNode> m_gradientNode;        /**< An XMLNode object with the gradientFill item */
        std::vector<XLGradientStop> m_gradientStops;
    };

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
         * @param node An XMLNode object with the fill XMLNode. If no input is provided, a null node is used.
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
         * @return The XLFillType derived from the name of the first child element of the fill node
         */
        XLFillType fillType() const;

        /**
         * @brief Create & set the base XML element describing the fill
         * @param newFillType that shall be set
         * @param force erase an existing fillType() if not equal newFillType
         * @return true for success, false for failure
         */
        bool setFillType(XLFillType newFillType, bool force = false);

    private:                                         // ---------- Switch to private context for two Methods  --------- //
        /**
         * @brief Throw XLException if fillType() matches typeToThrowOn
         * @param typeToThrowOn throw on this
         * @param functionName include this (calling function name) in the exception
         */
        void throwOnFillType(XLFillType typeToThrowOn, const char *functionName) const;

        /**
         * @brief Fetch a valid first element child of m_fillNode. Create with default if needed
         * @param fillTypeIfEmpty if no conflicting fill type exists, create a node with this fill type
         * @param functionName include this (calling function name) in a potential exception
         * @return An XMLNode containing a fill description
         * @throw XLException if fillTypeIfEmpty is in conflict with a current fillType()
         */
        XMLNode getValidFillDescription(XLFillType fillTypeIfEmpty, const char *functionName);

    public:                                          // ---------- Switch back to public context ---------------------- //

        /**
         * @brief Getter functions for gradientFill - will throwOnFillType(XLPatternFill, __func__)
         * @return The requested gradientFill property
         */
        XLGradientType gradientType(); // <gradientFill type="path" />
        double degree();
        double left();
        double right();
        double top();
        double bottom();
        XLGradientStops stops();

        /**
         * @brief Getter functions for patternFill - will throwOnFillType(XLGradientFill, __func__)
         * @return The requested patternFill property
         */
        XLPatternType patternType();
        XLColor color();
        XLColor backgroundColor();

        /**
         * @brief Setter functions for gradientFill - will throwOnFillType(XLPatternFill, __func__)
         * @param value that shall be set
         * @return true for success, false for failure
         * @note for gradient stops, use the stops() getter with the XLGradientStops access functions (create, stopByIndex, [])
         *       and the XLGradientStop setter functions
         */
        bool setGradientType(XLGradientType newType);
        bool setDegree(double newDegree);
        bool setLeft(double newLeft);
        bool setRight(double newRight);
        bool setTop(double newTop);
        bool setBottom(double newBottom);

        /**
         * @brief Setter functions for patternFill - will throwOnFillType(XLGradientFill, __func__)
         * @param value that shall be set
         * @return true for success, false for failure
         */
        bool setPatternType(XLPatternType newPatternType);
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
         * @param node An XMLNode object with the fills item. If no input is provided, a null node is used.
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
         * @param node An XMLNode object with the line XMLNode. If no input is provided, a null node is used.
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
  
        XLDataBarColor color() const; // <line><color /></line> where node can be left, right, top, bottom, diagonal, vertical, horizontal

        /**
         * @note Regarding setter functions for style parameters:
         * @note Please refer to XLBorder setLine / setLeft / setRight / setTop / setBottom / setDiagonal
         * @note  and XLDataBarColor setter functions that can be invoked via color()
         */

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
         * @param node An XMLNode object with the border XMLNode. If no input is provided, a null node is used.
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
         * @brief Get the outline property
         * @return true if set, otherwise false
         */
        bool outline() const;
  
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
         * @brief Get the vertical line property
         * @return An XLLine object
         */
        XLLine vertical() const;
  
        /**
         * @brief Get the horizontal line property
         * @return An XLLine object
         */
        XLLine horizontal() const;
  
        /**
         * @brief Setter functions for style parameters
         * @param value that shall be set
         * @param value2 (optional) that shall be set
         * @return true for success, false for failure
         */
        bool setDiagonalUp  (bool set = true);
        bool setDiagonalDown(bool set = true);
        bool setOutline     (bool set = true);
        bool setLine        (XLLineType lineType, XLLineStyle lineStyle, XLColor lineColor, double lineTint = 0.0);
        bool setLeft        (                     XLLineStyle lineStyle, XLColor lineColor, double lineTint = 0.0);
        bool setRight       (                     XLLineStyle lineStyle, XLColor lineColor, double lineTint = 0.0);
        bool setTop         (                     XLLineStyle lineStyle, XLColor lineColor, double lineTint = 0.0);
        bool setBottom      (                     XLLineStyle lineStyle, XLColor lineColor, double lineTint = 0.0);
        bool setDiagonal    (                     XLLineStyle lineStyle, XLColor lineColor, double lineTint = 0.0);
        bool setVertical    (                     XLLineStyle lineStyle, XLColor lineColor, double lineTint = 0.0);
        bool setHorizontal  (                     XLLineStyle lineStyle, XLColor lineColor, double lineTint = 0.0);

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
         * @param node An XMLNode object with the borders item. If no input is provided, a null node is used.
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
         * @param node An XMLNode object with the alignment XMLNode. If no input is provided, a null node is used.
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
         * @return An integer value, where an increment of 1 represents 3 spaces.
         */
        uint32_t indent() const;

        /**
         * @brief Get the relative indent setting
         * @return An integer value, where an increment of 1 represents 1 space, in addition to indent()*3 spaces
         * @note can be negative
         */
        int32_t relativeIndent() const;

        /**
         * @brief Check whether justification of last line is enabled
         * @return true if enabled, false otherwise
         */
        bool justifyLastLine() const;

        /**
         * @brief Check whether shrink to fit is enabled
         * @return true if enabled, false otherwise
         */
        bool shrinkToFit() const;

        /**
         * @brief Get the reading order setting
         * @return An integer value: 0 - Context Dependent, 1 - Left-to-Right, 2 - Right-to-Left (any other value should be invalid)
         */
        uint32_t readingOrder() const;

        /**
         * @brief Setter functions for style parameters
         * @param value that shall be set
         * @return true for success, false for failure
         */
        bool setHorizontal     (XLAlignmentStyle newStyle);
        bool setVertical       (XLAlignmentStyle newStyle);
        /**
         * @details on setTextRotation from XLSX specification:
         * Text rotation in cells. Expressed in degrees. Values range from 0 to 180. The first letter of the text 
         *  is considered the center-point of the arc.
         * For 0 - 90, the value represents degrees above horizon. For 91-180 the degrees below the horizon is calculated as:
         * [degrees below horizon] = 90 - [newRotation].
         * Examples: setTextRotation( 45): / (text is formatted along a line from lower left to upper right)
         *           setTextRotation(135): \ (text is formatted along a line from upper left to lower right)
         */
        bool setTextRotation   (uint16_t newRotation);
        bool setWrapText       (bool set = true);
        bool setIndent         (uint32_t newIndent);
        bool setRelativeIndent (int32_t newIndent);
        bool setJustifyLastLine(bool set = true);
        bool setShrinkToFit    (bool set = true);
        bool setReadingOrder   (uint32_t newReadingOrder); // can be used with XLReadingOrderContextual, XLReadingOrderLeftToRight, XLReadingOrderRightToLeft

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
         * @param node An XMLNode object with the xf XMLNode. If no input is provided, a null node is used.
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
         * @brief Get the id of a referred <xf> entry
         * @return The id referring to an index in cell style formats (cellStyleXfs)
         * @throw XLException when invoked from cellStyleFormats
         * @note - only permitted for cellFormats
         */
        XLStyleIndex xfId() const;

        /**
         * @brief Report whether number format is applied
         * @return true for a setting enabled, or false if disabled
         */
        bool applyNumberFormat() const;

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
         * @brief Report whether quotePrefix is enabled
         * @return true for a setting enabled, or false if disabled
         * @note from documentation: A boolean value indicating whether the text string in a cell should be prefixed by a single quote mark
         *       (e.g., 'text). In these cases, the quote is not stored in the Shared Strings Part.
         */
        bool quotePrefix() const;

        /**
         * @brief Report whether pivot button is applied
         * @return true for a setting enabled, or false if disabled
         * @note from documentation: A boolean value indicating whether the cell rendering includes a pivot table dropdown button.
         * @todo need to find a use case for this
         */
        bool pivotButton() const;

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
         * @brief Unsupported getter
         */
        XLUnsupportedElement extLst() const { return XLUnsupportedElement{}; } // <xf><extLst>...</extLst></xf>

        /**
         * @brief Setter functions for style parameters
         * @param value that shall be set
         * @return true for success, false for failure
         */
        bool setNumberFormatId   (uint32_t newNumFmtId);
        bool setFontIndex        (XLStyleIndex newFontIndex);
        bool setFillIndex        (XLStyleIndex newFillIndex);
        bool setBorderIndex      (XLStyleIndex newBorderIndex);
        bool setXfId             (XLStyleIndex newXfId); // NOTE: throws when invoked from cellStyleFormats
        bool setApplyNumberFormat(bool set = true);
        bool setApplyFont        (bool set = true);
        bool setApplyFill        (bool set = true);
        bool setApplyBorder      (bool set = true);
        bool setApplyAlignment   (bool set = true);
        bool setApplyProtection  (bool set = true);
        bool setQuotePrefix      (bool set = true);
        bool setPivotButton      (bool set = true);
        bool setLocked           (bool set = true);
        bool setHidden           (bool set = true);
        /**
         * @brief Unsupported setter
         */
        bool setExtLst          (XLUnsupportedElement const& newExtLst);

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
         * @param node An XMLNode object with the cell formats (cellXfs or cellStyleXfs) item. If no input is provided, a null node is used.
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
         * @param node An XMLNode object with the cellStyle item. If no input is provided, a null node is used.
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
         * @return The built-in id of the cell style
         * @todo need to find a use case for this
         */
        uint32_t builtinId() const;

        /**
         * @brief Get the outline style id (attribute iLevel) of the cell style
         * @return The outline style id of the cell style
         * @todo need to find a use case for this
         */
        uint32_t outlineStyle() const;

        /**
         * @brief Get the hidden flag of the cell style
         * @return The hidden flag status (true: applications should not show this style)
         */
        bool hidden() const;

        /**
         * @brief Get the custom buildin flag
         * @return true if this cell style shall customize a built-in style
         */
        bool customBuiltin() const;

        /**
         * @brief Unsupported getter
         */
        XLUnsupportedElement extLst() const { return XLUnsupportedElement{}; } // <cellStyle><extLst>...</extLst></cellStyle>

        /**
         * @brief Setter functions for style parameters
         * @param value that shall be set
         * @return true for success, false for failure
         */
        bool setName         (std::string newName);
        bool setXfId         (XLStyleIndex newXfId);
        bool setBuiltinId    (uint32_t newBuiltinId);
        bool setOutlineStyle (uint32_t newOutlineStyle);
        bool setHidden       (bool set = true);
        bool setCustomBuiltin(bool set = true);
        /**
         * @brief Unsupported setter
         */
        bool setExtLst   (XLUnsupportedElement const& newExtLst);

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
         * @param node An XMLNode object with the cellStyles item. If no input is provided, a null node is used.
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
         * @param suppressWarnings if true (SUPRESS_WARNINGS), messages such as "XLStyles: Ignoring currently unsupported <dxfs> node" will be silenced
         * @param stylesPrefix Prefix any newly created root style nodes with this text as pugi::node_pcdata
         */
        explicit XLStyles(XLXmlData* xmlData, bool suppressWarnings = false, std::string stylesPrefix = XLDefaultStylesPrefix);

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
        bool                                m_suppressWarnings; // if true, will suppress output of warnings where supported
        std::unique_ptr<XLNumberFormats>    m_numberFormats;    // handle to the underlying number formats
        std::unique_ptr<XLFonts>            m_fonts;            // handle to the underlying fonts
        std::unique_ptr<XLFills>            m_fills;            // handle to the underlying fills
        std::unique_ptr<XLBorders>          m_borders;          // handle to the underlying border descriptions
        std::unique_ptr<XLCellFormats>      m_cellStyleFormats; // handle to the underlying cell style formats descriptions
        std::unique_ptr<XLCellFormats>      m_cellFormats;      // handle to the underlying cell formats descriptions
        std::unique_ptr<XLCellStyles>       m_cellStyles;       // handle to the underlying cell styles
    };
}    // namespace OpenXLSX

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#   pragma warning(pop)
#endif // _MSC_VER

#endif    // OPENXLSX_XLSTYLES_HPP

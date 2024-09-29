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
#include <cstdint>      // uint32_t
#include <iostream>     // std::cout, std::cerr
#include <memory>       // std::make_unique
#include <pugixml.hpp>
#include <stdexcept>    // std::invalid_argument
#include <string>       // std::stoi, std::literals::string_literals
#include <vector>       // std::vector

// ===== OpenXLSX Includes ===== //
#include "XLColor.hpp"
#include "XLDocument.hpp"
#include <XLException.hpp>
#include "utilities/XLUtilities.hpp"    // OpenXLSX::ignore
#include "XLStyles.hpp"

using namespace OpenXLSX;


namespace     // anonymous namespace for module local functions
{
    constexpr const bool XLRemoveAttributes = true;

    enum XLStylesEntryType : uint8_t {
        XLStylesNumberFormats    =   0,
        XLStylesFonts            =   1,
        XLStylesFills            =   2,
        XLStylesBorders          =   3,
        XLStylesCellStyleFormats =   4,
        XLStylesCellFormats      =   5,
        XLStylesCellStyles       =   6,
        XLStylesColors           =   7,
        XLStylesFormats          =   8,
        XLStylesTableStyles      =   9,
        XLStylesExtLst           =  10,
        XLStylesInvalid          = 255
    };

    XLStylesEntryType XLStylesEntryTypeFromString(std::string name)
    {
        if (name == "numFmts")      return XLStylesNumberFormats;
        if (name == "fonts")        return XLStylesFonts;
        if (name == "fills")        return XLStylesFills;
        if (name == "borders")      return XLStylesBorders;
        if (name == "cellStyleXfs") return XLStylesCellStyleFormats;
        if (name == "cellXfs")      return XLStylesCellFormats;
        if (name == "cellStyles")   return XLStylesCellStyles;
        if (name == "colors")       return XLStylesColors;
        if (name == "dxfs")         return XLStylesFormats;
        if (name == "tableStyles")  return XLStylesTableStyles;
        if (name == "extLst")       return XLStylesExtLst;
        return XLStylesEntryType::XLStylesInvalid;
    }

    std::string XLStylesEntryTypeToString(XLStylesEntryType entryType)
    {
        switch (entryType) {
            case XLStylesNumberFormats:    return "numFmts";
            case XLStylesFonts:            return "fonts";
            case XLStylesFills:            return "fills";
            case XLStylesBorders:          return "borders";
            case XLStylesCellStyleFormats: return "cellStyleXfs";
            case XLStylesCellFormats:      return "cellXfs";
            case XLStylesCellStyles:       return "cellStyles";
            case XLStylesColors:           return "colors";
            case XLStylesFormats:          return "dxfs";
            case XLStylesTableStyles:      return "tableStyles";
            case XLStylesExtLst:           return "extLst";
            case XLStylesInvalid:          [[fallthrough]];
            default:                       return "(invalid)";
        }
    }

    XLUnderlineStyle XLUnderlineStyleFromString(std::string underline)
    {
        if (underline == ""
         || underline == "none")   return XLUnderlineNone;
        if (underline == "single") return XLUnderlineSingle;
        if (underline == "double") return XLUnderlineDouble;
        std::cerr << __func__ << ": invalid underline style " << underline << std::endl;
        return XLUnderlineInvalid;
    }

    std::string XLUnderlineStyleToString(XLUnderlineStyle underline)
    {
        switch (underline) {
            case XLUnderlineNone   : return "";
            case XLUnderlineSingle : return "single";
            case XLUnderlineDouble : return "double";
            case XLUnderlineInvalid: [[fallthrough]];
            default: return "(invalid)";
        }
    }

    XLFontSchemeStyle XLFontSchemeStyleFromString(std::string fontScheme)
    {
        if (fontScheme == ""
         || fontScheme == "none")  return XLFontSchemeNone;
        if (fontScheme == "major") return XLFontSchemeMajor;
        if (fontScheme == "minor") return XLFontSchemeMinor;
        std::cerr << __func__ << ": invalid font scheme " << fontScheme << std::endl;
        return XLFontSchemeInvalid;
    }

    std::string XLFontSchemeStyleToString(XLFontSchemeStyle fontScheme)
    {
        switch (fontScheme) {
            case XLFontSchemeNone   : return "none";
            case XLFontSchemeMajor  : return "major";
            case XLFontSchemeMinor  : return "minor";
            case XLFontSchemeInvalid: [[fallthrough]];
            default                 : return "(invalid)";
        }
    }

    XLVerticalAlignRunStyle XLVerticalAlignRunStyleFromString(std::string vertAlign)
    {
        if (vertAlign == ""
         || vertAlign == "baseline")    return XLBaseline;
        if (vertAlign == "subscript")   return XLSubscript;
        if (vertAlign == "superscript") return XLSuperscript;
        std::cerr << __func__ << ": invalid font vertical align run style " << vertAlign << std::endl;
        return XLVerticalAlignRunInvalid;
    }

    std::string XLVerticalAlignRunStyleToString(XLVerticalAlignRunStyle vertAlign)
    {
        switch (vertAlign) {
            case XLBaseline               : return "baseline";
            case XLSubscript              : return "subscript";
            case XLSuperscript            : return "superscript";
            case XLVerticalAlignRunInvalid: [[fallthrough]];
            default                       : return "(invalid)";
        }
    }

    XLFillType XLFillTypeFromString(std::string fillType)
    {
        if (fillType == "gradientFill") return XLGradientFill;
        if (fillType == "patternFill" ) return XLPatternFill;
        if (fillType != "")    // suppress error message for empty fillType (= no node exists yet)
            std::cerr << __func__ << ": invalid fillType \"" << fillType << "\"" << std::endl;
        return XLFillTypeInvalid;
    }

    std::string XLFillTypeToString(XLFillType fillType)
    {
        switch (fillType) {
            case XLGradientFill   : return "gradientFill";
            case XLPatternFill    : return "patternFill";
            case XLFillTypeInvalid: [[fallthrough]];
            default:                return "(invalid)";
        }
    }

    XLGradientType XLGradientTypeFromString(std::string gradientType)
    {
        if (gradientType == "linear") return XLGradientLinear;
        if (gradientType == "path")   return XLGradientPath;
        std::cerr << __func__ << ": invalid gradient type " << gradientType << std::endl;
        return XLGradientTypeInvalid;
    }
    std::string XLGradientTypeToString(XLGradientType gradientType)
    {
        switch (gradientType) {
            case XLGradientLinear     : return "linear";
            case XLGradientPath       : return "path";
            case XLGradientTypeInvalid: [[fallthrough]];
            default:                    return "(invalid)";
        }
    }

    XLPatternType XLPatternTypeFromString(std::string patternType)
    {
        if (patternType == ""
         || patternType == "none")            return XLPatternNone;
        if (patternType == "solid")           return XLPatternSolid;
        if (patternType == "mediumGray")      return XLPatternMediumGray;
        if (patternType == "darkGray")        return XLPatternDarkGray;
        if (patternType == "lightGray")       return XLPatternLightGray;
        if (patternType == "darkHorizontal")  return XLPatternDarkHorizontal;
        if (patternType == "darkVertical")    return XLPatternDarkVertical;
        if (patternType == "darkDown")        return XLPatternDarkDown;
        if (patternType == "darkUp")          return XLPatternDarkUp;
        if (patternType == "darkGrid")        return XLPatternDarkGrid;
        if (patternType == "darkTrellis")     return XLPatternDarkTrellis;
        if (patternType == "lightHorizontal") return XLPatternLightHorizontal;
        if (patternType == "lightVertical")   return XLPatternLightVertical;
        if (patternType == "lightDown")       return XLPatternLightDown;
        if (patternType == "lightUp")         return XLPatternLightUp;
        if (patternType == "lightGrid")       return XLPatternLightGrid;
        if (patternType == "lightTrellis")    return XLPatternLightTrellis;
        if (patternType == "gray125")         return XLPatternGray125;
        if (patternType == "gray0625")        return XLPatternGray0625;
        std::cerr << __func__ << ": invalid patternType " << patternType << std::endl;
        return XLPatternTypeInvalid;
    }

    std::string XLPatternTypeToString(XLPatternType patternType)
    {
        switch (patternType) {
            case XLPatternNone           : return "none";
            case XLPatternSolid          : return "solid";
            case XLPatternMediumGray     : return "mediumGray";
            case XLPatternDarkGray       : return "darkGray";
            case XLPatternLightGray      : return "lightGray";
            case XLPatternDarkHorizontal : return "darkHorizontal";
            case XLPatternDarkVertical   : return "darkVertical";
            case XLPatternDarkDown       : return "darkDown";
            case XLPatternDarkUp         : return "darkUp";
            case XLPatternDarkGrid       : return "darkGrid";
            case XLPatternDarkTrellis    : return "darkTrellis";
            case XLPatternLightHorizontal: return "lightHorizontal";
            case XLPatternLightVertical  : return "lightVertical";
            case XLPatternLightDown      : return "lightDown";
            case XLPatternLightUp        : return "lightUp";
            case XLPatternLightGrid      : return "lightGrid";
            case XLPatternLightTrellis   : return "lightTrellis";
            case XLPatternGray125        : return "gray125";
            case XLPatternGray0625       : return "gray0625";
            case XLPatternTypeInvalid    : [[fallthrough]];
            default:                       return "(invalid)";
        }
    }

#pragma GCC diagnostic push
#pragma GCC diagnostic ignored "-Wunused-function"
    XLLineType XLLineTypeFromString(std::string lineType)
    {
        if (lineType == "left")       return XLLineLeft;
        if (lineType == "right")      return XLLineRight;
        if (lineType == "top")        return XLLineTop;
        if (lineType == "bottom")     return XLLineBottom;
        if (lineType == "diagonal")   return XLLineDiagonal;
        if (lineType == "vertical")   return XLLineVertical;
        if (lineType == "horizontal") return XLLineHorizontal;
        std::cerr << __func__ << ": invalid line type" << lineType << std::endl;
        return XLLineInvalid;
    }
#pragma GCC diagnostic pop

    std::string XLLineTypeToString(XLLineType lineType)
    {
        switch (lineType) {
            case XLLineLeft:       return "left";
            case XLLineRight:      return "right";
            case XLLineTop:        return "top";
            case XLLineBottom:     return "bottom";
            case XLLineDiagonal:   return "diagonal";
            case XLLineVertical:   return "vertical";
            case XLLineHorizontal: return "horizontal";
            case XLLineInvalid:    [[fallthrough]];
            default:               return "(invalid)";
        }
    }

    XLLineStyle XLLineStyleFromString(std::string style)
    {
        if (style == "")                 return XLLineStyleNone;
        if (style == "thin")             return XLLineStyleThin;
        if (style == "medium")           return XLLineStyleMedium;
        if (style == "dashed")           return XLLineStyleDashed;
        if (style == "dotted")           return XLLineStyleDotted;
        if (style == "thick")            return XLLineStyleThick;
        if (style == "double")           return XLLineStyleDouble;
        if (style == "hair")             return XLLineStyleHair;
        if (style == "mediumDashed")     return XLLineStyleMediumDashed;
        if (style == "dashDot")          return XLLineStyleDashDot;
        if (style == "mediumDashDot")    return XLLineStyleMediumDashDot;
        if (style == "dashDotDot")       return XLLineStyleDashDotDot;
        if (style == "mediumDashDotDot") return XLLineStyleMediumDashDotDot;
        if (style == "slantDashDot")     return XLLineStyleSlantDashDot;
        std::cerr << __func__ << ": invalid line style " << style << std::endl;
        return XLLineStyleInvalid;
    }

    std::string XLLineStyleToString(XLLineStyle style)
    {
        switch (style) {
            case XLLineStyleNone            : return "";
            case XLLineStyleThin            : return "thin";
            case XLLineStyleMedium          : return "medium";
            case XLLineStyleDashed          : return "dashed";
            case XLLineStyleDotted          : return "dotted";
            case XLLineStyleThick           : return "thick";
            case XLLineStyleDouble          : return "double";
            case XLLineStyleHair            : return "hair";
            case XLLineStyleMediumDashed    : return "mediumDashed";
            case XLLineStyleDashDot         : return "dashDot";
            case XLLineStyleMediumDashDot   : return "mediumDashDot";
            case XLLineStyleDashDotDot      : return "dashDotDot";
            case XLLineStyleMediumDashDotDot: return "mediumDashDotDot";
            case XLLineStyleSlantDashDot    : return "slantDashDot";
            case XLLineStyleInvalid         : [[fallthrough]];
            default                         : return "(invalid)";
        }
    }

    XLAlignmentStyle XLAlignmentStyleFromString(std::string alignment)
    {
        if (alignment == ""
         || alignment == "general")          return XLAlignGeneral;
        if (alignment == "left")             return XLAlignLeft;
        if (alignment == "right")            return XLAlignRight;
        if (alignment == "center")           return XLAlignCenter;
        if (alignment == "top")              return XLAlignTop;
        if (alignment == "bottom")           return XLAlignBottom;
        if (alignment == "fill")             return XLAlignFill;
        if (alignment == "justify")          return XLAlignJustify;
        if (alignment == "centerContinuous") return XLAlignCenterContinuous;
        if (alignment == "distributed")      return XLAlignDistributed;
        std::cerr << __func__ << ": invalid alignment style " << alignment << std::endl;
        return XLAlignInvalid;
    }

    std::string XLAlignmentStyleToString(XLAlignmentStyle alignment)
    {
        switch (alignment) {
            case XLAlignGeneral          : return "";
            case XLAlignLeft             : return "left";
            case XLAlignRight            : return "right";
            case XLAlignCenter           : return "center";
            case XLAlignTop              : return "top";
            case XLAlignBottom           : return "bottom";
            case XLAlignFill             : return "fill";
            case XLAlignJustify          : return "justify";
            case XLAlignCenterContinuous : return "centerContinuous";
            case XLAlignDistributed      : return "distributed";
            case XLAlignInvalid: [[fallthrough]];
            default: return "(unknown)";
        }
    }

    std::string XLReadingOrderToString(uint32_t readingOrder)
    {
        switch (readingOrder) {
            case XLReadingOrderContextual : return "contextual";
            case XLReadingOrderLeftToRight: return "left-to-right";
            case XLReadingOrderRightToLeft: return "right-to-left";
            default: return "(unknown)";
        }
    }

    XMLNode appendAndGetNode(XMLNode & parent, std::string nodeName)
    {
        if (parent.empty()) return XMLNode{};
        XMLNode node = parent.child(nodeName.c_str());
        if (node.empty()) node = parent.append_child(nodeName.c_str());
        return node;
    }

    XMLAttribute appendAndGetAttribute(XMLNode & node, std::string attrName, std::string attrDefaultVal)
    {
        if (node.empty()) return XMLAttribute{};
        XMLAttribute attr = node.attribute(attrName.c_str());
        if (attr.empty() && not node.empty()) {
            attr = node.append_attribute(attrName.c_str());
            attr.set_value(attrDefaultVal.c_str());
        }
        return attr;
    }

    XMLAttribute appendAndSetAttribute(XMLNode & node, std::string attrName, std::string attrVal)
    {
        if (node.empty()) return XMLAttribute{};
        XMLAttribute attr = node.attribute(attrName.c_str());
        if (attr.empty() && not node.empty())
            attr = node.append_attribute(attrName.c_str());
        attr.set_value(attrVal.c_str()); // silently fails on empty attribute, which is intended here
        return attr;
    }

    XMLAttribute appendAndGetNodeAttribute(XMLNode & parent, std::string nodeName, std::string attrName, std::string attrDefaultVal)
    {
        if (parent.empty()) return XMLAttribute{};
        XMLNode node = appendAndGetNode(parent, nodeName);
        return appendAndGetAttribute(node, attrName, attrDefaultVal);
    }

    XMLAttribute appendAndSetNodeAttribute(XMLNode & parent, std::string nodeName, std::string attrName, std::string attrVal, bool removeAttributes = false)
    {
        if (parent.empty()) return XMLAttribute{};
        XMLNode node = appendAndGetNode(parent, nodeName);
        if (removeAttributes) node.remove_attributes();
        return appendAndSetAttribute(node, attrName, attrVal);
    }

    void copyXMLNode(XMLNode & destination, XMLNode & source)
    {
        if (not source.empty()) {
            // ===== Copy all XML child nodes
            for (XMLNode child = source.first_child(); not child.empty(); child = child.next_sibling())
                destination.append_copy(child);
            // ===== Copy all XML attributes
            for (XMLAttribute attr = source.first_attribute(); not attr.empty(); attr = attr.next_attribute())
                destination.append_copy(attr);
        }
    }

    void wrapNode(XMLNode parentNode, XMLNode & node, std::string prefix)
    {
        if (not node.empty() && prefix.length() > 0) {
            parentNode.insert_child_before(pugi::node_pcdata, node).set_value(prefix.c_str());    // insert prefix before node opening tag
            node.append_child(pugi::node_pcdata).set_value(prefix.c_str());                       // insert prefix before node closing tag (within node)
        }
    }

   /**
     * @brief Format val as a string with decimalPlaces
     * @param val The value to format
     * @param decimalPlaces The amount of digits following the decimal separator that shall be included in the formatted string
     * @return The value formatted as a string with the desired amount of decimal places
     */
    std::string formatDoubleAsString(double val, int decimalPlaces = 2)
    {
        std::string result = std::to_string(val);
        size_t decimalPos = result.find_first_of('.');
        assert(decimalPos < result.length()); // this should never throw
    
        // ===== Return the string representation of val with the decimal separator and decimalPlaces digits following
        return result.substr(0, decimalPos + 1 + decimalPlaces);
    }

    /**
     * @brief Check that a double value is within range, and format it as a string with decimalPlaces
     * @param val The value to check & format
     * @param min The lower bound of the valid value range
     * @param max The upper bound of the valid value range 0 must be larger than min
     * @param absTolerance The tolerance for rounding errors when checking the valid range - must be positive
     * @param decimalPlaces The amount of digits following the decimal separator that shall be included in the formatted string
     * @return The value formatted as a string with the desired amount of decimal places
     * @return An empty string ""s if the value falls outside the valid range
     * @note values that fall outside the range but within the tolerance will be rounded to the nearest range bound
     */
    std::string checkAndFormatDoubleAsString(double val, double min, double max, double absTolerance, int decimalPlaces = 2)
    {
        if (max <= min || absTolerance < 0.0) throw XLException("checkAndFormatDoubleAsString: max must be greater than min and absTolerance must be >= 0.0");
        if (val < min - absTolerance || val > max + absTolerance) return ""; // range check
        if (val < min) val = min; else if (val > max) val = max;             // fix rounding errors within tolerance
        return formatDoubleAsString(val, decimalPlaces);
    }
}    // anonymous namespace



/**
 * @details Constructor. Initializes an empty XLNumberFormat object
 */
XLNumberFormat::XLNumberFormat() : m_numberFormatNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLNumberFormat object.
 */
XLNumberFormat::XLNumberFormat(const XMLNode& node) : m_numberFormatNode(std::make_unique<XMLNode>(node)) {}

XLNumberFormat::~XLNumberFormat() = default;

XLNumberFormat::XLNumberFormat(const XLNumberFormat& other)
    : m_numberFormatNode(std::make_unique<XMLNode>(*other.m_numberFormatNode))
{}

XLNumberFormat& XLNumberFormat::operator=(const XLNumberFormat& other)
{
    if (&other != this) *m_numberFormatNode = *other.m_numberFormatNode;
    return *this;
}

/**
 * @details Returns the numFmtId value
 */
uint32_t XLNumberFormat::numberFormatId() const { return m_numberFormatNode->attribute("numFmtId").as_uint(XLInvalidUInt32); }

/**
 * @details Returns the formatCode value
 */
std::string XLNumberFormat::formatCode() const { return m_numberFormatNode->attribute("formatCode").value(); }

/**
 * @details Setter functions
 */
bool XLNumberFormat::setNumberFormatId(uint32_t newNumberFormatId)
    { return appendAndSetAttribute(*m_numberFormatNode, "numFmtId",   std::to_string(newNumberFormatId)).empty() == false; }
bool XLNumberFormat::setFormatCode    (std::string newFormatCode)
    { return appendAndSetAttribute(*m_numberFormatNode, "formatCode", newFormatCode.c_str()            ).empty() == false; }


/**
 * @details assemble a string summary about the number format
 */
std::string XLNumberFormat::summary() const
{
    using namespace std::literals::string_literals;
    return "numFmtId="s + std::to_string(numberFormatId())
         + ", formatCode="s + formatCode();
}


// ===== XLNumberFormats, parent of XLNumberFormat

/**
 * @details Constructor. Initializes an empty XLNumberFormats object
 */
XLNumberFormats::XLNumberFormats() : m_numberFormatsNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLNumberFormats object.
 */
XLNumberFormats::XLNumberFormats(const XMLNode& numberFormats)
    : m_numberFormatsNode(std::make_unique<XMLNode>(numberFormats))
{
    // initialize XLNumberFormat entries and m_numberFormats here
    XMLNode node = numberFormats.first_child_of_type(pugi::node_element);
    while (not node.empty()) {
        std::string nodeName = node.name();
        // std::cout << "XMLNumberFormats constructor, node name is " << nodeName << std::endl;
        if (nodeName == "numFmt")
            m_numberFormats.push_back(XLNumberFormat(node));
        else
            std::cerr << "WARNING: XLNumberFormats constructor: unknown subnode " << nodeName << std::endl;
        node = node.next_sibling_of_type(pugi::node_element);
    }
}

XLNumberFormats::~XLNumberFormats()
{
    m_numberFormats.clear(); // delete vector with all children
}

XLNumberFormats::XLNumberFormats(const XLNumberFormats& other)
    : m_numberFormatsNode(std::make_unique<XMLNode>(*other.m_numberFormatsNode)),
      m_numberFormats(other.m_numberFormats)
{
    // std::cout << __func__ << " copy constructor" << std::endl;
}

XLNumberFormats::XLNumberFormats(XLNumberFormats&& other)
    : m_numberFormatsNode(std::move(other.m_numberFormatsNode)),
      m_numberFormats(std::move(other.m_numberFormats))
{
    // std::cout << __func__ << " move constructor" << std::endl;
}


/**
 * @details Copy assignment operator
 */
XLNumberFormats& XLNumberFormats::operator=(const XLNumberFormats& other)
{
    // std::cout << "XLNumberFormats::" << __func__ << " copy assignment" << std::endl;
    if (&other != this) {
        *m_numberFormatsNode = *other.m_numberFormatsNode;
        m_numberFormats.clear();
        m_numberFormats = other.m_numberFormats;
    }
    return *this;
}

/**
 * @details Returns the amount of numberFormats held by the class
 */
size_t XLNumberFormats::count() const { return m_numberFormats.size(); }

/**
 * @details fetch XLNumberFormat from m_numberFormats by index
 */
XLNumberFormat XLNumberFormats::numberFormatByIndex(XLStyleIndex index) const
{
    if (index >= m_numberFormats.size()) {
        using namespace std::literals::string_literals;
        throw XLException("XLNumberFormats::"s + __func__ + ": index "s + std::to_string(index) + " is out of range"s);
    }
    return m_numberFormats.at(index);
}

/**
 * @details fetch XLNumberFormat from m_numberFormats by its numberFormatId
 */
XLNumberFormat XLNumberFormats::numberFormatById(uint32_t numberFormatId) const
{
    for (XLNumberFormat fmt : m_numberFormats)
        if (fmt.numberFormatId() == numberFormatId)
            return fmt;
    using namespace std::literals::string_literals;
    throw XLException("XLNumberFormats::"s + __func__ + ": numberFormatId "s + std::to_string(numberFormatId) + " not found"s);
}

/**
 * @details fetch a numFmtId from m_numberFormats by index
 */
uint32_t XLNumberFormats::numberFormatIdFromIndex(XLStyleIndex index) const
{
    if (index >= m_numberFormats.size()) {
        using namespace std::literals::string_literals;
        throw XLException("XLNumberFormats::"s + __func__ + ": index "s + std::to_string(index) + " is out of range"s);
    }
    return m_numberFormats[ index ].numberFormatId(); 
}

/**
 * @details append a new XLNumberFormat to m_numberFormats and m_numberFormatsNode, based on copyFrom
 */
XLStyleIndex XLNumberFormats::create(XLNumberFormat copyFrom, std::string styleEntriesPrefix)
{
    XLStyleIndex index = count();    // index for the number format to be created
    XMLNode newNode{};               // scope declaration

    // ===== Append new node prior to final whitespaces, if any
    XMLNode lastStyle = m_numberFormatsNode->last_child_of_type(pugi::node_element);
    if (lastStyle.empty()) newNode = m_numberFormatsNode->prepend_child("numFmt");
    else                   newNode = m_numberFormatsNode->insert_child_after("numFmt", lastStyle);
    if (newNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLNumberFormats::"s + __func__ + ": failed to append a new numFmt node"s);
    }
    if (styleEntriesPrefix.length() > 0)    // if a whitespace prefix is configured
        m_numberFormatsNode->insert_child_before(pugi::node_pcdata, newNode).set_value(styleEntriesPrefix.c_str());    // prefix the new node with styleEntriesPrefix

    XLNumberFormat newNumberFormat(newNode);
    if (copyFrom.m_numberFormatNode->empty()) {    // if no template is given
        // ===== Create a number format with default values
        newNumberFormat.setNumberFormatId(0);
        newNumberFormat.setFormatCode("General");
    }
    else
        copyXMLNode(newNode, *copyFrom.m_numberFormatNode); // will use copyFrom as template, does nothing if copyFrom is empty

    m_numberFormats.push_back(newNumberFormat);
    appendAndSetAttribute(*m_numberFormatsNode, "count", std::to_string(m_numberFormats.size())); // update array count in XML
    return index;
}


/**
 * @details Constructor. Initializes an empty XLFont object
 */
XLFont::XLFont() : m_fontNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLFont object.
 */
XLFont::XLFont(const XMLNode& node) : m_fontNode(std::make_unique<XMLNode>(node)) {}

XLFont::~XLFont() = default;

XLFont::XLFont(const XLFont& other)
    : m_fontNode(std::make_unique<XMLNode>(*other.m_fontNode))
{}

XLFont& XLFont::operator=(const XLFont& other)
{
    if (&other != this) *m_fontNode = *other.m_fontNode;
    return *this;
}


/**
 * @details Returns the font name property
 */
std::string XLFont::fontName() const
{
    XMLAttribute attr = appendAndGetNodeAttribute(*m_fontNode, "name", "val", OpenXLSX::XLDefaultFontName);
    return attr.value();
}

/**
 * @details Returns the font charset property
 */
size_t XLFont::fontCharset() const
{
    XMLAttribute attr = appendAndGetNodeAttribute(*m_fontNode, "charset", "val", std::to_string(OpenXLSX::XLDefaultFontCharset));
    return attr.as_uint();
}

/**
 * @details Returns the font family property
 */
size_t XLFont::fontFamily() const
{
    XMLAttribute attr = appendAndGetNodeAttribute(*m_fontNode, "family", "val", std::to_string(OpenXLSX::XLDefaultFontFamily));
    return attr.as_uint();
}

/**
 * @details Returns the font size property
 */
size_t XLFont::fontSize() const
{
    XMLAttribute attr = appendAndGetNodeAttribute(*m_fontNode, "sz", "val", std::to_string(OpenXLSX::XLDefaultFontSize));
    return attr.as_uint();
}

/**
 * @details Returns the font color property
 */
XLColor XLFont::fontColor() const
{
    using namespace std::literals::string_literals;
    // XMLAttribute attr = appendAndGetNodeAttribute(*m_fontNode, "color", "theme", OpenXLSX::XLDefaultFontColorTheme);
    // TBD what "theme" is and whether it should be supported at all
    XMLAttribute attr = appendAndGetNodeAttribute(*m_fontNode, "color", "rgb", XLDefaultFontColor);
    return XLColor(attr.value());
}

/**
 * @details getter functions: return the font's bold, italic, underline, strikethrough status
 */
bool                    XLFont::bold()          const { return                                   appendAndGetNodeAttribute(*m_fontNode, "b",         "val", "false").as_bool() ; }
bool                    XLFont::italic()        const { return                                   appendAndGetNodeAttribute(*m_fontNode, "i",         "val", "false").as_bool() ; }
bool                    XLFont::strikethrough() const { return                                   appendAndGetNodeAttribute(*m_fontNode, "strike",    "val", "false").as_bool() ; }
XLUnderlineStyle        XLFont::underline()     const { return XLUnderlineStyleFromString       (appendAndGetNodeAttribute(*m_fontNode, "u",         "val", ""     ).value()  ); }
XLFontSchemeStyle       XLFont::scheme()        const { return XLFontSchemeStyleFromString      (appendAndGetNodeAttribute(*m_fontNode, "scheme",    "val", ""     ).value()  ); }
XLVerticalAlignRunStyle XLFont::vertAlign()     const { return XLVerticalAlignRunStyleFromString(appendAndGetNodeAttribute(*m_fontNode, "vertAlign", "val", ""     ).value()  ); }
bool                    XLFont::outline()       const { return                                   appendAndGetNodeAttribute(*m_fontNode, "outline",   "val", "false").as_bool() ; }
bool                    XLFont::shadow()        const { return                                   appendAndGetNodeAttribute(*m_fontNode, "shadow",    "val", "false").as_bool() ; }
bool                    XLFont::condense()      const { return                                   appendAndGetNodeAttribute(*m_fontNode, "condense",  "val", "false").as_bool() ; }
bool                    XLFont::extend()        const { return                                   appendAndGetNodeAttribute(*m_fontNode, "extend",    "val", "false").as_bool() ; }

/**
 * @details Setter functions
 */
bool XLFont::setFontName     (std::string newName)    { return appendAndSetNodeAttribute(*m_fontNode, "name",     "val", newName.c_str()                        ).empty() == false; }
bool XLFont::setFontCharset  (size_t newCharset)      { return appendAndSetNodeAttribute(*m_fontNode, "charset",  "val", std::to_string(newCharset)             ).empty() == false; }
bool XLFont::setFontFamily   (size_t newFamily)       { return appendAndSetNodeAttribute(*m_fontNode, "family",   "val", std::to_string(newFamily)              ).empty() == false; }
bool XLFont::setFontSize     (size_t newSize)         { return appendAndSetNodeAttribute(*m_fontNode, "sz",       "val", std::to_string(newSize)                ).empty() == false; }
bool XLFont::setFontColor    (XLColor newColor)       { return appendAndSetNodeAttribute(*m_fontNode, "color",    "rgb", newColor.hex(), XLRemoveAttributes     ).empty() == false; }
bool XLFont::setBold         (bool set)               { return appendAndSetNodeAttribute(*m_fontNode, "b",        "val", (set ? "true" : "false")               ).empty() == false; }
bool XLFont::setItalic       (bool set)               { return appendAndSetNodeAttribute(*m_fontNode, "i",        "val", (set ? "true" : "false")               ).empty() == false; }
bool XLFont::setStrikethrough(bool set)               { return appendAndSetNodeAttribute(*m_fontNode, "strike",   "val", (set ? "true" : "false")               ).empty() == false; }
bool XLFont::setUnderline    (XLUnderlineStyle style)
    { return appendAndSetNodeAttribute(*m_fontNode, "u",         "val", XLUnderlineStyleToString       (style       ).c_str()).empty() == false; }
bool XLFont::setScheme       (XLFontSchemeStyle newScheme)
    { return appendAndSetNodeAttribute(*m_fontNode, "scheme",    "val", XLFontSchemeStyleToString      (newScheme   ).c_str()).empty() == false; }
bool XLFont::setVertAlign    (XLVerticalAlignRunStyle newVertAlign)
    { return appendAndSetNodeAttribute(*m_fontNode, "vertAlign", "val", XLVerticalAlignRunStyleToString(newVertAlign).c_str()).empty() == false; }
bool XLFont::setOutline (bool set)                    { return appendAndSetNodeAttribute(*m_fontNode, "outline",  "val", (set ? "true" : "false")               ).empty() == false; }
bool XLFont::setShadow  (bool set)                    { return appendAndSetNodeAttribute(*m_fontNode, "shadow",   "val", (set ? "true" : "false")               ).empty() == false; }
bool XLFont::setCondense(bool set)                    { return appendAndSetNodeAttribute(*m_fontNode, "condense", "val", (set ? "true" : "false")               ).empty() == false; }
bool XLFont::setExtend  (bool set)                    { return appendAndSetNodeAttribute(*m_fontNode, "extend",   "val", (set ? "true" : "false")               ).empty() == false; }


/**
 * @details assemble a string summary about the font
 */
std::string XLFont::summary() const
{
    using namespace std::literals::string_literals;
    return "font name is "s + fontName()
         + ", charset: "s + std::to_string(fontCharset())
         + ", font family: "s + std::to_string(fontFamily())
         + ", size: "s + std::to_string(fontSize())
         + ", color: "s + fontColor().hex()
         + (bold() ? ", +bold"s : ""s)
         + (italic() ? ", +italic"s : ""s)
         + (strikethrough() ? ", +strikethrough"s : ""s)
         + (underline() != XLUnderlineNone ? ", underline: "s + XLUnderlineStyleToString(underline()) : ""s)
         + (scheme() != XLFontSchemeNone ? ", scheme: "s + XLFontSchemeStyleToString(scheme()) : ""s)
         + (vertAlign() != XLBaseline ? ", vertAlign: "s + XLVerticalAlignRunStyleToString(vertAlign()) : ""s)
         + (outline() ? ", +outline"s : ""s)
         + (shadow() ? ", +shadow"s : ""s)
         + (condense() ? ", +condense"s : ""s)
         + (extend() ? ", +extend"s : ""s);
}


// ===== XLFonts, parent of XLFont

/**
 * @details Constructor. Initializes an empty XLFonts object
 */
XLFonts::XLFonts() : m_fontsNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLFonts object.
 */
XLFonts::XLFonts(const XMLNode& fonts)
    : m_fontsNode(std::make_unique<XMLNode>(fonts))
{
    // initialize XLFonts entries and m_fonts here
    XMLNode node = m_fontsNode->first_child_of_type(pugi::node_element);
    while (not node.empty()) {
        std::string nodeName = node.name();
        if (nodeName == "font")
            m_fonts.push_back(XLFont(node));
        else
            std::cerr << "WARNING: XLFonts constructor: unknown subnode " << nodeName << std::endl;
        node = node.next_sibling_of_type(pugi::node_element);
    }
}

XLFonts::~XLFonts()
{
    m_fonts.clear(); // delete vector with all children
}

XLFonts::XLFonts(const XLFonts& other)
    : m_fontsNode(std::make_unique<XMLNode>(*other.m_fontsNode)),
      m_fonts(other.m_fonts)
{
    // std::cout << __func__ << " copy constructor" << std::endl;
}

XLFonts::XLFonts(XLFonts&& other)
    : m_fontsNode(std::move(other.m_fontsNode)),
      m_fonts(std::move(other.m_fonts))
{
    // std::cout << __func__ << " move constructor" << std::endl;
}


/**
 * @details Copy assignment operator
 */
XLFonts& XLFonts::operator=(const XLFonts& other)
{
    // std::cout << "XLFonts::" << __func__ << " copy assignment" << std::endl;
    if (&other != this) {
        *m_fontsNode = *other.m_fontsNode;
        m_fonts.clear();
        m_fonts = other.m_fonts;
    }
    return *this;
}

/**
 * @details Returns the amount of fonts held by the class
 */
size_t XLFonts::count() const { return m_fonts.size(); }

/**
 * @details fetch XLFont from m_Fonts by index
 */
XLFont XLFonts::fontByIndex(XLStyleIndex index) const
{
    if (index >= m_fonts.size()) {
        using namespace std::literals::string_literals;
        throw XLException("XLFonts::"s + __func__ + ": attempted to access index "s + std::to_string(index) + " with count "s + std::to_string(m_fonts.size()));
    }
    return m_fonts.at(index);
}

/**
 * @details append a new XLFont to m_fonts and m_fontsNode, based on copyFrom
 */
XLStyleIndex XLFonts::create(XLFont copyFrom, std::string styleEntriesPrefix)
{
    XLStyleIndex index = count();    // index for the font to be created
    XMLNode newNode{};               // scope declaration

    // ===== Append new node prior to final whitespaces, if any
    XMLNode lastStyle = m_fontsNode->last_child_of_type(pugi::node_element);
    if (lastStyle.empty()) newNode = m_fontsNode->prepend_child("font");
    else                   newNode = m_fontsNode->insert_child_after("font", lastStyle);
    if (newNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLFonts::"s + __func__ + ": failed to append a new fonts node"s);
    }
    if (styleEntriesPrefix.length() > 0)    // if a whitespace prefix is configured
        m_fontsNode->insert_child_before(pugi::node_pcdata, newNode).set_value(styleEntriesPrefix.c_str());    // prefix the new node with styleEntriesPrefix

    XLFont newFont(newNode);
    if (copyFrom.m_fontNode->empty()) {    // if no template is given
        // ===== Create a font with default values
        // TODO: implement font defaults
        // newFont.setProperty(defaultValue);
        // ...
    }
    else
        copyXMLNode(newNode, *copyFrom.m_fontNode);    // will use copyFrom as template, does nothing if copyFrom is empty

    m_fonts.push_back(newFont);
    appendAndSetAttribute(*m_fontsNode, "count", std::to_string(m_fonts.size())); // update array count in XML
    return index;
}


// ===== XLDataBarColor, used by XLFills gradientFill and by XLLine (to be implemented)

/**
 * @details Constructor. Initializes an empty XLDataBarColor object
 */
XLDataBarColor::XLDataBarColor() : m_colorNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLDataBarColor object.
 */
XLDataBarColor::XLDataBarColor(const XMLNode& node) : m_colorNode(std::make_unique<XMLNode>(node)) {}

/**
 * @details Copy constructor - initializes the member variables from other
 */
XLDataBarColor::XLDataBarColor(const XLDataBarColor& other)
    : m_colorNode(std::make_unique<XMLNode>(*other.m_colorNode))
{}

/**
 * @details Assign values of other to this
 */
XLDataBarColor& XLDataBarColor::operator=(const XLDataBarColor& other)
{
    if (&other != this) *m_colorNode = *other.m_colorNode;
    return *this;
}

/**
 * @details Getter functions
 */
XLColor  XLDataBarColor::rgb()       const { return XLColor(m_colorNode->attribute("rgb"    ).as_string("ffffffff")); }
double   XLDataBarColor::tint()      const { return         m_colorNode->attribute("tint"   ).as_double(0.0)        ; }
bool     XLDataBarColor::automatic() const { return         m_colorNode->attribute("auto"   ).as_bool()             ; }
uint32_t XLDataBarColor::indexed()   const { return         m_colorNode->attribute("indexed").as_uint()             ; }
uint32_t XLDataBarColor::theme()     const { return         m_colorNode->attribute("theme"  ).as_uint()             ; }
/**
 * @details Setter functions
 */
bool XLDataBarColor::setRgb (XLColor newColor)     { return appendAndSetAttribute(*m_colorNode, "rgb",       newColor.hex()               ).empty() == false; }
bool XLDataBarColor::setTint(double newTint)
{
    std::string tintString = "";
    if (newTint != 0.0) {
        tintString = checkAndFormatDoubleAsString(newTint, -1.0, +1.0, 0.01);
        if (tintString.length() == 0) {
            using namespace std::literals::string_literals;
            throw XLException("XLDataBarColor::setTint: color tint "s + std::to_string(newTint) + " is not in range [-1.0;+1.0]"s);
        }
    }
    if (tintString.length() == 0) return m_colorNode->remove_attribute("tint");           // remove tint attribute for a value 0

    return (appendAndSetAttribute(*m_colorNode, "tint", tintString).empty() == false);    // else: set tint attribute
}
bool XLDataBarColor::setAutomatic(bool set)        { return appendAndSetAttribute(*m_colorNode, "auto",      (set ? "true" : "false")     ).empty() == false; }
bool XLDataBarColor::setIndexed(uint32_t newIndex) { return appendAndSetAttribute(*m_colorNode, "indexed",   std::to_string(newIndex)     ).empty() == false; }
bool XLDataBarColor::setTheme(uint32_t newTheme)   { return appendAndSetAttribute(*m_colorNode, "theme",     std::to_string(newTheme)     ).empty() == false; }

/**
 * @details assemble a string summary about the color
 */
std::string XLDataBarColor::summary() const
{
    using namespace std::literals::string_literals;
    return "rgb is "s + rgb().hex()
         + ", tint is "s + formatDoubleAsString(tint())
         + (automatic() ? ", +automatic"s : ""s)
         + (indexed() ? (", index is "s + std::to_string(indexed())) : "")
         + (theme()   ? (", theme is "s + std::to_string(theme()  )) : "")
         ;
}


/**
 * @details Constructor. Initializes an empty XLGradientStop object
 */
XLGradientStop::XLGradientStop() : m_stopNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLGradientStop object.
 */
XLGradientStop::XLGradientStop(const XMLNode& node) : m_stopNode(std::make_unique<XMLNode>(node)) {}

/**
 * @details Copy constructor - initializes the member variables from other
 */
XLGradientStop::XLGradientStop(const XLGradientStop& other)
    : m_stopNode(std::make_unique<XMLNode>(*other.m_stopNode))
{}

/**
 * @details Assign values of other to this
 */
XLGradientStop& XLGradientStop::operator=(const XLGradientStop& other)
{
    if (&other != this) *m_stopNode = *other.m_stopNode;
    return *this;
}

/**
 * @details Getter functions
 */
XLDataBarColor XLGradientStop::color() const
{
    XMLNode color = appendAndGetNode(*m_stopNode, "color");
    if (color.empty()) return XLDataBarColor{};
    return XLDataBarColor(color);
}
double XLGradientStop::position() const
{
    XMLAttribute attr = m_stopNode->attribute("position");
    return attr.as_double(0.0);
}

/**
 * @details Setter functions
 */
bool XLGradientStop::setPosition(double newPosition) { return appendAndSetAttribute(*m_stopNode, "position", formatDoubleAsString(newPosition)).empty() == false; }


/**
 * @details assemble a string summary about the stop
 */
std::string XLGradientStop::summary() const
{
    using namespace std::literals::string_literals;
    return "stop position is "s + formatDoubleAsString(position())
         + ", " + color().summary();
}


// ===== XLGradientStops, parent of XLGradientStop

/**
 * @details Constructor. Initializes an empty XLGradientStops object
 */
XLGradientStops::XLGradientStops() : m_gradientNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLGradientStops object.
 */
XLGradientStops::XLGradientStops(const XMLNode& gradient)
    : m_gradientNode(std::make_unique<XMLNode>(gradient))
{
    // initialize XLGradientStops entries and m_gradientStops here
    XMLNode node = m_gradientNode->first_child_of_type(pugi::node_element);
    while (not node.empty()) {
        std::string nodeName = node.name();
        if (nodeName == "stop")
            m_gradientStops.push_back(XLGradientStop(node));
        else
            std::cerr << "WARNING: XLGradientStops constructor: unknown subnode " << nodeName << std::endl;
        node = node.next_sibling_of_type(pugi::node_element);
    }
}

XLGradientStops::~XLGradientStops()
{
    m_gradientStops.clear(); // delete vector with all children
}

XLGradientStops::XLGradientStops(const XLGradientStops& other)
    : m_gradientNode(std::make_unique<XMLNode>(*other.m_gradientNode)),
      m_gradientStops(other.m_gradientStops)
{
    // std::cout << __func__ << " copy constructor" << std::endl;
}

XLGradientStops::XLGradientStops(XLGradientStops&& other)
    : m_gradientNode(std::move(other.m_gradientNode)),
      m_gradientStops(std::move(other.m_gradientStops))
{
    // std::cout << __func__ << " move constructor" << std::endl;
}


/**
 * @details Copy assignment operator
 */
XLGradientStops& XLGradientStops::operator=(const XLGradientStops& other)
{
    // std::cout << "XLGradientStops::" << __func__ << " copy assignment" << std::endl;
    if (&other != this) {
        *m_gradientNode = *other.m_gradientNode;
        m_gradientStops.clear();
        m_gradientStops = other.m_gradientStops;
    }
    return *this;
}

/**
 * @details Returns the amount of gradient stops held by the class
 */
size_t XLGradientStops::count() const { return m_gradientStops.size(); }

/**
 * @details fetch XLGradientStop from m_gradientStops by index
 */
XLGradientStop XLGradientStops::stopByIndex(XLStyleIndex index) const
{
    if (index >= m_gradientStops.size()) {
        using namespace std::literals::string_literals;
        throw XLException("XLGradientStops::"s + __func__ + ": attempted to access index "s + std::to_string(index) + " with count "s + std::to_string(m_gradientStops.size()));
    }
    return m_gradientStops.at(index);
}

/**
 * @details append a new XLGradientStop to m_gradientStops m_gradientNode, based on copyFrom
 */
XLStyleIndex XLGradientStops::create(XLGradientStop copyFrom, std::string styleEntriesPrefix)
{
    XLStyleIndex index = count();    // index for the gradient stop to be created
    XMLNode newNode{};               // scope declaration

    // ===== Append new node prior to final whitespaces, if any
    XMLNode lastStyle = m_gradientNode->last_child_of_type(pugi::node_element);
    if (lastStyle.empty()) newNode = m_gradientNode->prepend_child("stop");
    else                   newNode = m_gradientNode->insert_child_after("stop", lastStyle);
    if (newNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLGradientStops::"s + __func__ + ": failed to append a new stop node"s);
    }
    if (styleEntriesPrefix.length() > 0)    // if a whitespace prefix is configured
        m_gradientNode->insert_child_before(pugi::node_pcdata, newNode).set_value(styleEntriesPrefix.c_str());    // prefix the new node with styleEntriesPrefix

    XLGradientStop newStop(newNode);
    if (copyFrom.m_stopNode->empty()) {    // if no template is given
        // ===== Create a stop node with default values
        // TODO: implement stop defaults
        // newStop.setProperty(defaultValue);
        // ...
    }
    else
        copyXMLNode(newNode, *copyFrom.m_stopNode);    // will use copyFrom as template, does nothing if copyFrom is empty

    m_gradientStops.push_back(newStop);
    appendAndSetAttribute(*m_gradientNode, "count", std::to_string(m_gradientStops.size())); // update array count in XML
    return index;
}

std::string XLGradientStops::summary() const
{
    std::string result{};
    for(XLGradientStop stop : m_gradientStops)
        result += stop.summary() + ", ";
    if (result.length() < 2) return ""; // if no stop summary was created - return empty string
    return result.substr(0, result.length() - 2);
}

/**
 * @details Constructor. Initializes an empty XLFill object
 */
XLFill::XLFill() : m_fillNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLFill object.
 */
XLFill::XLFill(const XMLNode& node) : m_fillNode(std::make_unique<XMLNode>(node)) {}

XLFill::~XLFill() = default;

XLFill::XLFill(const XLFill& other)
    : m_fillNode(std::make_unique<XMLNode>(*other.m_fillNode))
{}

XLFill& XLFill::operator=(const XLFill& other)
{
    if (&other != this) *m_fillNode = *other.m_fillNode;
    return *this;
}

/**
 * @details Returns the name of the first element child of fill
 * @note an empty node ::name() returns an empty string "", leading to XLFillTypeInvalid
 */
XLFillType XLFill::fillType() const { return XLFillTypeFromString(m_fillNode->first_child_of_type(pugi::node_element).name()); }

/**
 * @details set the fill type for a fill node - if force is true, delete any existing fill properties
 */
bool XLFill::setFillType(XLFillType newFillType, bool force)
{
    XLFillType ft = fillType(); // determine once, use twice

    // ===== If desired filltype is already set
    if (ft == newFillType) return true;    // nothing to do

    // ===== If force == true or fillType is just not set at all, delete existing child nodes, otherwise throw
    if (!force && ft != XLFillTypeInvalid) {
        using namespace std::literals::string_literals;
        throw XLException("XLFill::setFillType: can not change the fillType from "s + XLFillTypeToString(fillType())
                        + " to "s + XLFillTypeToString(newFillType) + " - invoke with force == true if override is desired"s);
    }
    // ===== At this point, m_fillNode needs to be emptied for a setting / force-change of the fill type

    // ===== Delete all fill node children and insert a new node for the newFillType
    m_fillNode->remove_children();
    return (m_fillNode->append_child(XLFillTypeToString(newFillType).c_str()).empty() == false);
}

/**
 * @details Throw an XLException on a fill of typeToThrowOn
 */
void XLFill::throwOnFillType(XLFillType typeToThrowOn, const char *functionName) const
{
    using namespace std::literals::string_literals;
    if (fillType() == typeToThrowOn)
        throw XLException("XLFill::"s + functionName + " must not be invoked for a "s + XLFillTypeToString(typeToThrowOn));
}

/**
 * @details get the fill element XML, create element with default XLFillType if none exists
 */
XMLNode XLFill::getValidFillDescription(XLFillType fillTypeIfEmpty, const char *functionName)
{
    XLFillType throwOnThis = XLFillTypeInvalid;
    switch (fillTypeIfEmpty) {
        case XLGradientFill: throwOnThis = XLPatternFill; break;     // throw on non-matching fill type
        case XLPatternFill : throwOnThis = XLGradientFill; break;    //   "
        default: throw XLException("XLFill::getValidFillDescription was not invoked with XLPatternFill or XLGradientFill");
    }
    throwOnFillType(throwOnThis, functionName);
    XMLNode fillDescription = m_fillNode->first_child_of_type(pugi::node_element); // fetch an existing fill description
    if (fillDescription.empty() && setFillType(fillTypeIfEmpty))                   // if none exists, attempt to create a description
        fillDescription = m_fillNode->first_child_of_type(pugi::node_element);         // fetch newly inserted description
    return fillDescription;
}


/**
 * @details Getter functions for gradientFill
 */
XLGradientType  XLFill::gradientType() { return XLGradientTypeFromString(getValidFillDescription(XLGradientFill, __func__).attribute("type"  ).value()     ); }
double          XLFill::degree()       { return                          getValidFillDescription(XLGradientFill, __func__).attribute("degree").as_double(0) ; }
double          XLFill::left()         { return                          getValidFillDescription(XLGradientFill, __func__).attribute("left"  ).as_double(0) ; }
double          XLFill::right()        { return                          getValidFillDescription(XLGradientFill, __func__).attribute("right" ).as_double(0) ; }
double          XLFill::top()          { return                          getValidFillDescription(XLGradientFill, __func__).attribute("top"   ).as_double(0) ; }
double          XLFill::bottom()       { return                          getValidFillDescription(XLGradientFill, __func__).attribute("bottom").as_double(0) ; }
XLGradientStops XLFill::stops()        { return XLGradientStops(         getValidFillDescription(XLGradientFill, __func__)                                 ); }

/**
 * @details Getter functions for patternFill
 */
XLPatternType XLFill::patternType()
{
    XMLNode fillDescription = getValidFillDescription(XLPatternFill, __func__);
    if (fillDescription.empty()) return XLDefaultPatternType;    // if no description could be fetched: fail
    XMLAttribute attr = appendAndGetNodeAttribute(*m_fillNode, XLFillTypeToString(XLPatternFill), "patternType", XLPatternTypeToString(XLDefaultPatternType));
    return XLPatternTypeFromString(attr.value());
}

XLColor XLFill::color()
{
    XMLNode fillDescription = getValidFillDescription(XLPatternFill, __func__);
    if (fillDescription.empty()) return XLColor{}; // if no description could be fetched: fail
    XMLAttribute fgColorRGB = appendAndGetNodeAttribute(fillDescription, "fgColor", "rgb", XLDefaultPatternFgColor);
    return XLColor(fgColorRGB.value());
}
XLColor XLFill::backgroundColor()
{
    XMLNode fillDescription = getValidFillDescription(XLPatternFill, __func__);
    if (fillDescription.empty()) return XLColor{}; // if no description could be fetched: fail
    XMLAttribute bgColorRGB = appendAndGetNodeAttribute(fillDescription, "bgColor", "rgb", XLDefaultPatternBgColor);
    return XLColor(bgColorRGB.value());
}

/**
 * @details Setter functions for gradientFill
 */
bool XLFill::setGradientType(XLGradientType newType)
{
    XMLNode fillDescription = getValidFillDescription(XLGradientFill, __func__);
    return appendAndSetAttribute(fillDescription, "type", XLGradientTypeToString(newType)).empty() == false;
}
bool XLFill::setDegree(double newDegree)
{
    std::string degreeString = checkAndFormatDoubleAsString(newDegree, 0.0, 360.0, 0.01);
    if (degreeString.length() == 0) {
        using namespace std::literals::string_literals;
        throw XLException("XLFill::setDegree: gradientFill degree value "s + std::to_string(newDegree) + " is not in range [0.0;360.0]"s);
    }
    XMLNode fillDescription = getValidFillDescription(XLGradientFill, __func__);
    return appendAndSetAttribute(fillDescription, "degree", degreeString                         ).empty() == false;
}
bool XLFill::setLeft(double newLeft)
{
    XMLNode fillDescription = getValidFillDescription(XLGradientFill, __func__);
    return appendAndSetAttribute(fillDescription, "left",   formatDoubleAsString(newLeft  ).c_str()).empty() == false;
}
bool XLFill::setRight(double newRight)
{
    XMLNode fillDescription = getValidFillDescription(XLGradientFill, __func__);
    return appendAndSetAttribute(fillDescription, "right",  formatDoubleAsString(newRight ).c_str()).empty() == false;
}
bool XLFill::setTop(double newTop)
{
    XMLNode fillDescription = getValidFillDescription(XLGradientFill, __func__);
    return appendAndSetAttribute(fillDescription, "top",    formatDoubleAsString(newTop   ).c_str()).empty() == false;
}
bool XLFill::setBottom(double newBottom)
{
    XMLNode fillDescription = getValidFillDescription(XLGradientFill, __func__);
    return appendAndSetAttribute(fillDescription, "bottom", formatDoubleAsString(newBottom).c_str()).empty() == false;
}

/**
 * @details Setter functions for patternFill
 */
bool XLFill::setPatternType(XLPatternType newFillPattern)
{
    XMLNode fillDescription = getValidFillDescription(XLPatternFill, __func__);
    if (fillDescription.empty()) return false; // if no description could be fetched: fail
    return appendAndSetAttribute(fillDescription, "patternType", XLPatternTypeToString(newFillPattern)).empty() == false;
}
bool XLFill::setColor(XLColor newColor)
{
    XMLNode fillDescription = getValidFillDescription(XLPatternFill, __func__);
    if (fillDescription.empty()) return false; // if no description could be fetched: fail
    return appendAndSetNodeAttribute(fillDescription, "fgColor", "rgb", newColor.hex(), XLRemoveAttributes).empty() == false;
}
bool XLFill::setBackgroundColor(XLColor newBgColor)
{
    XMLNode fillDescription = getValidFillDescription(XLPatternFill, __func__);
    if (fillDescription.empty()) return false; // if no description could be fetched: fail
    return appendAndSetNodeAttribute(fillDescription, "bgColor", "rgb", newBgColor.hex(), XLRemoveAttributes).empty() == false;
}


/**
 * @details assemble a string summary about the fill
 */
std::string XLFill::summary()
{
    using namespace std::literals::string_literals;
    switch (fillType()) {
        case XLGradientFill:
            return "fill type is "s + XLFillTypeToString(fillType())
                 + ", gradient type is "s + XLGradientTypeToString(gradientType())
                 + ", degree is: "s + formatDoubleAsString(degree())
                 + ", left is: "s   + formatDoubleAsString(left())
                 + ", right is: "s  + formatDoubleAsString(right())
                 + ", top is: "s    + formatDoubleAsString(top())
                 + ", bottom is: "s + formatDoubleAsString(bottom())
                 + ", stops: "s + stops().summary();
        case XLPatternFill :
            return "fill type is "s + XLFillTypeToString(fillType())
                 + ", pattern type is "s + XLPatternTypeToString(patternType())
                 + ", fgcolor is: "s + color().hex() + ", bgcolor: "s + backgroundColor().hex();
        case XLFillTypeInvalid: [[fallthrough]];
        default               :
            return "fill type is invalid!"s;
    }
}

// ===== XLFills, parent of XLFill

/**
 * @details Constructor. Initializes an empty XLFills object
 */
XLFills::XLFills() : m_fillsNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLFills object.
 */
XLFills::XLFills(const XMLNode& fills)
    : m_fillsNode(std::make_unique<XMLNode>(fills))
{
    // initialize XLFills entries and m_fills here
    XMLNode node = fills.first_child_of_type(pugi::node_element);
    while (not node.empty()) {
        std::string nodeName = node.name();
        if (nodeName == "fill")
            m_fills.push_back(XLFill(node));
        else
            std::cerr << "WARNING: XLFills constructor: unknown subnode " << nodeName << std::endl;
        node = node.next_sibling_of_type(pugi::node_element);
    }
}

XLFills::~XLFills()
{
    m_fills.clear(); // delete vector with all children
}

XLFills::XLFills(const XLFills& other)
    : m_fillsNode(std::make_unique<XMLNode>(*other.m_fillsNode)),
      m_fills(other.m_fills)
{
    // std::cout << __func__ << " copy constructor" << std::endl;
}

XLFills::XLFills(XLFills&& other)
    : m_fillsNode(std::move(other.m_fillsNode)),
      m_fills(std::move(other.m_fills))
{
    // std::cout << __func__ << " move constructor" << std::endl;
}


/**
 * @details Copy assignment operator
 */
XLFills& XLFills::operator=(const XLFills& other)
{
    // std::cout << "XLFills::" << __func__ << " copy assignment" << std::endl;
    if (&other != this) {
        *m_fillsNode = *other.m_fillsNode;
        m_fills.clear();
        m_fills = other.m_fills;
    }
    return *this;
}

/**
 * @details Returns the amount of fills held by the class
 */
size_t XLFills::count() const { return m_fills.size(); }

/**
 * @details fetch XLFill from m_fills by index
 */
XLFill XLFills::fillByIndex(XLStyleIndex index) const
{
    if (index >= m_fills.size()) {
        using namespace std::literals::string_literals;
        throw XLException("XLFills::"s + __func__ + ": attempted to access index "s + std::to_string(index) + " with count "s + std::to_string(m_fills.size()));
    }
    return m_fills.at(index);
}

/**
 * @details append a new XLFill to m_fills and m_fillsNode, based on copyFrom
 */
XLStyleIndex XLFills::create(XLFill copyFrom, std::string styleEntriesPrefix)
{
    XLStyleIndex index = count();    // index for the fill to be created
    XMLNode newNode{};               // scope declaration

    // ===== Append new node prior to final whitespaces, if any
    XMLNode lastStyle = m_fillsNode->last_child_of_type(pugi::node_element);
    if (lastStyle.empty()) newNode = m_fillsNode->prepend_child("fill");
    else                   newNode = m_fillsNode->insert_child_after("fill", lastStyle);
    if (newNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLFills::"s + __func__ + ": failed to append a new fill node"s);
    }
    if (styleEntriesPrefix.length() > 0)    // if a whitespace prefix is configured
        m_fillsNode->insert_child_before(pugi::node_pcdata, newNode).set_value(styleEntriesPrefix.c_str());    // prefix the new node with styleEntriesPrefix

    XLFill newFill(newNode);
    if (copyFrom.m_fillNode->empty()) {    // if no template is given
        // ===== Create a fill with default values
        // TODO: implement fill defaults
        // newFill.setProperty(defaultValue);
        // ...
    }
    else
        copyXMLNode(newNode, *copyFrom.m_fillNode);    // will use copyFrom as template, does nothing if copyFrom is empty

    m_fills.push_back(newFill);
    appendAndSetAttribute(*m_fillsNode, "count", std::to_string(m_fills.size())); // update array count in XML
    return index;
}


/**
 * @details Constructor. Initializes an empty XLLine object
 */
XLLine::XLLine() : m_lineNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLLine object.
 */
XLLine::XLLine(const XMLNode& node) : m_lineNode(std::make_unique<XMLNode>(node)) {}

XLLine::~XLLine() = default;

XLLine::XLLine(const XLLine& other)
    : m_lineNode(std::make_unique<XMLNode>(*other.m_lineNode))
{}

XLLine& XLLine::operator=(const XLLine& other)
{
    if (&other != this) *m_lineNode = *other.m_lineNode;
    return *this;
}

/**
 * @details Returns the line style (XLLineStyleNone if line is not set)
 */
XLLineStyle XLLine::style() const
{
    if (m_lineNode->empty()) return XLLineStyleNone;
    XMLAttribute attr = appendAndGetAttribute(*m_lineNode, "style", OpenXLSX::XLDefaultLineStyle);
    return XLLineStyleFromString(attr.value());
}

/**
 * @details check if line is used (set) or not
 */
XLLine::operator bool() const { return (style() != XLLineStyleNone); }

/**
 * @details Returns the line data bar color object
 */
XLDataBarColor XLLine::color() const
{
    XMLNode color = appendAndGetNode(*m_lineNode, "color");
    if (color.empty()) return XLDataBarColor{};
    return XLDataBarColor(color);
}
// XLColor XLLine::color() const
// {
//     XMLNode colorDetails = m_lineNode->child("color");
//     if (colorDetails.empty()) return XLColor("ffffffff");
//     XMLAttribute colorRGB = colorDetails.attribute("rgb");
//     if (colorRGB.empty()) return XLColor("ffffffff");
//     return XLColor(colorRGB.value());
// }
// 
// /**
//  * @details Returns the line color tint
//  */
// double XLLine::colorTint() const
// {
//     XMLNode colorDetails = m_lineNode->child("color");
//     if (not colorDetails.empty()) {
//         XMLAttribute colorTint = colorDetails.attribute("tint");
//         if (not colorTint.empty())
//             return colorTint.as_double();
//     }
//     return 0.0;
// }

/**
 * @details assemble a string summary about the fill
 */
std::string XLLine::summary() const
{
    using namespace std::literals::string_literals;
    return "line style is "s + XLLineStyleToString(style()) + ", "s + color().summary();
    // double tint = colorTint();
    // std::string tintSummary = "colorTint is "s + (tint == 0.0 ? "(none)"s : std::to_string(tint));
    // size_t tintDecimalPos = tintSummary.find_last_of('.');
    // if (tintDecimalPos != std::string::npos) tintSummary = tintSummary.substr(0, tintDecimalPos + 3); // truncate colorTint double output 2 digits after the decimal separator
    // return "line style is "s + XLLineStyleToString(style()) + ", color is "s + color().hex() + ", "s + tintSummary;
}


/**
 * @details Constructor. Initializes an empty XLBorder object
 */
XLBorder::XLBorder() : m_borderNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLBorder object.
 */
XLBorder::XLBorder(const XMLNode& node) : m_borderNode(std::make_unique<XMLNode>(node)) {}

XLBorder::~XLBorder() = default;

XLBorder::XLBorder(const XLBorder& other)
    : m_borderNode(std::make_unique<XMLNode>(*other.m_borderNode))
{}

XLBorder& XLBorder::operator=(const XLBorder& other)
{
    if (&other != this) *m_borderNode = *other.m_borderNode;
    return *this;
}

/**
 * @details determines whether the diagonalUp property is set
 */
bool XLBorder::diagonalUp() const { return m_borderNode->attribute("diagonalUp").as_bool(); }

/**
 * @details determines whether the diagonalDown property is set
 */
bool XLBorder::diagonalDown() const { return m_borderNode->attribute("diagonalDown").as_bool(); }

/**
 * @details determines whether the outline property is set
 */
bool XLBorder::outline() const { return m_borderNode->attribute("outline").as_bool(); }

/**
 * @details fetch lines
 */
XLLine XLBorder::left()       const { return XLLine(m_borderNode->child("left")      ); }
XLLine XLBorder::right()      const { return XLLine(m_borderNode->child("right")     ); }
XLLine XLBorder::top()        const { return XLLine(m_borderNode->child("top")       ); }
XLLine XLBorder::bottom()     const { return XLLine(m_borderNode->child("bottom")    ); }
XLLine XLBorder::diagonal()   const { return XLLine(m_borderNode->child("diagonal")  ); }
XLLine XLBorder::vertical()   const { return XLLine(m_borderNode->child("vertical")  ); }
XLLine XLBorder::horizontal() const { return XLLine(m_borderNode->child("horizontal")); }

/**
 * @details Setter functions
 */
bool XLBorder::setDiagonalUp  (bool set) { return appendAndSetAttribute(*m_borderNode, "diagonalUp",   (set ? "true" : "false")).empty() == false; }
bool XLBorder::setDiagonalDown(bool set) { return appendAndSetAttribute(*m_borderNode, "diagonalDown", (set ? "true" : "false")).empty() == false; }
bool XLBorder::setOutline     (bool set) { return appendAndSetAttribute(*m_borderNode, "outline",      (set ? "true" : "false")).empty() == false; }
bool XLBorder::setLine(XLLineType lineType, XLLineStyle lineStyle, XLColor lineColor, double lineTint)
{
    // std::string tintString = "";
    // if (lineTint != 0.0) {
    //     tintString = checkAndFormatDoubleAsString(lineTint, -1.0, +1.0, 0.01);
    //     if (tintString.length() == 0) {
    //         using namespace std::literals::string_literals;
    //         throw XLException("XLBorder::setLine: line tint "s + std::to_string(lineTint) + " is not in range [-1.0;+1.0]"s);
    //     }
    // }

    XMLNode lineNode = appendAndGetNode(*m_borderNode, XLLineTypeToString(lineType));    // generate line node if not present
    bool success = (lineNode.empty() == false);
    if (success) success = (appendAndSetAttribute(lineNode, "style", XLLineStyleToString(lineStyle)).empty() == false); // set style attribute
    XMLNode colorNode{}; // empty node
    if (success) colorNode = appendAndGetNode(lineNode, "color");                        // generate color node if not present
    XLDataBarColor colorObject{colorNode};
    success = (colorNode.empty() == false);
    if (success) success = colorObject.setRgb(lineColor);
    if (success) success = colorObject.setTint(lineTint);
    // if (success) success = (appendAndSetAttribute(colorNode, "rgb", lineColor.hex()).empty() == false);                 // set rgb attribute
    // if (success) {
    //     if (tintString.length() == 0) success = colorNode.remove_attribute("tint");                                     // unset: remove tint attribute
    //     else success = (appendAndSetAttribute(colorNode, "tint", tintString).empty() == false);                         // set tint attribute
    // }
    return success;
}
bool XLBorder::setLeft      (XLLineStyle lineStyle, XLColor lineColor, double lineTint) { return setLine(XLLineLeft,       lineStyle, lineColor, lineTint); }
bool XLBorder::setRight     (XLLineStyle lineStyle, XLColor lineColor, double lineTint) { return setLine(XLLineRight,      lineStyle, lineColor, lineTint); }
bool XLBorder::setTop       (XLLineStyle lineStyle, XLColor lineColor, double lineTint) { return setLine(XLLineTop,        lineStyle, lineColor, lineTint); }
bool XLBorder::setBottom    (XLLineStyle lineStyle, XLColor lineColor, double lineTint) { return setLine(XLLineBottom,     lineStyle, lineColor, lineTint); }
bool XLBorder::setDiagonal  (XLLineStyle lineStyle, XLColor lineColor, double lineTint) { return setLine(XLLineDiagonal,   lineStyle, lineColor, lineTint); }
bool XLBorder::setVertical  (XLLineStyle lineStyle, XLColor lineColor, double lineTint) { return setLine(XLLineVertical,   lineStyle, lineColor, lineTint); }
bool XLBorder::setHorizontal(XLLineStyle lineStyle, XLColor lineColor, double lineTint) { return setLine(XLLineHorizontal, lineStyle, lineColor, lineTint); }


/**
 * @details assemble a string summary about the fill
 */
std::string XLBorder::summary() const
{
    using namespace std::literals::string_literals;
    std::string lineInfo{};
    lineInfo += ", left: "s + left().summary() + ", right: " + right().summary()
              + ", top: "s + top().summary() + ", bottom: " + bottom().summary()
              + ", diagonal: "s + diagonal().summary()
              + ", vertical: "s + vertical().summary()
              + ", horizontal: "s + horizontal().summary();
    return   "diagonalUp: "s   + (diagonalUp()   ? "true"s : "false"s)
         + ", diagonalDown: "s + (diagonalDown() ? "true"s : "false"s)
         + ", outline: "s      + (outline()      ? "true"s : "false"s)
         + lineInfo;
}


// ===== XLBorders, parent of XLBorder

/**
 * @details Constructor. Initializes an empty XLBorders object
 */
XLBorders::XLBorders() : m_bordersNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLBorders object.
 */
XLBorders::XLBorders(const XMLNode& borders)
    : m_bordersNode(std::make_unique<XMLNode>(borders))
{
    // initialize XLBorders entries and m_borders here
    XMLNode node = borders.first_child_of_type(pugi::node_element);
    while (not node.empty()) {
        std::string nodeName = node.name();
        if (nodeName == "border")
            m_borders.push_back(XLBorder(node));
        else
            std::cerr << "WARNING: XLBorders constructor: unknown subnode " << nodeName << std::endl;
        node = node.next_sibling_of_type(pugi::node_element);
    }
}

XLBorders::~XLBorders()
{
    m_borders.clear(); // delete vector with all children
}

XLBorders::XLBorders(const XLBorders& other)
    : m_bordersNode(std::make_unique<XMLNode>(*other.m_bordersNode)),
      m_borders(other.m_borders)
{
    // std::cout << __func__ << " copy constructor" << std::endl;
}

XLBorders::XLBorders(XLBorders&& other)
    : m_bordersNode(std::move(other.m_bordersNode)),
      m_borders(std::move(other.m_borders))
{
    // std::cout << __func__ << " move constructor" << std::endl;
}


/**
 * @details Copy assignment operator
 */
XLBorders& XLBorders::operator=(const XLBorders& other)
{
    // std::cout << "XLBorders::" << __func__ << " copy assignment" << std::endl;
    if (&other != this) {
        *m_bordersNode = *other.m_bordersNode;
        m_borders.clear();
        m_borders = other.m_borders;
    }
    return *this;
}

/**
 * @details Returns the amount of border descriptions held by the class
 */
size_t XLBorders::count() const { return m_borders.size(); }

/**
 * @details fetch XLBorder from m_borders by index
 */
XLBorder XLBorders::borderByIndex(XLStyleIndex index) const
{
    if (index >= m_borders.size()) {
        using namespace std::literals::string_literals;
        throw XLException("XLBorders::"s + __func__ + ": attempted to access index "s + std::to_string(index) + " with count "s + std::to_string(m_borders.size()));
    }
    return m_borders.at(index);
}

/**
 * @details append a new XLBorder to m_borders and m_bordersNode, based on copyFrom
 */
XLStyleIndex XLBorders::create(XLBorder copyFrom, std::string styleEntriesPrefix)
{
    XLStyleIndex index = count();    // index for the border to be created
    XMLNode newNode{};               // scope declaration

    // ===== Append new node prior to final whitespaces, if any
    XMLNode lastStyle = m_bordersNode->last_child_of_type(pugi::node_element);
    if (lastStyle.empty()) newNode = m_bordersNode->prepend_child("border");
    else                   newNode = m_bordersNode->insert_child_after("border", lastStyle);
    if (newNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLBorders::"s + __func__ + ": failed to append a new border node"s);
    }
    if (styleEntriesPrefix.length() > 0)    // if a whitespace prefix is configured
        m_bordersNode->insert_child_before(pugi::node_pcdata, newNode).set_value(styleEntriesPrefix.c_str());    // prefix the new node with styleEntriesPrefix

    XLBorder newBorder(newNode);
    if (copyFrom.m_borderNode->empty()) {    // if no template is given
        // ===== Create a border with default values
        // TODO: implement border defaults
        // newBorder.setProperty(defaultValue);
        // ...
    }
    else
        copyXMLNode(newNode, *copyFrom.m_borderNode);    // will use copyFrom as template, does nothing if copyFrom is empty

    m_borders.push_back(newBorder);
    appendAndSetAttribute(*m_bordersNode, "count", std::to_string(m_borders.size())); // update array count in XML
    return index;
}


/**
 * @details Constructor. Initializes an empty XLAlignment object
 */
XLAlignment::XLAlignment() : m_alignmentNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLAlignment object.
 */
XLAlignment::XLAlignment(const XMLNode& node) : m_alignmentNode(std::make_unique<XMLNode>(node)) {}

XLAlignment::~XLAlignment() = default;

XLAlignment::XLAlignment(const XLAlignment& other)
    : m_alignmentNode(std::make_unique<XMLNode>(*other.m_alignmentNode))
{}

XLAlignment& XLAlignment::operator=(const XLAlignment& other)
{
    if (&other != this) *m_alignmentNode = *other.m_alignmentNode;
    return *this;
}

/**
 * @details Returns the horizontal alignment style
 */
XLAlignmentStyle XLAlignment::horizontal() const { return XLAlignmentStyleFromString(m_alignmentNode->attribute("horizontal").value()); }

/**
 * @details Returns the vertical alignment style
 */
XLAlignmentStyle XLAlignment::vertical() const { return XLAlignmentStyleFromString(m_alignmentNode->attribute("vertical").value()); }

/**
 * @details Returns the text rotation
 */
uint16_t XLAlignment::textRotation() const { return m_alignmentNode->attribute("textRotation").as_uint(); }

/**
 * @details check if text wrapping is enabled
 */
bool XLAlignment::wrapText() const { return m_alignmentNode->attribute("wrapText").as_bool(); }

/**
 * @details Returns the indent setting
 */
uint32_t XLAlignment::indent() const { return m_alignmentNode->attribute("indent").as_uint(); }

/**
 * @details Returns the relative indent setting
 */
int32_t XLAlignment::relativeIndent() const { return m_alignmentNode->attribute("relativeIndent").as_int(); }

/**
 * @details check if justification of last line is enabled
 */
bool XLAlignment::justifyLastLine() const { return m_alignmentNode->attribute("justifyLastLine").as_bool(); }

/**
 * @details check if shrink to fit is enabled
 */
bool XLAlignment::shrinkToFit() const { return m_alignmentNode->attribute("shrinkToFit").as_bool(); }

/**
 * @details Returns the reading order setting
 */
uint32_t XLAlignment::readingOrder() const { return m_alignmentNode->attribute("readingOrder").as_uint(); }

/**
 * @details Setter functions
 */
bool XLAlignment::setHorizontal     (XLAlignmentStyle newStyle) { return appendAndSetAttribute(*m_alignmentNode, "horizontal",      XLAlignmentStyleToString(newStyle).c_str()).empty() == false; }
bool XLAlignment::setVertical       (XLAlignmentStyle newStyle) { return appendAndSetAttribute(*m_alignmentNode, "vertical",        XLAlignmentStyleToString(newStyle).c_str()).empty() == false; }
bool XLAlignment::setTextRotation   (uint16_t newRotation)      { return appendAndSetAttribute(*m_alignmentNode, "textRotation",    std::to_string(newRotation)               ).empty() == false; }
bool XLAlignment::setWrapText       (bool set)                  { return appendAndSetAttribute(*m_alignmentNode, "wrapText",        (set ? "true" : "false")                  ).empty() == false; }
bool XLAlignment::setIndent         (uint32_t newIndent)        { return appendAndSetAttribute(*m_alignmentNode, "indent",          std::to_string(newIndent)                 ).empty() == false; }
bool XLAlignment::setRelativeIndent (int32_t newRelativeIndent) { return appendAndSetAttribute(*m_alignmentNode, "relativeIndent",  std::to_string(newRelativeIndent)         ).empty() == false; }
bool XLAlignment::setJustifyLastLine(bool set)                  { return appendAndSetAttribute(*m_alignmentNode, "justifyLastLine", (set ? "true" : "false")                  ).empty() == false; }
bool XLAlignment::setShrinkToFit    (bool set)                  { return appendAndSetAttribute(*m_alignmentNode, "shrinkToFit",     (set ? "true" : "false")                  ).empty() == false; }
bool XLAlignment::setReadingOrder   (uint32_t newReadingOrder)  { return appendAndSetAttribute(*m_alignmentNode, "readingOrder",    std::to_string(newReadingOrder)           ).empty() == false; }

/**
 * @details assemble a string summary about the fill
 */
std::string XLAlignment::summary() const
{
    using namespace std::literals::string_literals;
    return "alignment horizontal="s + XLAlignmentStyleToString(horizontal()) + ", vertical="s + XLAlignmentStyleToString(vertical())
         + ", textRotation="s + std::to_string(textRotation())
         + ", wrapText="s + (wrapText() ? "true" : "false")
         + ", indent="s + std::to_string(indent())
         + ", relativeIndent="s + std::to_string(relativeIndent())
         + ", justifyLastLine="s + (justifyLastLine() ? "true" : "false")
         + ", shrinkToFit="s + (shrinkToFit() ? "true" : "false")
         + ", readingOrder="s + XLReadingOrderToString(readingOrder());
}


/**
 * @details Constructor. Initializes an empty XLCellFormat object
 */
XLCellFormat::XLCellFormat() : m_cellFormatNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLCellFormat object.
 */
XLCellFormat::XLCellFormat(const XMLNode& node, bool permitXfId)
 : m_cellFormatNode(std::make_unique<XMLNode>(node)),
   m_permitXfId(permitXfId)
{}

XLCellFormat::~XLCellFormat() = default;

XLCellFormat::XLCellFormat(const XLCellFormat& other)
    : m_cellFormatNode(std::make_unique<XMLNode>(*other.m_cellFormatNode)),
      m_permitXfId(other.m_permitXfId)
{}

XLCellFormat& XLCellFormat::operator=(const XLCellFormat& other)
{
    if (&other != this) {
        *m_cellFormatNode = *other.m_cellFormatNode;
        m_permitXfId = other.m_permitXfId;
    }
    return *this;
}

/**
 * @details determines the numberFormatId
 * @note returns XLInvalidUInt32 if attribute is not defined / set / empty
 */
uint32_t XLCellFormat::numberFormatId() const { return m_cellFormatNode->attribute("numFmtId").as_uint(XLInvalidUInt32); }

/**
 * @details determines the fontIndex
 * @note returns XLInvalidStyleIndex if attribute is not defined / set / empty
 */
XLStyleIndex XLCellFormat::fontIndex() const { return m_cellFormatNode->attribute("fontId").as_uint(XLInvalidStyleIndex); }

/**
 * @details determines the fillIndex
 * @note returns XLInvalidStyleIndex if attribute is not defined / set / empty
 */
XLStyleIndex XLCellFormat::fillIndex() const { return m_cellFormatNode->attribute("fillId").as_uint(XLInvalidStyleIndex); }

/**
 * @details determines the borderIndex
 * @note returns XLInvalidStyleIndex if attribute is not defined / set / empty
 */
XLStyleIndex XLCellFormat::borderIndex() const { return m_cellFormatNode->attribute("borderId").as_uint(XLInvalidStyleIndex); }

/**
 * @details Returns the xfId value
 */
XLStyleIndex XLCellFormat::xfId() const
{
    if (m_permitXfId) return m_cellFormatNode->attribute("xfId").as_uint(XLInvalidStyleIndex);
    throw XLException("XLCellFormat::xfId not permitted when m_permitXfId is false");
}

/**
 * @details determine the applyNumberFormat,applyFont,applyFill,applyBorder,applyAlignment,applyProtection status
 */
bool XLCellFormat::applyNumberFormat() const { return m_cellFormatNode->attribute("applyNumberFormat").as_bool(); }
bool XLCellFormat::applyFont()         const { return m_cellFormatNode->attribute("applyFont").as_bool();         }
bool XLCellFormat::applyFill()         const { return m_cellFormatNode->attribute("applyFill").as_bool();         }
bool XLCellFormat::applyBorder()       const { return m_cellFormatNode->attribute("applyBorder").as_bool();       }
bool XLCellFormat::applyAlignment()    const { return m_cellFormatNode->attribute("applyAlignment").as_bool();    }
bool XLCellFormat::applyProtection()   const { return m_cellFormatNode->attribute("applyProtection").as_bool();   }

/**
 * @details determine the quotePrefix, pivotButton status
 */
bool XLCellFormat::quotePrefix()       const { return m_cellFormatNode->attribute("quotePrefix").as_bool();       }
bool XLCellFormat::pivotButton()       const { return m_cellFormatNode->attribute("pivotButton").as_bool();       }

/**
 * @details determine the protection:locked,protection:hidden status
 */
bool XLCellFormat::locked() const { return m_cellFormatNode->child("protection").attribute("locked").as_bool(); }
bool XLCellFormat::hidden() const { return m_cellFormatNode->child("protection").attribute("hidden").as_bool(); }

/**
 * @details fetch alignment object
 */
XLAlignment XLCellFormat::alignment(bool createIfMissing) const
{
    XMLNode nodeAlignment = m_cellFormatNode->child("alignment");
    if (nodeAlignment.empty() && createIfMissing)
        nodeAlignment = m_cellFormatNode->append_child("alignment");
    return XLAlignment(nodeAlignment);
}

/**
 * @details Setter functions
 */
bool XLCellFormat::setNumberFormatId   (uint32_t newNumFmtId)        { return appendAndSetAttribute(*m_cellFormatNode, "numFmtId", std::to_string(newNumFmtId   )).empty() == false; }
bool XLCellFormat::setFontIndex        (XLStyleIndex newXfIndex)     { return appendAndSetAttribute(*m_cellFormatNode, "fontId",   std::to_string(newXfIndex    )).empty() == false; }
bool XLCellFormat::setFillIndex        (XLStyleIndex newFillIndex)   { return appendAndSetAttribute(*m_cellFormatNode, "fillId",   std::to_string(newFillIndex  )).empty() == false; }
bool XLCellFormat::setBorderIndex      (XLStyleIndex newBorderIndex) { return appendAndSetAttribute(*m_cellFormatNode, "borderId", std::to_string(newBorderIndex)).empty() == false; }
bool XLCellFormat::setXfId             (XLStyleIndex newXfId)
{
    if (m_permitXfId)                                                  return appendAndSetAttribute(*m_cellFormatNode, "xfId",     std::to_string(newXfId       )).empty() == false;
    throw XLException("XLCellFormat::setXfId not permitted when m_permitXfId is false");
}
bool XLCellFormat::setApplyNumberFormat(bool set)            { return appendAndSetAttribute    (*m_cellFormatNode, "applyNumberFormat",           (set ? "true" : "false")).empty() == false; }
bool XLCellFormat::setApplyFont        (bool set)            { return appendAndSetAttribute    (*m_cellFormatNode, "applyFont",                   (set ? "true" : "false")).empty() == false; }
bool XLCellFormat::setApplyFill        (bool set)            { return appendAndSetAttribute    (*m_cellFormatNode, "applyFill",                   (set ? "true" : "false")).empty() == false; }
bool XLCellFormat::setApplyBorder      (bool set)            { return appendAndSetAttribute    (*m_cellFormatNode, "applyBorder",                 (set ? "true" : "false")).empty() == false; }
bool XLCellFormat::setApplyAlignment   (bool set)            { return appendAndSetAttribute    (*m_cellFormatNode, "applyAlignment",              (set ? "true" : "false")).empty() == false; }
bool XLCellFormat::setApplyProtection  (bool set)            { return appendAndSetAttribute    (*m_cellFormatNode, "applyProtection",             (set ? "true" : "false")).empty() == false; }
bool XLCellFormat::setQuotePrefix      (bool set)            { return appendAndSetAttribute    (*m_cellFormatNode, "quotePrefix",                 (set ? "true" : "false")).empty() == false; }
bool XLCellFormat::setPivotButton      (bool set)            { return appendAndSetAttribute    (*m_cellFormatNode, "pivotButton",                 (set ? "true" : "false")).empty() == false; }
bool XLCellFormat::setLocked           (bool set)            { return appendAndSetNodeAttribute(*m_cellFormatNode, "protection",        "locked", (set ? "true" : "false")).empty() == false; }
bool XLCellFormat::setHidden           (bool set)            { return appendAndSetNodeAttribute(*m_cellFormatNode, "protection",        "hidden", (set ? "true" : "false")).empty() == false; }

/**
 * @brief Unsupported setter function
 */
bool XLCellFormat::setExtLst(XLUnsupportedElement const& newExtLst) { OpenXLSX::ignore(newExtLst); return false; }


/**
 * @details assemble a string summary about the cell format
 */
std::string XLCellFormat::summary() const
{
    using namespace std::literals::string_literals;
    return "numberFormatId="s + std::to_string(numberFormatId())
         + ", fontIndex="s + std::to_string(fontIndex())
         + ", fillIndex="s + std::to_string(fillIndex())
         + ", borderIndex="s + std::to_string(borderIndex())
         + ( m_permitXfId ? (", xfId: "s + std::to_string(xfId())) : ""s)
         + ", applyNumberFormat: "s + (applyNumberFormat() ? "true"s : "false"s)
         + ", applyFont: "s + (applyFont() ? "true"s : "false"s)
         + ", applyFill: "s + (applyFill() ? "true"s : "false"s)
         + ", applyBorder: "s + (applyBorder() ? "true"s : "false"s)
         + ", applyAlignment: "s + (applyAlignment() ? "true"s : "false"s)
         + ", applyProtection: "s + (applyProtection() ? "true"s : "false"s)
         + ", quotePrefix: "s + (quotePrefix() ? "true"s : "false"s)
         + ", pivotButton: "s + (pivotButton() ? "true"s : "false"s)
         + ", locked: "s + (locked() ? "true"s : "false"s)
         + ", hidden: "s + (hidden() ? "true"s : "false"s)
         + ", " + alignment().summary();
}


// ===== XLCellFormats, one parent of XLCellFormat (the other being XLCellFormats)

/**
 * @details Constructor. Initializes an empty XLCellFormats object
 */
XLCellFormats::XLCellFormats() : m_cellFormatsNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLCellFormats object.
 */
XLCellFormats::XLCellFormats(const XMLNode& cellStyleFormats, bool permitXfId)
    : m_cellFormatsNode(std::make_unique<XMLNode>(cellStyleFormats)),
      m_permitXfId(permitXfId)
{
    // initialize XLCellFormats entries and m_cellFormats here
    XMLNode node = cellStyleFormats.first_child_of_type(pugi::node_element);
    while (not node.empty()) {
        std::string nodeName = node.name();
        if (nodeName == "xf")
            m_cellFormats.push_back(XLCellFormat(node, m_permitXfId));
        else
            std::cerr << "WARNING: XLCellFormats constructor: unknown subnode " << nodeName << std::endl;
        node = node.next_sibling_of_type(pugi::node_element);
    }
}

XLCellFormats::~XLCellFormats()
{
    m_cellFormats.clear(); // delete vector with all children
}

XLCellFormats::XLCellFormats(const XLCellFormats& other)
    : m_cellFormatsNode(std::make_unique<XMLNode>(*other.m_cellFormatsNode)),
      m_cellFormats(other.m_cellFormats),
      m_permitXfId(other.m_permitXfId)
{
    // std::cout << __func__ << " copy constructor" << std::endl;
}

XLCellFormats::XLCellFormats(XLCellFormats&& other)
    : m_cellFormatsNode(std::move(other.m_cellFormatsNode)),
      m_cellFormats(std::move(other.m_cellFormats)),
      m_permitXfId(other.m_permitXfId)
{
    // std::cout << __func__ << " move constructor" << std::endl;
}


/**
 * @details Copy assignment operator
 */
XLCellFormats& XLCellFormats::operator=(const XLCellFormats& other)
{
    // std::cout << "XLCellFormats::" << __func__ << " copy assignment" << std::endl;
    if (&other != this) {
        *m_cellFormatsNode = *other.m_cellFormatsNode;
        m_cellFormats.clear();
        m_cellFormats = other.m_cellFormats;
        m_permitXfId = other.m_permitXfId;
    }
    return *this;
}

/**
 * @details Returns the amount of cell format descriptions held by the class
 */
size_t XLCellFormats::count() const { return m_cellFormats.size(); }

/**
 * @details fetch XLCellFormat from m_cellFormats by index
 */
XLCellFormat XLCellFormats::cellFormatByIndex(XLStyleIndex index) const
{
    if (index >= m_cellFormats.size()) {
        using namespace std::literals::string_literals;
        throw XLException("XLCellFormats::"s + __func__ + ": attempted to access index "s + std::to_string(index)
                        + " with count "s + std::to_string(m_cellFormats.size()));
    }
    return m_cellFormats.at(index);
}

/**
 * @details append a new XLCellFormat to m_cellFormats and m_cellFormatsNode, based on copyFrom
 */
XLStyleIndex XLCellFormats::create(XLCellFormat copyFrom, std::string styleEntriesPrefix)
{
    XLStyleIndex index = count();    // index for the cell format to be created
    XMLNode newNode{};               // scope declaration

    // ===== Append new node prior to final whitespaces, if any
    XMLNode lastStyle = m_cellFormatsNode->last_child_of_type(pugi::node_element);
    if (lastStyle.empty()) newNode = m_cellFormatsNode->prepend_child("xf");
    else                   newNode = m_cellFormatsNode->insert_child_after("xf", lastStyle);
    if (newNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLCellFormats::"s + __func__ + ": failed to append a new xf node"s);
    }
    if (styleEntriesPrefix.length() > 0)    // if a whitespace prefix is configured
        m_cellFormatsNode->insert_child_before(pugi::node_pcdata, newNode).set_value(styleEntriesPrefix.c_str());    // prefix the new node with styleEntriesPrefix

    XLCellFormat newCellFormat(newNode, m_permitXfId);
    if (copyFrom.m_cellFormatNode->empty()) {    // if no template is given
        // ===== Create a cell format with default values
        // TODO: implement cell format defaults
        // newStyle.setProperty(defaultValue);
        // ...
    }
    else
        copyXMLNode(newNode, *copyFrom.m_cellFormatNode); // will use copyFrom as template, does nothing if copyFrom is empty

    m_cellFormats.push_back(newCellFormat);
    appendAndSetAttribute(*m_cellFormatsNode, "count", std::to_string(m_cellFormats.size())); // update array count in XML
    return index;
}


/**
 * @details Constructor. Initializes an empty XLCellStyle object
 */
XLCellStyle::XLCellStyle() : m_cellStyleNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLCellStyle object.
 */
XLCellStyle::XLCellStyle(const XMLNode& node) : m_cellStyleNode(std::make_unique<XMLNode>(node)) {}

XLCellStyle::~XLCellStyle() = default;

XLCellStyle::XLCellStyle(const XLCellStyle& other)
    : m_cellStyleNode(std::make_unique<XMLNode>(*other.m_cellStyleNode))
{}

XLCellStyle& XLCellStyle::operator=(const XLCellStyle& other)
{
    if (&other != this) *m_cellStyleNode = *other.m_cellStyleNode;
    return *this;
}

/**
 * @details Returns the style empty status
 */
bool XLCellStyle::empty() const { return m_cellStyleNode->empty(); }

/**
 * @details Getter functions
 */
std::string  XLCellStyle::name         () const { return m_cellStyleNode->attribute("name"         ).value();                      }
XLStyleIndex XLCellStyle::xfId         () const { return m_cellStyleNode->attribute("xfId"         ).as_uint(XLInvalidStyleIndex); }
uint32_t     XLCellStyle::builtinId    () const { return m_cellStyleNode->attribute("builtinId"    ).as_uint(XLInvalidUInt32);     }
uint32_t     XLCellStyle::outlineStyle () const { return m_cellStyleNode->attribute("iLevel"       ).as_uint(XLInvalidUInt32);     }
bool         XLCellStyle::hidden       () const { return m_cellStyleNode->attribute("hidden"       ).as_bool();                    }
bool         XLCellStyle::customBuiltin() const { return m_cellStyleNode->attribute("customBuiltin").as_bool();                    }

/**
 * @details Setter functions
 */
bool XLCellStyle::setName         (std::string newName)      { return appendAndSetAttribute(*m_cellStyleNode, "name",          newName).empty() == false;                         }
bool XLCellStyle::setXfId         (XLStyleIndex newXfId)     { return appendAndSetAttribute(*m_cellStyleNode, "xfId",          std::to_string(newXfId)).empty() == false;         }
bool XLCellStyle::setBuiltinId    (uint32_t newBuiltinId)    { return appendAndSetAttribute(*m_cellStyleNode, "builtinId",     std::to_string(newBuiltinId)).empty() == false;    }
bool XLCellStyle::setOutlineStyle (uint32_t newOutlineStyle) { return appendAndSetAttribute(*m_cellStyleNode, "iLevel",        std::to_string(newOutlineStyle)).empty() == false; }
bool XLCellStyle::setHidden       (bool set)                 { return appendAndSetAttribute(*m_cellStyleNode, "hidden",        (set ? "true" : "false")).empty() == false;        }
bool XLCellStyle::setCustomBuiltin(bool set)                 { return appendAndSetAttribute(*m_cellStyleNode, "customBuiltin", (set ? "true" : "false")).empty() == false;        }

/**
 * @brief Unsupported setter function
 */
bool XLCellStyle::setExtLst(XLUnsupportedElement const& newExtLst) { OpenXLSX::ignore(newExtLst); return false; }

/**
 * @details assemble a string summary about the cell style
 */
std::string XLCellStyle::summary() const
{
    using namespace std::literals::string_literals;
    uint32_t iLevel = outlineStyle();
    return "name="s + name()
         + ", xfId="s + std::to_string(xfId())
         + ", builtinId="s + std::to_string(builtinId())
         + (iLevel != XLInvalidUInt32 ? ", iLevel="s + std::to_string(outlineStyle()) : ""s)
         + (hidden() ? ", hidden=true"s : ""s)
         + (customBuiltin() ? ", customBuiltin=true"s : ""s);
}

// ===== XLCellStyles, parent of XLCellStyle

/**
 * @details Constructor. Initializes an empty XLCellStyles object
 */
XLCellStyles::XLCellStyles() : m_cellStylesNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLCellStyles object.
 */
XLCellStyles::XLCellStyles(const XMLNode& cellStyles)
    : m_cellStylesNode(std::make_unique<XMLNode>(cellStyles))
{
    // initialize XLCellStyles entries and m_cellStyles here
    XMLNode node = cellStyles.first_child_of_type(pugi::node_element);
    while (not node.empty()) {
        std::string nodeName = node.name();
        // std::cout << "XLCellStyles constructor, node name is " << nodeName << std::endl;
        if (nodeName == "cellStyle")
            m_cellStyles.push_back(XLCellStyle(node));
        else
            std::cerr << "WARNING: XLCellStyles constructor: unknown subnode " << nodeName << std::endl;
        node = node.next_sibling_of_type(pugi::node_element);
    }
}

XLCellStyles::~XLCellStyles()
{
    m_cellStyles.clear(); // delete vector with all children
}

XLCellStyles::XLCellStyles(const XLCellStyles& other)
    : m_cellStylesNode(std::make_unique<XMLNode>(*other.m_cellStylesNode)),
      m_cellStyles(other.m_cellStyles)
{
    // std::cout << __func__ << " copy constructor" << std::endl;
}

XLCellStyles::XLCellStyles(XLCellStyles&& other)
    : m_cellStylesNode(std::move(other.m_cellStylesNode)),
      m_cellStyles(std::move(other.m_cellStyles))
{
    // std::cout << __func__ << " move constructor" << std::endl;
}


/**
 * @details Copy assignment operator
 */
XLCellStyles& XLCellStyles::operator=(const XLCellStyles& other)
{
    // std::cout << "XLCellStyles::" << __func__ << " copy assignment" << std::endl;
    if (&other != this) {
        *m_cellStylesNode = *other.m_cellStylesNode;
        m_cellStyles.clear();
        m_cellStyles = other.m_cellStyles;
    }
    return *this;
}

/**
 * @details Returns the amount of numberFormats held by the class
 */
size_t XLCellStyles::count() const { return m_cellStyles.size(); }

/**
 * @details fetch XLCellStyle from m_cellStyles by index
 */
XLCellStyle XLCellStyles::cellStyleByIndex(XLStyleIndex index) const
{
    if (index >= m_cellStyles.size()) {
        using namespace std::literals::string_literals;
        throw XLException("XLCellStyles::"s + __func__ + ": attempted to access index "s + std::to_string(index)
                        + " with count "s + std::to_string(m_cellStyles.size()));
    }
    return m_cellStyles.at(index);
}

/**
 * @details append a new XLCellStyle to m_cellStyles and m_cellStyleNode, based on copyFrom
 */
XLStyleIndex XLCellStyles::create(XLCellStyle copyFrom, std::string styleEntriesPrefix)
{
    XLStyleIndex index = count();    // index for the cell style to be created
    XMLNode newNode{};               // scope declaration

    // ===== Append new node prior to final whitespaces, if any
    XMLNode lastStyle = m_cellStylesNode->last_child_of_type(pugi::node_element);
    if (lastStyle.empty()) newNode = m_cellStylesNode->prepend_child("cellStyle");
    else                   newNode = m_cellStylesNode->insert_child_after("cellStyle", lastStyle);
    if (newNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLCellStyles::"s + __func__ + ": failed to append a new cellStyle node"s);
    }
    if (styleEntriesPrefix.length() > 0)    // if a whitespace prefix is configured
        m_cellStylesNode->insert_child_before(pugi::node_pcdata, newNode).set_value(styleEntriesPrefix.c_str());    // prefix the new node with styleEntriesPrefix

    XLCellStyle newCellStyle(newNode);
    if (copyFrom.m_cellStyleNode->empty()) {    // if no template is given
        // ===== Create a cell style with default values
        // TODO: implement cell style defaults
        // newCellStyle.setProperty(defaultValue);
        // ...
    }
    else
        copyXMLNode(newNode, *copyFrom.m_cellStyleNode); // will use copyFrom as template, does nothing if copyFrom is empty

    m_cellStyles.push_back(newCellStyle);
    appendAndSetAttribute(*m_cellStylesNode, "count", std::to_string(m_cellStyles.size())); // update array count in XML
    return index;
}


/**
 * @details Creates an XLStyles object, which will initialize from the given xmlData
 */
XLStyles::XLStyles(XLXmlData* xmlData, std::string stylesPrefix)
    : XLXmlFile(xmlData)
{
    XMLDocument & doc = xmlDocument();
    if (doc.document_element().empty())    // handle a bad (no document element) xl/styles.xml
        doc.load_string(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
                "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
                "</styleSheet>",
                pugi_parse_settings
        );

    XMLNode node = doc.document_element().first_child_of_type(pugi::node_element);
    while (not node.empty()) {
        XLStylesEntryType e = XLStylesEntryTypeFromString (node.name());
        // std::cout << "node.name() is " << node.name() << ", resulting XLStylesEntryType is " << std::to_string(e) << std::endl;
        switch (e) {
            case XLStylesNumberFormats:
                // std::cout << "found XLStylesNumberFormats, node name is " << node.name() << std::endl;
                m_numberFormats = std::make_unique<XLNumberFormats>(node);
                break;
            case XLStylesFonts:
                // std::cout << "found XLStylesFonts, node name is " << node.name() << std::endl;
                m_fonts = std::make_unique<XLFonts>(node);
                break;
            case XLStylesFills:
                // std::cout << "found XLStylesFills, node name is " << node.name() << std::endl;
                m_fills = std::make_unique<XLFills>(node);
                break;
            case XLStylesBorders:
                // std::cout << "found XLStylesBorders, node name is " << node.name() << std::endl;
                m_borders = std::make_unique<XLBorders>(node);
                break;
            case XLStylesCellStyleFormats:
                // std::cout << "found XLStylesCellStyleFormats, node name is " << node.name() << std::endl;
                m_cellStyleFormats = std::make_unique<XLCellFormats>(node);
                break;
            case XLStylesCellFormats:
                // std::cout << "found XLStylesCellFormats, node name is " << node.name() << std::endl;
                m_cellFormats = std::make_unique<XLCellFormats>(node, XLPermitXfID);
                break;
            case XLStylesCellStyles:
                // std::cout << "found XLStylesCellStyles, node name is " << node.name() << std::endl;
                m_cellStyles = std::make_unique<XLCellStyles>(node);
                break;
            case XLStylesColors:      [[fallthrough]];
            case XLStylesFormats:     [[fallthrough]];
            case XLStylesTableStyles: [[fallthrough]];
            case XLStylesExtLst:
                std::cout << "XLStyles: Ignoring currently unsupported <" << XLStylesEntryTypeToString(e) + "> node" << std::endl;
                break;
            case XLStylesInvalid: [[fallthrough]];
            default:
                std::cerr << "WARNING: XLStyles constructor: invalid styles subnode \"" << node.name() << "\"" << std::endl;
                break;
        }
        node = node.next_sibling_of_type(pugi::node_element);
    }

    // ===== Fallbacks: create root style nodes (in reverse order, using prepend_child)
    if (!m_cellStyles) {
        node = doc.document_element().prepend_child(XLStylesEntryTypeToString(XLStylesCellStyles).c_str());
        wrapNode (doc.document_element(), node, stylesPrefix);
        m_cellStyles = std::make_unique<XLCellStyles>(node);
    }
    if (!m_cellFormats) {
        node = doc.document_element().prepend_child(XLStylesEntryTypeToString(XLStylesCellFormats).c_str());
        wrapNode (doc.document_element(), node, stylesPrefix);
        m_cellFormats = std::make_unique<XLCellFormats>(node, XLPermitXfID);
    }
    if (m_cellFormats->count() == 0) {    // if the cell formats array is empty
        // ===== Create a default empty cell format with ID 0 (== XLDefaultCellFormat) because when XLDefaultCellFormat
        //        is assigned to an XLRow, the intention is interpreted as "set the cell back to default formatting",
        //        which does not trigger setting the attribute customFormat="true".
        //       To avoid confusing the user when the first style created does not work for rows, and setting a cell's
        //        format back to XLDefaultCellFormat would cause an actual formatting (if assigned format ID 0), this
        //        initial entry with no properties is created and henceforth ignored
        m_cellFormats->create();
    }

    if (!m_cellStyleFormats) {
        node = doc.document_element().prepend_child(XLStylesEntryTypeToString(XLStylesCellStyleFormats).c_str());
        wrapNode (doc.document_element(), node, stylesPrefix);
        m_cellStyleFormats = std::make_unique<XLCellFormats>(node);
    }
    if (!m_borders) {
        node = doc.document_element().prepend_child(XLStylesEntryTypeToString(XLStylesBorders).c_str());
        wrapNode (doc.document_element(), node, stylesPrefix);
        m_borders = std::make_unique<XLBorders>(node);
    }
    if (!m_fills) {
        node = doc.document_element().prepend_child(XLStylesEntryTypeToString(XLStylesFills).c_str());
        wrapNode (doc.document_element(), node, stylesPrefix);
        m_fills = std::make_unique<XLFills>(node);
    }
    if (!m_fonts) {
        node = doc.document_element().prepend_child(XLStylesEntryTypeToString(XLStylesFonts).c_str());
        wrapNode (doc.document_element(), node, stylesPrefix);
        m_fonts = std::make_unique<XLFonts>(node);
    }
    if (!m_numberFormats) {
        node = doc.document_element().prepend_child(XLStylesEntryTypeToString(XLStylesNumberFormats).c_str());
        wrapNode (doc.document_element(), node, stylesPrefix);
        m_numberFormats = std::make_unique<XLNumberFormats>(node);
    }
}

XLStyles::~XLStyles()
{
}

/**
 * @details return a handle to the underlying number formats
 */
XLNumberFormats& XLStyles::numberFormats() const { return *m_numberFormats; }

/**
 * @details return a handle to the underlying fonts
 */
XLFonts& XLStyles::fonts() const { return *m_fonts; }

/**
 * @details return a handle to the underlying fills
 */
XLFills& XLStyles::fills() const { return *m_fills; }

/**
 * @details return a handle to the underlying borders
 */
XLBorders& XLStyles::borders() const { return *m_borders; }

/**
 * @details return a handle to the underlying cell style formats
 */
XLCellFormats& XLStyles::cellStyleFormats() const { return *m_cellStyleFormats; }

/**
 * @details return a handle to the underlying cell formats
 */
XLCellFormats& XLStyles::cellFormats() const { return *m_cellFormats; }

/**
 * @details return a handle to the underlying cell styles
 */
XLCellStyles& XLStyles::cellStyles() const { return *m_cellStyles; }

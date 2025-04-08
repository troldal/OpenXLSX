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
#include <algorithm> // std::max
#include <cctype>    // std::isdigit (issue #330)
#include <limits>    // std::numeric_limits
#include <map>       // std::multimap
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLCellRange.hpp"
#include "XLDocument.hpp"
#include "XLMergeCells.hpp"
#include "XLSheet.hpp"
#include "utilities/XLUtilities.hpp"

using namespace OpenXLSX;

namespace OpenXLSX
{
    // // Forward declaration. Implementation is in the XLUtilities.hpp file
    // XMLNode getRowNode(XMLNode sheetDataNode, uint32_t rowNumber);

    /**
     * @brief Function for setting tab color.
     * @param xmlDocument XMLDocument object
     * @param color Thr color to set
     */
    void setTabColor(const XMLDocument& xmlDocument, const XLColor& color)
    {
        if (!xmlDocument.document_element().child("sheetPr")) xmlDocument.document_element().prepend_child("sheetPr");

        if (!xmlDocument.document_element().child("sheetPr").child("tabColor"))
            xmlDocument.document_element().child("sheetPr").prepend_child("tabColor");

        auto colorNode = xmlDocument.document_element().child("sheetPr").child("tabColor");
        for (auto attr : colorNode.attributes()) colorNode.remove_attribute(attr);

        colorNode.prepend_attribute("rgb").set_value(color.hex().c_str());
    }

    /**
     * @brief Set the tab selected property to desired value
     * @param xmlDocument
     * @param selected
     */
    void setTabSelected(const XMLDocument& xmlDocument, bool selected)
    {    // 2024-04-30: whitespace support
        unsigned int value       = (selected ? 1 : 0);
        XMLNode      sheetView   = xmlDocument.document_element().child("sheetViews").first_child_of_type(pugi::node_element);
        XMLAttribute tabSelected = sheetView.attribute("tabSelected");
        if (tabSelected.empty())
            tabSelected = sheetView.prepend_attribute("tabSelected");    // BUGFIX 2025-03-15 issue #337: assign tabSelected value with newly created attribute if it didn't exist
        tabSelected.set_value(value);
    }

    /**
     * @brief Function for checking if the tab is selected.
     * @param xmlDocument
     * @return
     * @note pugi::xml_attribute::as_bool defaults to false for a non-existing (= empty) attribute
     */
    bool tabIsSelected(const XMLDocument& xmlDocument)
    {    // 2024-04-30: whitespace support
        return xmlDocument.document_element()
            .child("sheetViews")
            .first_child_of_type(pugi::node_element)
            .attribute("tabSelected")
            .as_bool();    // BUGFIX 2024-05-01: .value() "0" was evaluating to true
    }


    /**
     * @brief get the correct XLCfType from the OOXML cfRule type attribute string
     * @param typeString the string as used in the OOXML
     * @return the corresponding XLCfType enum value
     */
    XLCfType XLCfTypeFromString(std::string const& typeString)
    {
        if (typeString == "expression")        return XLCfType::Expression;
        if (typeString == "cellIs")            return XLCfType::CellIs;
        if (typeString == "colorScale")        return XLCfType::ColorScale;
        if (typeString == "dataBar")           return XLCfType::DataBar;
        if (typeString == "iconSet")           return XLCfType::IconSet;
        if (typeString == "top10")             return XLCfType::Top10;
        if (typeString == "uniqueValues")      return XLCfType::UniqueValues;
        if (typeString == "duplicateValues")   return XLCfType::DuplicateValues;
        if (typeString == "containsText")      return XLCfType::ContainsText;
        if (typeString == "notContainsText")   return XLCfType::NotContainsText;
        if (typeString == "beginsWith")        return XLCfType::BeginsWith;
        if (typeString == "endsWith")          return XLCfType::EndsWith;
        if (typeString == "containsBlanks")    return XLCfType::ContainsBlanks;
        if (typeString == "notContainsBlanks") return XLCfType::NotContainsBlanks;
        if (typeString == "containsErrors")    return XLCfType::ContainsErrors;
        if (typeString == "notContainsErrors") return XLCfType::NotContainsErrors;
        if (typeString == "timePeriod")        return XLCfType::TimePeriod;
        if (typeString == "aboveAverage")      return XLCfType::AboveAverage;
        return XLCfType::Invalid;
    }

    /**
     * @brief inverse of XLCfTypeFromString
     * @param cfType the type for which to get the OOXML string
     */
    std::string XLCfTypeToString(XLCfType cfType)
    {
        switch (cfType) {
            case XLCfType::Expression:         return "expression";
            case XLCfType::CellIs:             return "cellIs";
            case XLCfType::ColorScale:         return "colorScale";
            case XLCfType::DataBar:            return "dataBar";
            case XLCfType::IconSet:            return "iconSet";
            case XLCfType::Top10:              return "top10";
            case XLCfType::UniqueValues:       return "uniqueValues";
            case XLCfType::DuplicateValues:    return "duplicateValues";
            case XLCfType::ContainsText:       return "containsText";
            case XLCfType::NotContainsText:    return "notContainsText";
            case XLCfType::BeginsWith:         return "beginsWith";
            case XLCfType::EndsWith:           return "endsWith";
            case XLCfType::ContainsBlanks:     return "containsBlanks";
            case XLCfType::NotContainsBlanks:  return "notContainsBlanks";
            case XLCfType::ContainsErrors:     return "containsErrors";
            case XLCfType::NotContainsErrors:  return "notContainsErrors";
            case XLCfType::TimePeriod:         return "timePeriod";
            case XLCfType::AboveAverage:       return "aboveAverage";
            case XLCfType::Invalid:            [[fallthrough]];
            default:                           return "(invalid)";
        }
    }

    /**
     * @brief get the correct XLCfOperator from the OOXML cfRule operator attribute string
     * @param operatorString the string as used in the OOXML
     * @return the corresponding XLCfOperator enum value
     */
    XLCfOperator XLCfOperatorFromString(std::string const& operatorString)
    {
        if (operatorString == "lessThan")           return XLCfOperator::LessThan;
        if (operatorString == "lessThanOrEqual")    return XLCfOperator::LessThanOrEqual;
        if (operatorString == "equal")              return XLCfOperator::Equal;
        if (operatorString == "notEqual")           return XLCfOperator::NotEqual;
        if (operatorString == "greaterThanOrEqual") return XLCfOperator::GreaterThanOrEqual;
        if (operatorString == "greaterThan")        return XLCfOperator::GreaterThan;
        if (operatorString == "between")            return XLCfOperator::Between;
        if (operatorString == "notBetween")         return XLCfOperator::NotBetween;
        if (operatorString == "containsText")       return XLCfOperator::ContainsText;
        if (operatorString == "notContains")        return XLCfOperator::NotContains;
        if (operatorString == "beginsWith")         return XLCfOperator::BeginsWith;
        if (operatorString == "endsWith")           return XLCfOperator::EndsWith;
        return XLCfOperator::Invalid;
    }

    /**
     * @brief inverse of XLCfOperatorFromString
     * @param cfOperator the XLCfOperator for which to get the OOXML string
     */
    std::string XLCfOperatorToString(XLCfOperator cfOperator)
    {
        switch (cfOperator) {
            case XLCfOperator::LessThan:            return "lessThan";
            case XLCfOperator::LessThanOrEqual:     return "lessThanOrEqual";
            case XLCfOperator::Equal:               return "equal";
            case XLCfOperator::NotEqual:            return "notEqual";
            case XLCfOperator::GreaterThanOrEqual:  return "greaterThanOrEqual";
            case XLCfOperator::GreaterThan:         return "greaterThan";
            case XLCfOperator::Between:             return "between";
            case XLCfOperator::NotBetween:          return "notBetween";
            case XLCfOperator::ContainsText:        return "containsText";
            case XLCfOperator::NotContains:         return "notContains";
            case XLCfOperator::BeginsWith:          return "beginsWith";
            case XLCfOperator::EndsWith:            return "endsWith";
            case XLCfOperator::Invalid:             [[fallthrough]];
            default:                                return "(invalid)";
        }
    }

    /**
     * @brief get the correct XLCfTimePeriod from the OOXML cfRule timePeriod attribute string
     * @param operatorString the string as used in the OOXML
     * @return the corresponding XLCfTimePeriod enum value
     */
    XLCfTimePeriod XLCfTimePeriodFromString(std::string const& timePeriodString)
    {
        if (timePeriodString == "today")     return XLCfTimePeriod::Today;
        if (timePeriodString == "yesterday") return XLCfTimePeriod::Yesterday;
        if (timePeriodString == "tomorrow")  return XLCfTimePeriod::Tomorrow;
        if (timePeriodString == "last7Days") return XLCfTimePeriod::Last7Days;
        if (timePeriodString == "thisMonth") return XLCfTimePeriod::ThisMonth;
        if (timePeriodString == "lastMonth") return XLCfTimePeriod::LastMonth;
        if (timePeriodString == "nextMonth") return XLCfTimePeriod::NextMonth;
        if (timePeriodString == "thisWeek")  return XLCfTimePeriod::ThisWeek;
        if (timePeriodString == "lastWeek")  return XLCfTimePeriod::LastWeek;
        if (timePeriodString == "nextWeek")  return XLCfTimePeriod::NextWeek;
        return XLCfTimePeriod::Invalid;
    }

    /**
     * @brief inverse of XLCfOperatorFromString
     * @param cfOperator the XLCfTimePeriod for which to get the OOXML string
     */
    std::string XLCfTimePeriodToString(XLCfTimePeriod cfTimePeriod)
    {
        switch (cfTimePeriod) {
            case XLCfTimePeriod::Today:     return "today";
            case XLCfTimePeriod::Yesterday: return "yesterday";
            case XLCfTimePeriod::Tomorrow:  return "tomorrow";
            case XLCfTimePeriod::Last7Days: return "last7Days";
            case XLCfTimePeriod::ThisMonth: return "thisMonth";
            case XLCfTimePeriod::LastMonth: return "lastMonth";
            case XLCfTimePeriod::NextMonth: return "nextMonth";
            case XLCfTimePeriod::ThisWeek:  return "thisWeek";
            case XLCfTimePeriod::LastWeek:  return "lastWeek";
            case XLCfTimePeriod::NextWeek:  return "nextWeek";
            case XLCfTimePeriod::Invalid:   [[fallthrough]];
            default:                        return "(invalid)";
        }
    }
}    // namespace OpenXLSX

// ========== XLSheet Member Functions

/**
 * @details The constructor begins by constructing an instance of its superclass, XLAbstractXMLFile. The default
 * sheet type is WorkSheet and the default sheet state is Visible.
 */
XLSheet::XLSheet(XLXmlData* xmlData) : XLXmlFile(xmlData)
{
    if (xmlData->getXmlType() == XLContentType::Worksheet)
        m_sheet = XLWorksheet(xmlData);
    else if (xmlData->getXmlType() == XLContentType::Chartsheet)
        m_sheet = XLChartsheet(xmlData);
    else
        throw XLInternalError("Invalid XML data.");
}

/**
 * @details This method uses visitor pattern to return the name() member variable of the underlying
 * sheet object (XLWorksheet or XLChartsheet).
 */
std::string XLSheet::name() const
{
    return std::visit([](auto&& arg) { return arg.name(); }, m_sheet);
}

/**
 * @details This method sets the name of the sheet to a new name, by calling the setName()
 * member function of the underlying sheet object (XLWorksheet or XLChartsheet).
 */
void XLSheet::setName(const std::string& name)
{
    std::visit([&](auto&& arg) { return arg.setName(name); }, m_sheet);
}

/**
 * @details This method uses visitor pattern to return the visibility() member variable of the underlying
 * sheet object (XLWorksheet or XLChartsheet).
 */
XLSheetState XLSheet::visibility() const
{
    return std::visit([](auto&& arg) { return arg.visibility(); }, m_sheet);
}

/**
 * @details This method sets the visibility state of the sheet, by calling the setVisibility()
 * member function of the underlying sheet object (XLWorksheet or XLChartsheet).
 */
void XLSheet::setVisibility(XLSheetState state)
{
    std::visit([&](auto&& arg) { return arg.setVisibility(state); }, m_sheet);
}

/**
 * @details This method uses visitor pattern to return the color() member variable of the underlying
 * sheet object (XLWorksheet or XLChartsheet).
 */
XLColor XLSheet::color() const
{
    return std::visit([](auto&& arg) { return arg.color(); }, m_sheet);
}

/**
 * @details This method sets the color of the sheet, by calling the setColor()
 * member function of the underlying sheet object (XLWorksheet or XLChartsheet).
 */
void XLSheet::setColor(const XLColor& color)
{
    std::visit([&](auto&& arg) { return arg.setColor(color); }, m_sheet);
}

/**
 * @details This reports the selection state of the sheet, by calling the isSelected()
 * member function of the underlying sheet object (XLWorksheet or XLChartsheet).
 */
bool XLSheet::isSelected() const
{
    return std::visit([](auto&& arg) { return arg.isSelected(); }, m_sheet);
}

/**
 * @details This method sets the selection state of the sheet, by calling the setSelected()
 * member function of the underlying sheet object (XLWorksheet or XLChartsheet).
 */
void XLSheet::setSelected(bool selected)
{
    std::visit([&](auto&& arg) { return arg.setSelected(selected); }, m_sheet);
}

/**
 * @details This reports the active state of the sheet, by calling the isActive()
 * member function of the underlying sheet object (XLWorksheet or XLChartsheet).
 */
bool XLSheet::isActive() const
{
    return std::visit([](auto&& arg) { return arg.isActive(); }, m_sheet);
}

/**
 * @details This method sets the active state of the sheet, by calling the setActive()
 * member function of the underlying sheet object (XLWorksheet or XLChartsheet).
 */
bool XLSheet::setActive()
{
    return std::visit([&](auto&& arg) { return arg.setActive(); }, m_sheet);
}

/**
 * @details Clones the sheet by calling the clone() method in the underlying sheet object
 * (XLWorksheet or XLChartsheet), using the visitor pattern.
 */
void XLSheet::clone(const std::string& newName)
{
    std::visit([&](auto&& arg) { arg.clone(newName); }, m_sheet);
}

/**
 * @details Get the index of the sheet, by calling the index() method in the underlying
 * sheet object (XLWorksheet or XLChartsheet), using the visitor pattern.
 */
uint16_t XLSheet::index() const
{
    return std::visit([](auto&& arg) { return arg.index(); }, m_sheet);
}

/**
 * @details This method sets the index of the sheet (i.e. move the sheet), by calling the setIndex()
 * member function of the underlying sheet object (XLWorksheet or XLChartsheet).
 */
void XLSheet::setIndex(uint16_t index)
{
    std::visit([&](auto&& arg) { arg.setIndex(index); }, m_sheet);
}

/**
 * @details Implicit conversion operator to XLWorksheet. Calls the get<>() member function, with XLWorksheet as
 * the template argument.
 */
XLSheet::operator XLWorksheet() const { return this->get<XLWorksheet>(); }

/**
 * @details Implicit conversion operator to XLChartsheet. Calls the get<>() member function, with XLChartsheet as
 * the template argument.
 */
XLSheet::operator XLChartsheet() const { return this->get<XLChartsheet>(); }

/**
 * @details Print the underlying XML using pugixml::xml_node::print
 */
void XLSheet::print(std::basic_ostream<char>& ostr) const { xmlDocument().document_element().print(ostr); }


// ========== BEGIN <conditionalFormatting> related member function definitions

// ========== XLCfRule member functions
/**
 * @details Constructor. Initializes an empty XLCfRule object
 */
XLCfRule::XLCfRule() : m_cfRuleNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLCfRule object.
 */
XLCfRule::XLCfRule(const XMLNode& node) : m_cfRuleNode(std::make_unique<XMLNode>(node)) {}

XLCfRule::XLCfRule(const XLCfRule& other)
    : m_cfRuleNode(std::make_unique<XMLNode>(*other.m_cfRuleNode))
{}

XLCfRule::~XLCfRule() = default;

XLCfRule& XLCfRule::operator=(const XLCfRule& other)
{
    if (&other != this) *m_cfRuleNode = *other.m_cfRuleNode;
    return *this;
}

/**
 * @details Returns the node empty status
 */
bool XLCfRule::empty() const { return m_cfRuleNode->empty(); }

/**
 * @details Element getter functions
 */
std::string XLCfRule::formula() const
{
    // XMLNode formulaTextNode = 
    return m_cfRuleNode->child("formula").first_child_of_type(pugi::node_pcdata).value();
}

/**
 * @details Unsupported element getter functions
 */
XLUnsupportedElement XLCfRule::colorScale() const { return XLUnsupportedElement(); } // <cfRule><colorScale>...</colorScale></cfRule>
XLUnsupportedElement XLCfRule::dataBar()    const { return XLUnsupportedElement(); } // <cfRule><dataBar>...</dataBar></cfRule>
XLUnsupportedElement XLCfRule::iconSet()    const { return XLUnsupportedElement(); } // <cfRule><iconSet>...</iconSet></cfRule>
XLUnsupportedElement XLCfRule::extLst()     const { return XLUnsupportedElement{}; } // <cfRule><extLst>...</extLst></cfRule>

/**
 * @details Attribute getter functions
 */
XLCfType       XLCfRule::type()         const { return XLCfTypeFromString      (m_cfRuleNode->attribute("type").value()                     ); }
XLStyleIndex   XLCfRule::dxfId()        const { return                          m_cfRuleNode->attribute("dxfId").as_uint(XLInvalidStyleIndex); }
uint16_t       XLCfRule::priority()     const { return                          m_cfRuleNode->attribute("priority").as_uint(XLPriorityNotSet); }
bool           XLCfRule::stopIfTrue()   const { return                          m_cfRuleNode->attribute("stopIfTrue").as_bool(false)         ; }
// TODO TBD: how does true default value manifest itself for aboveAverage? Evaluate as true if aboveAverage=""?
bool           XLCfRule::aboveAverage() const { return                          m_cfRuleNode->attribute("aboveAverage").as_bool(false)       ; }
bool           XLCfRule::percent()      const { return                          m_cfRuleNode->attribute("percent").as_bool(false)            ; }
bool           XLCfRule::bottom()       const { return                          m_cfRuleNode->attribute("bottom").as_bool(false)             ; }
XLCfOperator   XLCfRule::Operator()     const { return XLCfOperatorFromString  (m_cfRuleNode->attribute("operator").value()                 ); }
std::string    XLCfRule::text()         const { return                          m_cfRuleNode->attribute("text").value()                      ; }
XLCfTimePeriod XLCfRule::timePeriod()   const { return XLCfTimePeriodFromString(m_cfRuleNode->attribute("timePeriod").value()               ); }
uint16_t       XLCfRule::rank()         const { return                          m_cfRuleNode->attribute("rank").as_uint()                    ; }
int16_t        XLCfRule::stdDev()       const { return                          m_cfRuleNode->attribute("stdDev").as_int()                   ; }
bool           XLCfRule::equalAverage() const { return                          m_cfRuleNode->attribute("equalAverage").as_bool(false)       ; }

/**
* @brief Element setter functions
*/
bool XLCfRule::setFormula(std::string const& newFormula)
{
    XMLNode formula = appendAndGetNode(*m_cfRuleNode, "formula", m_nodeOrder);
    if (formula.empty()) return false;
    formula.remove_children(); // no-op if no children
    return formula.append_child(pugi::node_pcdata).set_value(newFormula.c_str());
}

/**
 * @brief Unsupported element setter function
 */
bool XLCfRule::setColorScale(XLUnsupportedElement const& newColorScale) { OpenXLSX::ignore(newColorScale); return false; }
bool XLCfRule::setDataBar(   XLUnsupportedElement const& newDataBar   ) { OpenXLSX::ignore(newDataBar   ); return false; }
bool XLCfRule::setIconSet(   XLUnsupportedElement const& newIconSet   ) { OpenXLSX::ignore(newIconSet   ); return false; }
bool XLCfRule::setExtLst(    XLUnsupportedElement const& newExtLst    ) { OpenXLSX::ignore(newExtLst    ); return false; }

/**
 * @details Attribute setter functions
 */
bool XLCfRule::setType        (XLCfType newType    )         { return appendAndSetAttribute(*m_cfRuleNode, "type",         XLCfTypeToString(        newType      )).empty() == false; }
bool XLCfRule::setDxfId       (XLStyleIndex newDxfId)        { return appendAndSetAttribute(*m_cfRuleNode, "dxfId",        std::to_string(          newDxfId     )).empty() == false; }
bool XLCfRule::setPriority    (uint16_t newPriority)         { return appendAndSetAttribute(*m_cfRuleNode, "priority",     std::to_string(          newPriority  )).empty() == false; }
// TODO TBD whether true / false work with MS Office booleans here, or if 1 and 0 are mandatory
bool XLCfRule::setStopIfTrue  (bool set)                     { return appendAndSetAttribute(*m_cfRuleNode, "stopIfTrue",   (set ? "true" : "false")               ).empty() == false; }
bool XLCfRule::setAboveAverage(bool set)                     { return appendAndSetAttribute(*m_cfRuleNode, "aboveAverage", (set ? "true" : "false")               ).empty() == false; }
bool XLCfRule::setPercent     (bool set)                     { return appendAndSetAttribute(*m_cfRuleNode, "percent",      (set ? "true" : "false")               ).empty() == false; }
bool XLCfRule::setBottom      (bool set)                     { return appendAndSetAttribute(*m_cfRuleNode, "bottom",       (set ? "true" : "false")               ).empty() == false; }
bool XLCfRule::setOperator    (XLCfOperator newOperator)     { return appendAndSetAttribute(*m_cfRuleNode, "operator",     XLCfOperatorToString(  newOperator    )).empty() == false; }
bool XLCfRule::setText        (std::string const& newText)   { return appendAndSetAttribute(*m_cfRuleNode, "text",                                newText.c_str() ).empty() == false; }
bool XLCfRule::setTimePeriod  (XLCfTimePeriod newTimePeriod) { return appendAndSetAttribute(*m_cfRuleNode, "timePeriod",   XLCfTimePeriodToString(newTimePeriod)  ).empty() == false; }
bool XLCfRule::setRank        (uint16_t newRank)             { return appendAndSetAttribute(*m_cfRuleNode, "rank",         std::to_string(          newRank      )).empty() == false; }
bool XLCfRule::setStdDev      (int16_t newStdDev)            { return appendAndSetAttribute(*m_cfRuleNode, "stdDev",       std::to_string(          newStdDev    )).empty() == false; }
bool XLCfRule::setEqualAverage(bool set)                     { return appendAndSetAttribute(*m_cfRuleNode, "equalAverage", (set ? "true" : "false")               ).empty() == false; }


/**
 * @details assemble a string summary about the conditional formatting
 */
std::string XLCfRule::summary() const
{
    using namespace std::literals::string_literals;
    XLStyleIndex dxfIndex = dxfId();
    return "formula is "s + formula()
         + ", type is "s + XLCfTypeToString(type())
         + ", dxfId is "s + (dxfIndex == XLInvalidStyleIndex ? "(invalid)"s : std::to_string(dxfId()))
         + ", priority is "s + std::to_string(priority())
         + ", stopIfTrue: "s + (stopIfTrue() ? "true"s : "false"s)
         + ", aboveAverage: "s + (aboveAverage() ? "true"s : "false"s)
         + ", percent: "s + (percent() ? "true"s : "false"s)
         + ", bottom: "s + (bottom() ? "true"s : "false"s)
         + ", operator is "s + XLCfOperatorToString(Operator())
         + ", text is \""s + text() + "\""s
         + ", timePeriod is "s + XLCfTimePeriodToString(timePeriod())
         + ", rank is "s + std::to_string(rank())
         + ", stdDev is "s + std::to_string(stdDev())
         + ", equalAverage: "s + (equalAverage() ? "true"s : "false"s)
    ;
}


// ========== XLCfRules member functions, parent of XLCfRule
/**
 * @details Constructor. Initializes an empty XLCfRules object
 */
XLCfRules::XLCfRules() : m_conditionalFormattingNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLCellStyle object.
 */
XLCfRules::XLCfRules(const XMLNode& node) : m_conditionalFormattingNode(std::make_unique<XMLNode>(node)) {}

XLCfRules::XLCfRules(const XLCfRules& other)
    : m_conditionalFormattingNode(std::make_unique<XMLNode>(*other.m_conditionalFormattingNode))
{}

XLCfRules::~XLCfRules() = default;

XLCfRules& XLCfRules::operator=(const XLCfRules& other)
{
    if (&other != this) *m_conditionalFormattingNode = *other.m_conditionalFormattingNode;
    return *this;
}

/**
 * @details Returns the node empty status
 */
bool XLCfRules::empty() const { return m_conditionalFormattingNode->empty(); }

/**
 * @details Returns the maximum numerical priority value that a cfRule is using (= lowest rule priority)
 */
uint16_t XLCfRules::maxPriorityValue() const
{
    XMLNode node = m_conditionalFormattingNode->first_child_of_type(pugi::node_element);
    while (not node.empty() && std::string(node.name()) != "cfRule")
        node = node.next_sibling_of_type(pugi::node_element);
    uint16_t maxPriority = XLPriorityNotSet;
    while (not node.empty() && std::string(node.name()) == "cfRule") {
        maxPriority = std::max(maxPriority, XLCfRule(node).priority());
        node = node.next_sibling_of_type(pugi::node_element);
    }
    return maxPriority;
}

/**
 * @details ensures that newPriority is unique (by incrementing all existing priorities >= newPriority),
 *  then assigns newPriority to the rule with cfRuleIndex
 * @note has a check for no-op if desired priority is already set
 */
bool XLCfRules::setPriority(size_t cfRuleIndex, uint16_t newPriority)
{
    XLCfRule affectedRule = cfRuleByIndex(cfRuleIndex);        // throws for cfRuleIndex out of bounds
    if (newPriority == affectedRule.priority()) return true;   // no-op if priority is already set

    XMLNode node = m_conditionalFormattingNode->first_child_of_type(pugi::node_element);
    while (not node.empty() && std::string(node.name()) != "cfRule") // skip past non cfRule elements
        node = node.next_sibling_of_type(pugi::node_element);

    // iterate once over rules to check if newPriority is in use by another rule
    XMLNode node2 = node;
    bool newPriorityFree = true;
    while (newPriorityFree && not node2.empty() && std::string(node2.name()) == "cfRule") {
        if (XLCfRule(node2).priority() == newPriority) newPriorityFree = false; // check if newPriority is in use
        node2 = node2.next_sibling_of_type(pugi::node_element);
    }

    if (!newPriorityFree) { // need to free up newPriority
        size_t index = 0;
        while (not node.empty() && std::string(node.name()) == "cfRule") { // loop over cfRule elements
            if (index != cfRuleIndex) { // for all rules that are not at cfRuleIndex: increase priority if >= newPriority
                XLCfRule rule(node);
                uint16_t prio = rule.priority();
                if (prio >= newPriority) rule.setPriority(prio + 1);
            }
            node = node.next_sibling_of_type(pugi::node_element);
            ++index;
        }
    }
    return affectedRule.setPriority(newPriority);
}

/**
 * @details sort cfRule entries by priority and assign a continuous sequence of values using a given increment
 * @note this mimics the BASIC functionality of line renumbering
 */
void XLCfRules::renumberPriorities(uint16_t increment)
{
    if (increment == 0)   // not allowed
        throw XLException("XLCfRules::renumberPriorities: increment must not be 0");

    std::multimap< uint16_t, XLCfRule > rules;

    XMLNode node = m_conditionalFormattingNode->first_child_of_type(pugi::node_element);
    while (not node.empty() && std::string(node.name()) != "cfRule") // skip past non cfRule elements
        node = node.next_sibling_of_type(pugi::node_element);

    while (not node.empty() && std::string(node.name()) == "cfRule") { // loop over cfRule elements
        XLCfRule rule(node);
        rules.insert(std::pair(rule.priority(), std::move(rule)));
        node = node.next_sibling_of_type(pugi::node_element);
    }

    if (rules.size() * increment > std::numeric_limits< uint16_t >::max()) { // first rule always gets assigned "1*increment"
        using namespace std::literals::string_literals;
        throw XLException("XLCfRules::renumberPriorities: amount of rules "s + std::to_string(rules.size())
        /**/              + " with given increment "s + std::to_string(increment) + " exceeds max range of uint16_t"s);
    }

    uint16_t prio = 0;
    for (auto& [key, rule]: rules) {
        prio += increment;
        rule.setPriority(prio);
    }
    rules.clear();
}

/**
 * @details Returns the amount of cfRule entries held by the class
 */
size_t XLCfRules::count() const
{
    XMLNode node = m_conditionalFormattingNode->first_child_of_type(pugi::node_element);
    while (not node.empty() && std::string(node.name()) != "cfRule")
        node = node.next_sibling_of_type(pugi::node_element);
    size_t count = 0;
    while (not node.empty() && std::string(node.name()) == "cfRule") {
        ++count;
        node = node.next_sibling_of_type(pugi::node_element);
    }
    return count;
}

/**
 * @details fetch XLCfRule from m_conditionalFormattingNode by index
 */
XLCfRule XLCfRules::cfRuleByIndex(size_t index) const
{
    XMLNode node = m_conditionalFormattingNode->first_child_of_type(pugi::node_element);
    while (not node.empty() && std::string(node.name()) != "cfRule")
        node = node.next_sibling_of_type(pugi::node_element);

    if (not node.empty()) {
        size_t count = 0;
        while (not node.empty() && std::string(node.name()) == "cfRule" && count < index) {
            ++count;
            node = node.next_sibling_of_type(pugi::node_element);
        }
        if (count == index && std::string(node.name()) == "cfRule")
            return XLCfRule(node);
    }
    using namespace std::literals::string_literals;
    throw XLException("XLCfRules::"s + __func__ + ": cfRule with index "s + std::to_string(index) + " does not exist");
}

/**
 * @details append a new XLCfRule to m_conditionalFormattingNode, based on copyFrom
 */
size_t XLCfRules::create([[maybe_unused]] XLCfRule copyFrom, std::string cfRulePrefix)
{
    uint16_t maxPrio = maxPriorityValue();
    if (maxPrio == std::numeric_limits< uint16_t >::max()) {
        using namespace std::literals::string_literals;
        throw XLException("XLCfRules::"s + __func__
         + ": can not create a new cfRule entry: no available priority value - please renumberPriorities or otherwise free up the highest value"s);
    }

    size_t index = count();   // index for the cfRule to be created
    XMLNode newNode{};        // scope declaration

    // ===== Append new node prior to final whitespaces, if any
    if (index == 0) newNode = appendAndGetNode(*m_conditionalFormattingNode, "cfRule", m_nodeOrder);
    else {
        XMLNode lastCfRule = *cfRuleByIndex(index - 1).m_cfRuleNode;
        if (not lastCfRule.empty())
            newNode = m_conditionalFormattingNode->insert_child_after("cfRule", lastCfRule);
    }
    if (newNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLCfRules::"s + __func__ + ": failed to create a new cfRule entry");
    }

    // copyXMLNode(newNode, *copyFrom.m_cfRuleNode); // will use copyFrom as template, does nothing if copyFrom is empty
    m_conditionalFormattingNode->insert_child_before(pugi::node_pcdata, newNode).set_value(cfRulePrefix.c_str()); // prefix the new node with

    // regardless of whether a template was provided or not: set the newest rule to the lowest possible priority == highest value in use plus 1
    cfRuleByIndex(index).setPriority(maxPrio + 1);

    return index;
}

/**
 * @details Getter functions
 */
// TODO

/**
 * @details Setter functions
 */
// TODO / N/A


/**
 * @details assemble a string summary about the conditional formatting rules
 */
std::string XLCfRules::summary() const
{
    using namespace std::literals::string_literals;
    size_t rulesCount = count();
    if (rulesCount == 0)
        return "(no cfRule entries)";
    std::string result = "";
    for (size_t idx = 0; idx < rulesCount; ++idx) {
        using namespace std::literals::string_literals;
        result += "cfRule["s + std::to_string(idx) + "] "s + cfRuleByIndex(idx).summary();
        if (idx + 1 < rulesCount) result += ", ";
    }
    return result;
}


// ========== XLConditionalFormat Member Functions
/**
 * @details Constructor. Initializes an empty XLConditionalFormat object
 */
XLConditionalFormat::XLConditionalFormat() : m_conditionalFormattingNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLCellStyle object.
 */
XLConditionalFormat::XLConditionalFormat(const XMLNode& node) : m_conditionalFormattingNode(std::make_unique<XMLNode>(node)) {}

XLConditionalFormat::XLConditionalFormat(const XLConditionalFormat& other)
    : m_conditionalFormattingNode(std::make_unique<XMLNode>(*other.m_conditionalFormattingNode))
{}

XLConditionalFormat::~XLConditionalFormat() = default;

XLConditionalFormat& XLConditionalFormat::operator=(const XLConditionalFormat& other)
{
    if (&other != this) *m_conditionalFormattingNode = *other.m_conditionalFormattingNode;
    return *this;
}

/**
 * @details Returns the node empty status
 */
bool XLConditionalFormat::empty() const { return m_conditionalFormattingNode->empty(); }

/**
 * @details Getter functions
 */
std::string  XLConditionalFormat::sqref  () const { return m_conditionalFormattingNode->attribute("sqref").value(); }
XLCfRules    XLConditionalFormat::cfRules() const { return XLCfRules(*m_conditionalFormattingNode);                 }

/**
 * @details Setter functions
 */
bool XLConditionalFormat::setSqref(std::string newSqref) { return appendAndSetAttribute(*m_conditionalFormattingNode, "sqref", newSqref).empty() == false; }

/**
 * @brief Unsupported setter function
 */
bool XLConditionalFormat::setExtLst(XLUnsupportedElement const& newExtLst) { OpenXLSX::ignore(newExtLst); return false; }

/**
 * @details assemble a string summary about the conditional formatting
 */
std::string XLConditionalFormat::summary() const
{
    using namespace std::literals::string_literals;
    return "sqref is "s + sqref()
         + ", cfRules: "s + cfRules().summary();
}

// ========== XLConditionalFormats member functions, parent of XLConditionalFormat
/**
 * @details Constructor. Initializes an empty XLConditionalFormats object
 */
XLConditionalFormats::XLConditionalFormats() : m_sheetNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLConditionalFormats object.
 */
XLConditionalFormats::XLConditionalFormats(const XMLNode& sheet)
    : m_sheetNode(std::make_unique<XMLNode>(sheet))
{}

XLConditionalFormats::~XLConditionalFormats() {}

XLConditionalFormats::XLConditionalFormats(const XLConditionalFormats& other)
    : m_sheetNode(std::make_unique<XMLNode>(*other.m_sheetNode))
{}

XLConditionalFormats::XLConditionalFormats(XLConditionalFormats&& other)
    : m_sheetNode(std::move(other.m_sheetNode))
{}


/**
 * @details Copy assignment operator
 */
XLConditionalFormats& XLConditionalFormats::operator=(const XLConditionalFormats& other)
{
    if (&other != this) {
        *m_sheetNode = *other.m_sheetNode;
    }
    return *this;
}

/**
 * @details Returns the node empty status
 */
bool XLConditionalFormats::empty() const { return m_sheetNode->empty(); }

/**
 * @details Returns the amount of conditionalFormatting entries held by the class
 */
size_t XLConditionalFormats::count() const
{
    XMLNode node = m_sheetNode->first_child_of_type(pugi::node_element);
    while (not node.empty() && std::string(node.name()) != "conditionalFormatting")
        node = node.next_sibling_of_type(pugi::node_element);
    size_t count = 0;
    while (not node.empty() && std::string(node.name()) == "conditionalFormatting") {
        ++count;
        node = node.next_sibling_of_type(pugi::node_element);
    }
    return count;
}

/**
 * @details fetch XLConditionalFormat from m_sheetNode by index
 */
XLConditionalFormat XLConditionalFormats::conditionalFormatByIndex(size_t index) const
{
    XMLNode node = m_sheetNode->first_child_of_type(pugi::node_element);
    while (not node.empty() && std::string(node.name()) != "conditionalFormatting")
        node = node.next_sibling_of_type(pugi::node_element);

    if (not node.empty()) {
        size_t count = 0;
        while (not node.empty() && std::string(node.name()) == "conditionalFormatting" && count < index) {
            ++count;
            node = node.next_sibling_of_type(pugi::node_element);
        }
        if (count == index && std::string(node.name()) == "conditionalFormatting")
            return XLConditionalFormat(node);
    }
    using namespace std::literals::string_literals;
    throw XLException("XLConditionalFormats::"s + __func__ + ": conditional format with index "s + std::to_string(index) + " does not exist");
}

/**
 * @details append a new XLConditionalFormat to m_sheetNode, based on copyFrom
 */
size_t XLConditionalFormats::create([[maybe_unused]] XLConditionalFormat copyFrom, std::string conditionalFormattingPrefix)
{
    size_t index = count();   // index for the conditional formatting to be created
    XMLNode newNode{};        // scope declaration

    // ===== Append new node prior to final whitespaces, if any
    if (index == 0) newNode = appendAndGetNode(*m_sheetNode, "conditionalFormatting", m_nodeOrder);
    else {
        XMLNode lastConditionalFormat = *conditionalFormatByIndex(index - 1).m_conditionalFormattingNode;
        if (not lastConditionalFormat.empty())
            newNode = m_sheetNode->insert_child_after("conditionalFormatting", lastConditionalFormat);
    }
    if (newNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLException("XLConditionalFormats::"s + __func__ + ": failed to create a new conditional formatting entry");
	 }

    // copyXMLNode(newNode, *copyFrom.m_conditionalFormattingNode); // will use copyFrom as template, does nothing if copyFrom is empty
    m_sheetNode->insert_child_before(pugi::node_pcdata, newNode).set_value(conditionalFormattingPrefix.c_str());    // prefix the new node with conditionalFormattingPrefix

    return index;
}

/**
 * @details assemble a string summary about the conditional formattings
 */
std::string XLConditionalFormats::summary() const
{
    using namespace std::literals::string_literals;
    size_t conditionalFormatsCount = count();
    if (conditionalFormatsCount == 0)
        return "(no conditionalFormatting entries)";
    std::string result = "";
    for (size_t idx = 0; idx < conditionalFormatsCount; ++idx) {
        using namespace std::literals::string_literals;
        result += "conditionalFormatting["s + std::to_string(idx) + "] "s + conditionalFormatByIndex(idx).summary();
        if (idx + 1 < conditionalFormatsCount) result += ", ";
    }
    return result;
}
// ========== END <conditionalFormatting> related member function definitions


// ========== XLWorksheet Member Functions

/**
 * @details The constructor does some slight reconfiguration of the XML file, in order to make parsing easier.
 * For example, columns with identical formatting are by default grouped under the same node. However, this makes it more difficult to
 * parse, so the constructor reconfigures it so each column has it's own formatting.
 */
XLWorksheet::XLWorksheet(XLXmlData* xmlData) : XLSheetBase(xmlData)
{
    // ===== Read the dimensions of the Sheet and set data members accordingly.
    if (const std::string dimensions = xmlDocument().document_element().child("dimension").attribute("ref").value();
        dimensions.find(':') == std::string::npos)
        xmlDocument().document_element().child("dimension").set_value("A1");
    else
        xmlDocument().document_element().child("dimension").set_value(dimensions.substr(dimensions.find(':') + 1).c_str());

    // If Column properties are grouped, divide them into properties for individual Columns.
    if (xmlDocument().document_element().child("cols").type() != pugi::node_null) {
        auto currentNode = xmlDocument().document_element().child("cols").first_child_of_type(pugi::node_element);
        while (not currentNode.empty()) {
            uint16_t min {};
            uint16_t max {};
            try {
                min = std::stoi(currentNode.attribute("min").value());
                max = std::stoi(currentNode.attribute("max").value());
            }
            catch (...) {
                throw XLInternalError("Worksheet column min and/or max attributes are invalid.");
            }
            if (min != max) {
                currentNode.attribute("min").set_value(max);
                for (uint16_t i = min; i < max; i++) {    // NOLINT
                    auto newnode = xmlDocument().document_element().child("cols").insert_child_before("col", currentNode);
                    auto attr    = currentNode.first_attribute();
                    while (not attr.empty()) {    // NOLINT
                        newnode.append_attribute(attr.name()) = attr.value();
                        attr                                  = attr.next_attribute();
                    }
                    newnode.attribute("min") = i;
                    newnode.attribute("max") = i;
                }
            }
            currentNode = currentNode.next_sibling_of_type(pugi::node_element);
        }
    }
}

/**
 * @details copy-construct an XLWorksheet object from other
 */
XLWorksheet::XLWorksheet(const XLWorksheet& other) : XLSheetBase<XLWorksheet>(other)
{
    m_relationships = other.m_relationships;  // invoke XLRelationships copy assignment operator
    m_merges        = other.m_merges;         //  "     XLMergeCells       "
    m_vmlDrawing    = other.m_vmlDrawing;     //  "     XLVmlDrawing
    m_comments      = other.m_comments;       //  "     XLComments         "
    m_tables        = other.m_tables;         //  "     XLTables           "
}

/**
 * @details move-construct an XLWorksheet object from other
 */
XLWorksheet::XLWorksheet(XLWorksheet&& other) : XLSheetBase< XLWorksheet >(other)
{
    m_relationships = std::move(other.m_relationships);  // invoke XLRelationships move assignment operator
    m_merges        = std::move(other.m_merges);         //  "     XLMergeCells       "
    m_vmlDrawing    = std::move(other.m_vmlDrawing);     //  "     XLVmlDrawing
    m_comments      = std::move(other.m_comments);       //  "     XLComments         "
    m_tables        = std::move(other.m_tables);         //  "     XLTables           "
}

/**
 * @details copy-assign an XLWorksheet from other
 */
XLWorksheet& XLWorksheet::operator=(const XLWorksheet& other)
{
    XLSheetBase<XLWorksheet>::operator=(other); // invoke base class copy assignment operator
    m_relationships = other.m_relationships;
    m_merges        = other.m_merges;
    m_vmlDrawing    = other.m_vmlDrawing;
    m_comments      = other.m_comments;
    m_tables        = other.m_tables;
    return *this;
}

/**
 * @details move-assign an XLWorksheet from other
 */
XLWorksheet& XLWorksheet::operator=(XLWorksheet&& other)
{
    XLSheetBase<XLWorksheet>::operator=(other); // invoke base class move assignment operator
    m_relationships = std::move(other.m_relationships);
    m_merges        = std::move(other.m_merges);
    m_vmlDrawing    = std::move(other.m_vmlDrawing);
    m_comments      = std::move(other.m_comments);
    m_tables        = std::move(other.m_tables);
    return *this;
}


/**
 * @details
 */
XLColor XLWorksheet::getColor_impl() const
{
    return XLColor(xmlDocument().document_element().child("sheetPr").child("tabColor").attribute("rgb").value());
}

/**
 * @details
 */
void XLWorksheet::setColor_impl(const XLColor& color) { setTabColor(xmlDocument(), color); }

/**
 * @details
 */
bool XLWorksheet::isSelected_impl() const { return tabIsSelected(xmlDocument()); }

/**
 * @details
 */
void XLWorksheet::setSelected_impl(bool selected) { setTabSelected(xmlDocument(), selected); }

/**
 * @details
 */
bool XLWorksheet::isActive_impl() const
{
    return parentDoc().execQuery(XLQuery(XLQueryType::QuerySheetIsActive).setParam("sheetID", relationshipID())).result<bool>();
}

/**
 * @details
 */
bool XLWorksheet::setActive_impl()
{
    return parentDoc().execCommand(XLCommand(XLCommandType::SetSheetActive).setParam("sheetID", relationshipID()));
}

/**
 * @details
 */
XLCellAssignable XLWorksheet::cell(const std::string& ref) const { return cell(XLCellReference(ref)); }

/**
 * @details
 */
XLCellAssignable XLWorksheet::cell(const XLCellReference& ref) const { return cell(ref.row(), ref.column()); }

/**
 * @details This function returns a pointer to an XLCell object in the worksheet. This particular overload
 * also serves as the main function, called by the other overloads.
 */
XLCellAssignable XLWorksheet::cell(uint32_t rowNumber, uint16_t columnNumber) const
{
    const XMLNode rowNode  = getRowNode(xmlDocument().document_element().child("sheetData"), rowNumber);
    const XMLNode cellNode = getCellNode(rowNode, columnNumber, rowNumber);
    // ===== Move-construct XLCellAssignable from temporary XLCell
    return XLCellAssignable(XLCell(cellNode, parentDoc().sharedStrings()));
}

/**
 * @details
 */
XLCellAssignable XLWorksheet::findCell(const std::string& ref) const { return findCell(XLCellReference(ref)); }

/**
 * @details
 */
XLCellAssignable XLWorksheet::findCell(const XLCellReference& ref) const { return findCell(ref.row(), ref.column()); }

/**
 * @details This function attempts to find a cell, but creates neither the row nor the cell XML if missing - and returns an empty XLCellAssignable instead
 */
XLCellAssignable XLWorksheet::findCell(uint32_t rowNumber, uint16_t columnNumber) const
{
    return XLCellAssignable(XLCell(findCellNode(findRowNode( xmlDocument().document_element().child("sheetData"), rowNumber ), columnNumber), parentDoc().sharedStrings()));
}

/**
 * @details
 */
XLCellRange XLWorksheet::range() const { return range(XLCellReference("A1"), lastCell()); }

/**
 * @details
 */
XLCellRange XLWorksheet::range(const XLCellReference& topLeft, const XLCellReference& bottomRight) const
{
    return XLCellRange(xmlDocument().document_element().child("sheetData"),
                       topLeft,
                       bottomRight,
                       parentDoc().sharedStrings());
}

/**
 * @details Get a range based on two cell reference strings
 */
XLCellRange XLWorksheet::range(std::string const& topLeft, std::string const& bottomRight) const
{
    return range(XLCellReference(topLeft), XLCellReference(bottomRight));
}

/**
 * @details Get a range based on a reference string
 */
XLCellRange XLWorksheet::range(std::string const& rangeReference) const
{
    size_t pos = rangeReference.find_first_of(':');
    return range(rangeReference.substr(0, pos), rangeReference.substr(pos + 1, std::string::npos));
}

/**
 * @details
 * @pre
 * @post
 */
XLRowRange XLWorksheet::rows() const    // 2024-04-29: patched for whitespace
{
    const auto sheetDataNode = xmlDocument().document_element().child("sheetData");
    return XLRowRange(sheetDataNode,
                      1,
                      (sheetDataNode.last_child_of_type(pugi::node_element).empty()
                           ? 1
                           : static_cast<uint32_t>(sheetDataNode.last_child_of_type(pugi::node_element).attribute("r").as_ullong())),
                      parentDoc().sharedStrings());
}

/**
 * @details
 * @pre
 * @post
 */
XLRowRange XLWorksheet::rows(uint32_t rowCount) const
{
    return XLRowRange(xmlDocument().document_element().child("sheetData"),
                      1,
                      rowCount,
                      parentDoc().sharedStrings());
}

/**
 * @details
 * @pre
 * @post
 */
XLRowRange XLWorksheet::rows(uint32_t firstRow, uint32_t lastRow) const
{
    return XLRowRange(xmlDocument().document_element().child("sheetData"),
                      firstRow,
                      lastRow,
                      parentDoc().sharedStrings());
}

/**
 * @details Get the XLRow object corresponding to the given row number. In the XML file, all cell data are stored under
 * the corresponding row, and all rows have to be ordered in ascending order. If a row have no data, there may not be a
 * node for that row.
 */
XLRow XLWorksheet::row(uint32_t rowNumber) const
{
    return XLRow { getRowNode(xmlDocument().document_element().child("sheetData"), rowNumber),
                   parentDoc().sharedStrings() };
}

/**
 * @details Get the XLColumn object corresponding to the given column number. In the underlying XML data structure,
 * column nodes do not hold any cell data. Columns are used solely to hold data regarding column formatting.
 * @todo Consider simplifying this function. Can any standard algorithms be used?
 */
XLColumn XLWorksheet::column(uint16_t columnNumber) const
{
    using namespace std::literals::string_literals;

    if (columnNumber < 1 || columnNumber > OpenXLSX::MAX_COLS)    // 2024-08-05: added range check
        throw XLException("XLWorksheet::column: columnNumber "s + std::to_string(columnNumber) + " is outside allowed range [1;"s + std::to_string(MAX_COLS) + "]"s);

    // If no columns exists, create the <cols> node in the XML document.
    if (xmlDocument().document_element().child("cols").empty())
        xmlDocument().document_element().insert_child_before("cols", xmlDocument().document_element().child("sheetData"));

    // ===== Find the column node, if it exists
    auto columnNode = xmlDocument().document_element().child("cols").find_child([&](const XMLNode node) {
        return (columnNumber >= node.attribute("min").as_int() && columnNumber <= node.attribute("max").as_int()) ||
               node.attribute("min").as_int() > columnNumber;
    });

    uint16_t minColumn {};
    uint16_t maxColumn {};
    if (not columnNode.empty()) {
        minColumn = columnNode.attribute("min").as_int();    // only look it up once for multiple access
        maxColumn = columnNode.attribute("max").as_int();    //   "
    }

    // ===== If the node exists for the column, and only spans that column, then continue...
    if (not columnNode.empty() && (minColumn == columnNumber) && (maxColumn == columnNumber)) {
    }

    // ===== If the node exists for the column, but spans several columns, split it into individual nodes, and set columnNode to the right
    // one...
    // BUGFIX 2024-04-27 - old if condition would split a multi-column setting even if columnNumber is < minColumn (see lambda return value
    // above)
    // NOTE 2024-04-27: the column splitting for loaded files is already handled in the constructor, technically this code is not necessary
    // here
    else if (not columnNode.empty() && (columnNumber >= minColumn) && (minColumn != maxColumn)) {
        // ===== Split the node in individual columns...
        columnNode.attribute("min").set_value(maxColumn);    // Limit the original node to a single column
        for (int i = minColumn; i < maxColumn; ++i) {
            auto node = xmlDocument().document_element().child("cols").insert_copy_before(columnNode, columnNode);
            node.attribute("min").set_value(i);
            node.attribute("max").set_value(i);
        }
        // BUGFIX 2024-04-27: the original node should not be deleted, but - in line with XLWorksheet constructor - is now limited above to
        // min = max attribute
        // // ===== Delete the original node
        // columnNode = columnNode.previous_sibling_of_type(pugi::node_element); // due to insert loop, previous node should be guaranteed
        // // to be an element node
        // xmlDocument().document_element().child("cols").remove_child(columnNode.next_sibling());

        // ===== Find the node corresponding to the column number - BUGFIX 2024-04-27: loop should abort on empty node
        while (not columnNode.empty() && columnNode.attribute("min").as_int() != columnNumber)
            columnNode = columnNode.previous_sibling_of_type(pugi::node_element);
        if (columnNode.empty())
            throw XLInternalError("XLWorksheet::"s + __func__ + ": column node for index "s + std::to_string(columnNumber) +
                                  "not found after splitting column nodes"s);
    }

    // ===== If a node for the column does NOT exist, but a node for a higher column exists...
    else if (not columnNode.empty() && minColumn > columnNumber) {
        columnNode                                 = xmlDocument().document_element().child("cols").insert_child_before("col", columnNode);
        columnNode.append_attribute("min")         = columnNumber;
        columnNode.append_attribute("max")         = columnNumber;
        columnNode.append_attribute("width")       = 9.8;    // NOLINT
        columnNode.append_attribute("customWidth") = 0;
    }

    // ===== Otherwise, the end of the list is reached, and a new node is appended
    else if (columnNode.empty()) {
        columnNode                                 = xmlDocument().document_element().child("cols").append_child("col");
        columnNode.append_attribute("min")         = columnNumber;
        columnNode.append_attribute("max")         = columnNumber;
        columnNode.append_attribute("width")       = 9.8;    // NOLINT
        columnNode.append_attribute("customWidth") = 0;
    }

    if (columnNode.empty()) {
        using namespace std::literals::string_literals;
        throw XLInternalError("XLWorksheet::"s + __func__ + ": was unable to find or create node for column "s +
                              std::to_string(columnNumber));
    }
    return XLColumn(columnNode);
}

/**
 * @details Get the column with the given column reference by converting the columnRef into a column number
 *          and forwarding the implementation to XLColumn XLWorksheet::column(uint16_t columnNumber) const
 */
XLColumn XLWorksheet::column(std::string const& columnRef) const { return column(XLCellReference::columnAsNumber(columnRef)); }

/**
 * @details Returns an XLCellReference to the last cell using rowCount() and columnCount() methods.
 */
XLCellReference XLWorksheet::lastCell() const noexcept { return { rowCount(), columnCount() }; }

/**
 * @details Iterates through the rows and finds the maximum number of cells.
 */
uint16_t XLWorksheet::columnCount() const noexcept
{
    uint16_t maxCount = 0;    // Pull request: Update XLSheet.cpp with correct type #176, Explicitely cast to unsigned short int #163
    XLRowRange rowsRange = rows();
    for (XLRowIterator rowIt = rowsRange.begin(); rowIt != rowsRange.end(); ++rowIt ) {
        if (rowIt.rowExists()) {
            uint16_t cellCount = rowIt->cellCount();
            maxCount = std::max(cellCount, maxCount);
        }
    }
    return maxCount;
}

/**
 * @details Finds the last row (node) and returns the row number.
 */
uint32_t XLWorksheet::rowCount() const noexcept
{
    return static_cast<uint32_t>(
        xmlDocument().document_element().child("sheetData").last_child_of_type(pugi::node_element).attribute("r").as_ullong());
}

/**
 * @details finds a given row and deletes it
 */
bool XLWorksheet::deleteRow(uint32_t rowNumber)
{
    XMLNode row = xmlDocument().document_element().child("sheetData").first_child_of_type(pugi::node_element);
    XMLNode lastRow = xmlDocument().document_element().child("sheetData").last_child_of_type(pugi::node_element);

    // ===== Fail if rowNumber is not in XML
    if (rowNumber < row.attribute("r").as_ullong() || rowNumber > lastRow.attribute("r").as_ullong()) return false;

    // ===== If rowNumber is closer to first (existing) row than to last row, search forwards
    if (rowNumber - row.attribute("r").as_ullong() < lastRow.attribute("r").as_ullong() - rowNumber)
        while (not row.empty() && (row.attribute("r").as_ullong() < rowNumber)) row = row.next_sibling_of_type(pugi::node_element);

    // ===== Otherwise, search backwards
    else {
        row = lastRow;
        while (not row.empty() && (row.attribute("r").as_ullong() > rowNumber)) row = row.previous_sibling_of_type(pugi::node_element);
    }

    if (row.attribute("r").as_ullong() != rowNumber) return false;    // row not found in XML

    // ===== If row was located: remove it
    return xmlDocument().document_element().child("sheetData").remove_child(row);
}

/**
 * @details
 */
void XLWorksheet::updateSheetName(const std::string& oldName, const std::string& newName)    // 2024-04-29: patched for whitespace
{
    // ===== Set up temporary variables
    std::string oldNameTemp = oldName;
    std::string newNameTemp = newName;
    std::string formula;

    // ===== If the sheet name contains spaces, it should be enclosed in single quotes (')
    if (oldName.find(' ') != std::string::npos) oldNameTemp = "\'" + oldName + "\'";
    if (newName.find(' ') != std::string::npos) newNameTemp = "\'" + newName + "\'";

    // ===== Ensure only sheet names are replaced (references to sheets always ends with a '!')
    oldNameTemp += '!';
    newNameTemp += '!';

    // ===== Iterate through all defined names
    XMLNode row = xmlDocument().document_element().child("sheetData").first_child_of_type(pugi::node_element);
    for (; not row.empty(); row = row.next_sibling_of_type(pugi::node_element)) {
        for (XMLNode cell = row.first_child_of_type(pugi::node_element);
             not cell.empty();
             cell         = cell.next_sibling_of_type(pugi::node_element))
        {
            if (!XLCell(cell, parentDoc().sharedStrings()).hasFormula()) continue;

            formula = XLCell(cell, parentDoc().sharedStrings()).formula().get();

            // ===== Skip if formula contains a '[' and ']' (means that the defined refers to external workbook)
            if (formula.find('[') == std::string::npos && formula.find(']') == std::string::npos) {
                // ===== For all instances of the old sheet name in the formula, replace with the new name.
                while (formula.find(oldNameTemp) != std::string::npos) {    // NOLINT
                    formula.replace(formula.find(oldNameTemp), oldNameTemp.length(), newNameTemp);
                }
                XLCell(cell, parentDoc().sharedStrings()).formula() = formula;
            }
        }
    }
}

/**
 * @details upon first access, ensure that the worksheet's <mergeCells> tag exists, and create an XLMergeCells object
 */
XLMergeCells & XLWorksheet::merges()
{
    if (!m_merges.valid())
        m_merges = XLMergeCells(xmlDocument().document_element(), m_nodeOrder);
    return m_merges;
}

/**
 * @details create a new mergeCell element in the worksheet mergeCells array, with the only
 *          attribute being ref="A1:D6" (with the top left and bottom right cells of the range)
 *          The mergeCell element otherwise remains empty (no value)
 *          The mergeCells attribute count will be increased by one
 *          If emptyHiddenCells is true (XLEmptyHiddenCells), the values of all but the top left cell of the range
 *          will be deleted (not from shared strings)
 *          If any cell within the range is the sheet's active cell, the active cell will be set to the
 *          top left corner of rangeToMerge
 */
void XLWorksheet::mergeCells(XLCellRange const& rangeToMerge, bool emptyHiddenCells)
{
    if (rangeToMerge.numRows() * rangeToMerge.numColumns() < 2) {
        using namespace std::literals::string_literals;
        throw XLInputError("XLWorksheet::"s + __func__ + ": rangeToMerge must comprise at least 2 cells"s);
    }

    merges().appendMerge(rangeToMerge.address());
    if (emptyHiddenCells) {
        // ===== Iterate over rangeToMerge, delete values & attributes (except r and s) for all but the first cell in the range
        XLCellIterator it = rangeToMerge.begin();
        ++it; // leave first cell untouched
        while (it != rangeToMerge.end()) {
            // ===== When merging with emptyHiddenCells in LO, subsequent unmerging restores the cell styles -> so keep them here as well
            it->clear(XLKeepCellStyle);    // clear cell contents except style
            ++it;
        }
    }
}

/**
 * @details Convenience wrapper for the previous function, using a std::string range reference
 */
void XLWorksheet::mergeCells(const std::string& rangeReference, bool emptyHiddenCells) {
    mergeCells(range(rangeReference), emptyHiddenCells);
}

/**
 * @details check if rangeToUnmerge exists in mergeCells array & remove it
 */
void XLWorksheet::unmergeCells(XLCellRange const& rangeToUnmerge)
{
    int32_t mergeIndex = merges().findMerge(rangeToUnmerge.address());
    if (mergeIndex != -1)
        merges().deleteMerge(mergeIndex);
    else {
        using namespace std::literals::string_literals;
        throw XLInputError("XLWorksheet::"s + __func__ + ": merged range "s + rangeToUnmerge.address() + " does not exist"s);
    }
}

/**
 * @details Convenience wrapper for the previous function, using a std::string range reference
 */
void XLWorksheet::unmergeCells(const std::string& rangeReference)
{
    unmergeCells(range(rangeReference));
}

/**
 * @details Retrieve the column's format
 */
XLStyleIndex XLWorksheet::getColumnFormat(uint16_t columnNumber) const { return column(columnNumber).format(); }
XLStyleIndex XLWorksheet::getColumnFormat(const std::string& columnNumber) const { return getColumnFormat(XLCellReference::columnAsNumber(columnNumber)); }

/**
 * @details Set the style for the identified column,
 *          then iterate over all rows to set the same style for existing cells in that column
 */
bool XLWorksheet::setColumnFormat(uint16_t columnNumber, XLStyleIndex cellFormatIndex)
{
    if (!column(columnNumber).setFormat(cellFormatIndex)) // attempt to set column format
        return false; // early fail

    XLRowRange allRows = rows();
    for (XLRowIterator rowIt = allRows.begin(); rowIt != allRows.end(); ++rowIt) {
        XLCell curCell = rowIt->findCell(columnNumber);
        if (curCell && !curCell.setCellFormat(cellFormatIndex))    // attempt to set cell format for a non empty cell
            return false; // failure to set cell format
    }
    return true; // if loop finished nominally: success
}
bool XLWorksheet::setColumnFormat(const std::string& columnNumber, XLStyleIndex cellFormatIndex) { return setColumnFormat(XLCellReference::columnAsNumber(columnNumber), cellFormatIndex); }

/**
 * @details Retrieve the row's format
 */
XLStyleIndex XLWorksheet::getRowFormat(uint16_t rowNumber) const { return row(rowNumber).format(); }

/**
 * @details Set the style for the identified row,
 *          then iterate over all existing cells in the row to set the same style 
 */
bool XLWorksheet::setRowFormat(uint32_t rowNumber, XLStyleIndex cellFormatIndex)
{
    if (!row(rowNumber).setFormat(cellFormatIndex)) // attempt to set row format
        return false; // early fail

    XLRowDataRange rowCells = row(rowNumber).cells();
    for (XLRowDataIterator cellIt = rowCells.begin(); cellIt != rowCells.end(); ++cellIt)
        if (!cellIt->setCellFormat(cellFormatIndex)) // attempt to set cell format
            return false;
    return true; // if loop finished nominally: success
}

/**
 * @details Provide access to worksheet conditional formats
 */
XLConditionalFormats XLWorksheet::conditionalFormats() const { return XLConditionalFormats(xmlDocument().document_element()); }

/**
 * @brief Set the <sheetProtection> attributes sheet, objects and scenarios respectively
 */
bool XLWorksheet::protectSheet(bool set) {
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "sheet",     (set ? "true" : "false"),
    /**/                             XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::protectObjects(bool set) {
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "objects",   (set ? "true" : "false"),
    /**/                             XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::protectScenarios(bool set) {
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "scenarios", (set ? "true" : "false"),
    /**/                             XLKeepAttributes, m_nodeOrder).empty() == false;
}


/**
 * @brief Set the <sheetProtection> attributes insertColumns, insertRows, deleteColumns, deleteRows, selectLockedCells, selectUnlockedCells
 */
bool XLWorksheet::allowInsertColumns(bool set) {
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "insertColumns",       (!set ? "true" : "false"), // invert allow set(ting) for the protect setting
    /**/                             XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowInsertRows(bool set) {
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "insertRows",          (!set ? "true" : "false"), // invert allow set(ting) for the protect setting
    /**/                             XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowDeleteColumns(bool set) {
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "deleteColumns",       (!set ? "true" : "false"), // invert allow set(ting) for the protect setting
    /**/                             XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowDeleteRows(bool set) {
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "deleteRows",          (!set ? "true" : "false"), // invert allow set(ting) for the protect setting
    /**/                             XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowSelectLockedCells(bool set) {
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "selectLockedCells",   (!set ? "true" : "false"), // invert allow set(ting) for the protect setting
    /**/                             XLKeepAttributes, m_nodeOrder).empty() == false;
}
bool XLWorksheet::allowSelectUnlockedCells(bool set) {
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "selectUnlockedCells", (!set ? "true" : "false"), // invert allow set(ting) for the protect setting
    /**/                             XLKeepAttributes, m_nodeOrder).empty() == false;
}

/**
 * @details store the password hash to the underlying XML node
 */
bool XLWorksheet::setPasswordHash(std::string hash)
{
    XMLNode sheetNode = xmlDocument().document_element();
    return appendAndSetNodeAttribute(sheetNode, "sheetProtection", "password", hash,
    /**/                             XLKeepAttributes, m_nodeOrder).empty() == false;
}

/**
 * @details calculate the password hash and store it to the underlying XML node
 */
bool XLWorksheet::setPassword(std::string password) {
    return setPasswordHash(password.length() ? ExcelPasswordHashAsString(password) : std::string(""));
}
/**
 * @details
 */
bool XLWorksheet::clearPassword() { return setPasswordHash(""); }

/**
 * @details
 */
bool XLWorksheet::clearSheetProtection() { return xmlDocument().document_element().remove_child("sheetProtection"); }

/**
 * @details
*/
bool XLWorksheet::sheetProtected()     const { return xmlDocument().document_element().child("sheetProtection").attribute("sheet").as_bool(false);     }
bool XLWorksheet::objectsProtected()   const { return xmlDocument().document_element().child("sheetProtection").attribute("objects").as_bool(false);   }
bool XLWorksheet::scenariosProtected() const { return xmlDocument().document_element().child("sheetProtection").attribute("scenarios").as_bool(false); }

/**
 * @details the following settings are inverted to get from "protected" meaning in OOXML to "allowed" meaning in the OpenXLSX API
 * @details insertColumns, insertRows, deleteColumns, deleteRows are protected by default
 * @details selectLockedCells, selectUnlockedCells are allowed by default
*/
bool XLWorksheet::insertColumnsAllowed()       const { return not xmlDocument().document_element().child("sheetProtection").attribute("insertColumns").as_bool(true);        }
bool XLWorksheet::insertRowsAllowed()          const { return not xmlDocument().document_element().child("sheetProtection").attribute("insertRows").as_bool(true);           }
bool XLWorksheet::deleteColumnsAllowed()       const { return not xmlDocument().document_element().child("sheetProtection").attribute("deleteColumns").as_bool(true);        }
bool XLWorksheet::deleteRowsAllowed()          const { return not xmlDocument().document_element().child("sheetProtection").attribute("deleteRows").as_bool(true);           }
bool XLWorksheet::selectLockedCellsAllowed()   const { return not xmlDocument().document_element().child("sheetProtection").attribute("selectLockedCells").as_bool(false);   }
bool XLWorksheet::selectUnlockedCellsAllowed() const { return not xmlDocument().document_element().child("sheetProtection").attribute("selectUnlockedCells").as_bool(false); }
/**
 * @details
*/
std::string XLWorksheet::passwordHash()        const { return     xmlDocument().document_element().child("sheetProtection").attribute("password").value();                   }

/**
 * @details
 */
bool XLWorksheet::passwordIsSet()              const { return passwordHash().length() > 0; }

/**
 * @details
 */
std::string XLWorksheet::sheetProtectionSummary() const
{
    using namespace std::literals::string_literals;
    return   "sheet is"s               + ( sheetProtected()             ? ""s : " not"s ) + " protected"s
         + ", objects are"s            + ( objectsProtected()           ? ""s : " not"s ) + " protected"s
         + ", scenarios are"s          + ( scenariosProtected()         ? ""s : " not"s ) + " protected"s
         + ", insertColumns is"s       + ( insertColumnsAllowed()       ? ""s : " not"s ) + " allowed"s
         + ", insertRows is"s          + ( insertRowsAllowed()          ? ""s : " not"s ) + " allowed"s
         + ", deleteColumns is"s       + ( deleteColumnsAllowed()       ? ""s : " not"s ) + " allowed"s
         + ", deleteRows is"s          + ( deleteRowsAllowed()          ? ""s : " not"s ) + " allowed"s
         + ", selectLockedCells is"s   + ( selectLockedCellsAllowed()   ? ""s : " not"s ) + " allowed"s
         + ", selectUnlockedCells is"s + ( selectUnlockedCellsAllowed() ? ""s : " not"s ) + " allowed"s
         + ", password is"s + ( passwordIsSet() ? ""s : " not"s ) + " set"s
         + ( passwordIsSet() ? ( ", passwordHash is "s + passwordHash() ) : ""s )
    ;
}

/**
 * @details functions to test whether a VMLDrawing / Comments / Tables entry exists for this worksheet (without creating those entries)
 */
bool XLWorksheet::hasRelationships() const { return parentDoc().hasSheetRelationships(sheetXmlNumber()); }
bool XLWorksheet::hasVmlDrawing()    const { return parentDoc().hasSheetVmlDrawing(   sheetXmlNumber()); }
bool XLWorksheet::hasComments()      const { return parentDoc().hasSheetComments(     sheetXmlNumber()); }
bool XLWorksheet::hasTables()        const { return parentDoc().hasSheetTables(       sheetXmlNumber()); }

/**
 * @details fetches XLVmlDrawing for the sheet - creates & assigns the class if empty
 */
XLVmlDrawing& XLWorksheet::vmlDrawing()
{
    if (!m_vmlDrawing.valid()) {
        // ===== Append xdr namespace attribute to worksheet if not present
        XMLNode docElement = xmlDocument().document_element();
        XMLAttribute xdrNamespace = appendAndGetAttribute(docElement, "xmlns:xdr", "");
        xdrNamespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

        std::ignore = relationships(); // create sheet relationships if not existing

        // ===== Trigger parentDoc to create drawing XML file and return it
        uint16_t sheetXmlNo = sheetXmlNumber();
        m_vmlDrawing = parentDoc().sheetVmlDrawing(sheetXmlNo); // fetch drawing for this worksheet
        if (!m_vmlDrawing.valid())
            throw XLException("XLWorksheet::vmlDrawing(): could not create drawing XML");
        std::string drawingRelativePath = getPathARelativeToPathB( m_vmlDrawing.getXmlPath(), getXmlPath() );
        XLRelationshipItem vmlDrawingRelationship;
        if (!m_relationships.targetExists(drawingRelativePath))
            vmlDrawingRelationship = m_relationships.addRelationship(XLRelationshipType::VMLDrawing, drawingRelativePath);
        else
            vmlDrawingRelationship = m_relationships.relationshipByTarget(drawingRelativePath);
        if (vmlDrawingRelationship.empty())
            throw XLException("XLWorksheet::vmlDrawing(): could not add determine sheet relationship for VML Drawing");
        XMLNode legacyDrawing = appendAndGetNode(docElement, "legacyDrawing", m_nodeOrder);
        if (legacyDrawing.empty())
            throw XLException("XLWorksheet::vmlDrawing(): could not add <legacyDrawing> element to worksheet XML");
        appendAndSetAttribute(legacyDrawing, "r:id", vmlDrawingRelationship.id());
    }

    return m_vmlDrawing;
}

/**
 * @details fetches XLComments for the sheet - creates & assigns the class if empty
 */
XLComments& XLWorksheet::comments()
{
    if (!m_comments.valid()) {
        std::ignore = relationships(); // create sheet relationships if not existing
        // ===== Unfortunately, xl/vmlDrawing#.vml is needed to support comments: append xdr namespace attribute to worksheet if not present
        std::ignore = vmlDrawing();    // create sheet VML drawing if not existing


        // ===== Trigger parentDoc to create comment XML file and return it
        uint16_t sheetXmlNo = sheetXmlNumber();
        // std::cout << "worksheet comments for sheetId " << sheetXmlNo << std::endl;
        m_comments = parentDoc().sheetComments(sheetXmlNo); // fetch comments for this worksheet
        if (!m_comments.valid())
            throw XLException("XLWorksheet::comments(): could not create comments XML");
        m_comments.setVmlDrawing(m_vmlDrawing); // link XLVmlDrawing object to the comments so it can be modified from there
        std::string commentsRelativePath = getPathARelativeToPathB( m_comments.getXmlPath(), getXmlPath() );
        if (!m_relationships.targetExists(commentsRelativePath))
            m_relationships.addRelationship(XLRelationshipType::Comments, commentsRelativePath);
    }

    return m_comments;
}

/**
 * @details fetches XLTables for the sheet - creates & assigns the class if empty
 */
XLTables& XLWorksheet::tables()
{
    if (!m_tables.valid()) {
        std::ignore = relationships(); // create sheet relationships if not existing


        // ===== Trigger parentDoc to create tables XML file and return it
        uint16_t sheetXmlNo = sheetXmlNumber();
        // std::cout << "worksheet tables for sheetId " << sheetXmlNo << std::endl;
        m_tables = parentDoc().sheetTables(sheetXmlNo); // fetch tables for this worksheet
        if (!m_tables.valid())
            throw XLException("XLWorksheet::tables(): could not create tables XML");
        std::string tablesRelativePath = getPathARelativeToPathB( m_tables.getXmlPath(), getXmlPath() );
        if (!m_relationships.targetExists(tablesRelativePath))
            m_relationships.addRelationship(XLRelationshipType::Table, tablesRelativePath);
    }

    return m_tables;
}

/**
 * @details perform a pattern matching on getXmlPath for (regex) .*xl/worksheets/sheet([0-9]*)\.xml$ and extract the numeric part \1
 */
uint16_t XLWorksheet::sheetXmlNumber() const
{
   constexpr const char *searchPattern = "xl/worksheets/sheet";
   std::string xmlPath = getXmlPath();
   size_t pos = xmlPath.find(searchPattern);                      // ensure compatibility with expected path
   if (pos == std::string::npos) return 0;
   pos += strlen(searchPattern);
   size_t pos2 = pos;
   while (std::isdigit(xmlPath[pos2])) ++pos2;                    // find sheet number in xmlPath - aborts on end of string
   if (pos2 == pos || xmlPath.substr(pos2) != ".xml" ) return 0;  // ensure compatibility with expected path

   // success: convert the sheet number part of xmlPath to uint16_t
   return static_cast<uint16_t>(std::stoi(xmlPath.substr(pos, pos2 - pos)));
}

/**
 * @details fetches XLRelationships for the sheet - creates & assigns the class if empty
 */
XLRelationships& XLWorksheet::relationships()
{
    if (!m_relationships.valid()){
        // trigger parentDoc to create relationships XML file and relationship and return it
        m_relationships = parentDoc().sheetRelationships(sheetXmlNumber()); // fetch relationships for this worksheet
    }
    if (!m_relationships.valid())
        throw XLException("XLWorksheet::relationships(): could not create relationships XML");
    return m_relationships;
}

/**
 * @details Constructor
 */
XLChartsheet::XLChartsheet(XLXmlData* xmlData) : XLSheetBase(xmlData) {}

/**
 * @details Destructor. Default implementation used.
 */
XLChartsheet::~XLChartsheet() = default;

/**
 * @details
 */
XLColor XLChartsheet::getColor_impl() const
{
    return XLColor(xmlDocument().document_element().child("sheetPr").child("tabColor").attribute("rgb").value());
}

/**
 * @details Calls the setTabColor() free function.
 */
void XLChartsheet::setColor_impl(const XLColor& color) { setTabColor(xmlDocument(), color); }

/**
 * @details Calls the tabIsSelected() free function.
 */
bool XLChartsheet::isSelected_impl() const { return tabIsSelected(xmlDocument()); }

/**
 * @details Calls the setTabSelected() free function.
 */
void XLChartsheet::setSelected_impl(bool selected) { setTabSelected(xmlDocument(), selected); }

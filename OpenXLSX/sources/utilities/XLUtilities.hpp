//
// Created by Kenneth Balslev on 24/08/2020.
//

#ifndef OPENXLSX_XLUTILITIES_HPP
#define OPENXLSX_XLUTILITIES_HPP

#include <fstream>
#include <pugixml.hpp>
#include <string>       // 2024-04-25 needed for xml_node_type_string
#include <string_view>  // std::string_view
#include <vector>       // std::vector< std::string_view >

#include "XLConstants.hpp"        // 2024-05-28 OpenXLSX::MAX_ROWS
#include "XLCellReference.hpp"
#include "XLCellValue.hpp"        // OpenXLSX::XLValueType
#include "XLContentTypes.hpp"     // OpenXLSX::XLContentType
#include "XLRelationships.hpp"    // OpenXLSX::XLRelationshipType
#include "XLStyles.hpp"           // OpenXLSX::XLStyleIndex
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    constexpr const bool XLRemoveAttributes = true;   // helper variables for appendAndSetNodeAttribute, parameter removeAttributes
    constexpr const bool XLKeepAttributes   = false;  //

    /**
     * @brief Get rid of compiler warnings about unused variables (-Wunused-variable) or unused parameters (-Wunusued-parameter)
     * @param (unnamed) the variable / parameter which is intentionally unused (e.g. for function stubs)
     */
    template<class T> void ignore( const T& ) {}

    /**
     * @brief Get a string representation of pugi::xml_node_type
     * @param t the pugi::xml_node_type of a node
     * @return a std::string containing the descriptive name of the node type
     */
    inline std::string XLValueTypeString( XLValueType t )
    {
        using namespace std::literals::string_literals;
        switch( t ) {
            case XLValueType::Empty      : return "Empty"s;
            case XLValueType::Boolean    : return "Boolean"s;
            case XLValueType::Integer    : return "Integer"s;
            case XLValueType::Float      : return "Float"s;
            case XLValueType::Error      : return "Error"s;
            case XLValueType::String     : return "String"s;
        }
        throw XLInternalError(__func__ + "Invalid XLValueType."s);
    }

    /**
     * @brief Get a string representation of pugi::xml_node_type
     * @param t the pugi::xml_node_type of a node
     * @return a std::string containing the descriptive name of the node type
     */
    inline std::string xml_node_type_string( pugi::xml_node_type t )
    {
        using namespace std::literals::string_literals;
        switch( t ) {
            case pugi::node_null        : return "node_null"s;
            case pugi::node_document    : return "node_document"s;
            case pugi::node_element     : return "node_element"s;
            case pugi::node_pcdata      : return "node_pcdata"s;
            case pugi::node_cdata       : return "node_cdata"s;
            case pugi::node_comment     : return "node_comment"s;
            case pugi::node_pi          : return "node_pi"s;
            case pugi::node_declaration : return "node_declaration"s;
            case pugi::node_doctype     : return "node_doctype"s;
        }
        throw XLInternalError("Invalid XML node type.");
    }

    /**
     * @brief Get a string representation of OpenXLSX::XLContentType
     * @param t an OpenXLSX::XLContentType value
     * @return a std::string containing the descriptive name of the content type
     */
    inline std::string XLContentTypeString( OpenXLSX::XLContentType const & t )
    {
        using namespace OpenXLSX;
        switch (t) {
            case XLContentType::Workbook: return "Workbook";
            case XLContentType::Relationships: return "Relationships";
            case XLContentType::WorkbookMacroEnabled: return "WorkbookMacroEnabled";
            case XLContentType::Worksheet: return "Worksheet";
            case XLContentType::Chartsheet: return "Chartsheet";
            case XLContentType::ExternalLink: return "ExternalLink";
            case XLContentType::Theme: return "Theme";
            case XLContentType::Styles: return "Styles";
            case XLContentType::SharedStrings: return "SharedStrings";
            case XLContentType::Drawing: return "Drawing";
            case XLContentType::Chart: return "Chart";
            case XLContentType::ChartStyle: return "ChartStyle";
            case XLContentType::ChartColorStyle: return "ChartColorStyle";
            case XLContentType::ControlProperties: return "ControlProperties";
            case XLContentType::CalculationChain: return "CalculationChain";
            case XLContentType::VBAProject: return "VBAProject";
            case XLContentType::CoreProperties: return "CoreProperties";
            case XLContentType::ExtendedProperties: return "ExtendedProperties";
            case XLContentType::CustomProperties: return "CustomProperties";
            case XLContentType::VMLDrawing: return "VMLDrawing";
            case XLContentType::Comments: return "Comments";
            case XLContentType::Table: return "Table";
            case XLContentType::Unknown: return "Unknown";
        }
        return "invalid";
    }

    /**
     * @brief Get a string representation of OpenXLSX::XLRelationshipType
     * @param t an OpenXLSX::XLRelationshipType value
     * @return std::string containing the descriptive name of the relationship type
     */
    inline std::string XLRelationshipTypeString( OpenXLSX::XLRelationshipType const & t )
    {
        using namespace OpenXLSX;
        switch (t) {
            case XLRelationshipType::CoreProperties: return "CoreProperties";
            case XLRelationshipType::ExtendedProperties: return "ExtendedProperties";
            case XLRelationshipType::CustomProperties: return "CustomProperties";
            case XLRelationshipType::Workbook: return "Workbook";
            case XLRelationshipType::Worksheet: return "Worksheet";
            case XLRelationshipType::Chartsheet: return "Chartsheet";
            case XLRelationshipType::Dialogsheet: return "Dialogsheet";
            case XLRelationshipType::Macrosheet: return "Macrosheet";
            case XLRelationshipType::CalculationChain: return "CalculationChain";
            case XLRelationshipType::ExternalLink: return "ExternalLink";
            case XLRelationshipType::ExternalLinkPath: return "ExternalLinkPath";
            case XLRelationshipType::Theme: return "Theme";
            case XLRelationshipType::Styles: return "Styles";
            case XLRelationshipType::Chart: return "Chart";
            case XLRelationshipType::ChartStyle: return "ChartStyle";
            case XLRelationshipType::ChartColorStyle: return "ChartColorStyle";
            case XLRelationshipType::Image: return "Image";
            case XLRelationshipType::Drawing: return "Drawing";
            case XLRelationshipType::VMLDrawing: return "VMLDrawing";
            case XLRelationshipType::SharedStrings: return "SharedStrings";
            case XLRelationshipType::PrinterSettings: return "PrinterSettings";
            case XLRelationshipType::VBAProject: return "VBAProject";
            case XLRelationshipType::ControlProperties: return "ControlProperties";
            case XLRelationshipType::Comments: return "Comments";
            case XLRelationshipType::Table: return "Table";
            case XLRelationshipType::Unknown: return "Unknown";
        }
        return "invalid";
    }

    /**
     * @details
     */
    inline XMLNode getRowNode(XMLNode sheetDataNode, uint32_t rowNumber)
    {
        if (rowNumber < 1 || rowNumber > OpenXLSX::MAX_ROWS) {    // 2024-05-28: added range check
            using namespace std::literals::string_literals;
            throw XLCellAddressError("rowNumber "s + std::to_string( rowNumber ) + " is outside valid range [1;"s + std::to_string(OpenXLSX::MAX_ROWS) + "]"s);
        }

        // ===== Get the last child of sheetDataNode that is of type node_element.
        XMLNode result = sheetDataNode.last_child_of_type(pugi::node_element);

        // ===== If there are now rows in the worksheet, or the requested row is beyond the current max row, append a new row to the end.
        if (result.empty() || (rowNumber > result.attribute("r").as_ullong())) {
            result = sheetDataNode.append_child("row");
            result.append_attribute("r") = rowNumber;
            //            result.append_attribute("x14ac:dyDescent") = "0.2";
            //            result.append_attribute("spans")           = "1:1";
        }

        // ===== If the requested node is closest to the end, start from the end and search backwards.
        else if (result.attribute("r").as_ullong() - rowNumber < rowNumber) {
            while (not result.empty() && (result.attribute("r").as_ullong() > rowNumber)) result = result.previous_sibling_of_type(pugi::node_element);
            // ===== If the backwards search failed to locate the requested row
            if (result.empty() || (result.attribute("r").as_ullong() != rowNumber)) {
                if (result.empty())
                    result = sheetDataNode.prepend_child("row"); // insert a new row node at datasheet begin. When saving, this will keep whitespace formatting towards next row node
                else
                    result = sheetDataNode.insert_child_after("row", result);
                result.append_attribute("r") = rowNumber;
                //                result.append_attribute("x14ac:dyDescent") = "0.2";
                //                result.append_attribute("spans")           = "1:1";
            }
        }

        // ===== Otherwise, start from the beginning
        else {
            // ===== At this point, it is guaranteed that there is at least one node_element in the row that is not empty.
            result = sheetDataNode.first_child_of_type(pugi::node_element);

            // ===== It has been verified above that the requested rowNumber is <= the row number of the last node_element, therefore this loop will halt.
            while (result.attribute("r").as_ullong() < rowNumber) result = result.next_sibling_of_type(pugi::node_element);
            // ===== If the forwards search failed to locate the requested row
            if (result.attribute("r").as_ullong() > rowNumber) {
                result = sheetDataNode.insert_child_before("row", result);

                result.append_attribute("r") = rowNumber;
                //                result.append_attribute("x14ac:dyDescent") = "0.2";
                //                result.append_attribute("spans")           = "1:1";
            }
        }

        return result;
    }

    /**
     * @brief get the style attribute s for the indicated column, if any is set
     * @param rowNode the row node from which to obtain the parent that should hold the <cols> node
     * @param colNo the column for which to obtain the style
     * @return the XLStyleIndex stored for the column, or XLDefaultCellFormat if none
     */
    inline XLStyleIndex getColumnStyle(XMLNode rowNode, uint16_t colNo)
    {
        XMLNode cols = rowNode.parent().parent().child("cols");
        if (not cols.empty()) {
            XMLNode col = cols.first_child_of_type(pugi::node_element);
            while (not col.empty()) {
                if (col.attribute("min").as_int(MAX_COLS+1) <= colNo && col.attribute("max").as_int(0) >= colNo) // found
                    return col.attribute("style").as_uint(XLDefaultCellFormat);
               col = col.next_sibling_of_type(pugi::node_element);
            }
        }
        return XLDefaultCellFormat; // if no col style was found
    }

    /**
     * @brief set the cell reference, and a default cell style attribute if and only if row or column style is != XLDefaultCellFormat
     * @note row style takes precedence over column style
     * @param cellNode the cell XML node
     * @param cellRef the cell reference (attribute r) to set
     * @param rowNode the row node for the cell
     * @param colNo the column number for this cell (to try and fetch a column style)
     * @param colStyles an optional std::vector<XLStyleIndex> that contains all pre-evaluated column styles,
     *         and can be used to avoid performance impact from lookup
     */
    inline void setDefaultCellAttributes(XMLNode cellNode, const std::string & cellRef, XMLNode rowNode, uint16_t colNo,
    /**/                                 std::vector<XLStyleIndex> const & colStyles = {})
    {
        cellNode.append_attribute("r").set_value(cellRef.c_str());
        XMLAttribute rowStyle = rowNode.attribute("s");
        XLStyleIndex cellStyle;    // scope declaration
        if (rowStyle.empty()) {    // if no explicit row style is set
            if (colStyles.size() > 0)    // if colStyles were provided
                cellStyle = (colNo > colStyles.size() ? XLDefaultCellFormat : colStyles[ colNo - 1 ]);    // determine cellStyle from colStyles
            else                         // else: no colStyles provided
                cellStyle = getColumnStyle(rowNode, colNo);                                               // so perform a lookup
        }
        else                       // else: an explicit row style is set
            cellStyle = rowStyle.as_uint(XLDefaultCellFormat);    // use the row style

        if (cellStyle != XLDefaultCellFormat) // if cellStyle was determined as not the default style (no point in setting that)
            cellNode.append_attribute("s").set_value(cellStyle);
    }

    /**
     * @brief Retrieve the xml node representing the cell at the given row and column. If the node doesn't
     * exist, it will be created.
     * @param rowNode The row node under which to find the cell.
     * @param columnNumber The column at which to find the cell.
     * @param rowNumber (optional) row number of the row node, if already known, defaults to 0
     * @param colStyles an optional std::vector<XLStyleIndex> that contains all pre-evaluated column styles,
     *         and can be used to avoid performance impact from lookup
     * @return The xml node representing the requested cell.
     */
    inline XMLNode getCellNode(XMLNode rowNode, uint16_t columnNumber, uint32_t rowNumber = 0,
    /**/                       std::vector<XLStyleIndex> const & colStyles = {})
    {
        if (columnNumber < 1 || columnNumber > OpenXLSX::MAX_COLS) {    // 2024-08-05: added range check
            using namespace std::literals::string_literals;
            throw XLException("XLWorksheet::column: columnNumber "s + std::to_string(columnNumber) + " is outside allowed range [1;"s + std::to_string(MAX_COLS) + "]"s);
        }
        if (rowNode.empty()) return XMLNode{};    // 2024-05-28: return an empty node in case of empty rowNode

        XMLNode cellNode = rowNode.last_child_of_type(pugi::node_element);
        if (!rowNumber) rowNumber = rowNode.attribute("r").as_uint(); // if not provided, determine from rowNode
        auto cellRef  = XLCellReference(rowNumber, columnNumber);

        // ===== If there are no cells in the current row, or the requested cell is beyond the last cell in the row...
        if (cellNode.empty() || (XLCellReference(cellNode.attribute("r").value()).column() < columnNumber)) {
            // ===== append a new node to the end.
            cellNode = rowNode.append_child("c");
            setDefaultCellAttributes(cellNode, cellRef.address(), rowNode, columnNumber, colStyles);
        }
        // ===== If the requested node is closest to the end, start from the end and search backwards...
        else if (XLCellReference(cellNode.attribute("r").value()).column() - columnNumber < columnNumber) {
            while (not cellNode.empty() && (XLCellReference(cellNode.attribute("r").value()).column() > columnNumber))
                cellNode = cellNode.previous_sibling_of_type(pugi::node_element);
            // ===== If the backwards search failed to locate the requested cell
            if (cellNode.empty() || (XLCellReference(cellNode.attribute("r").value()).column() < columnNumber)) {
                if (cellNode.empty()) // If between row begin and higher column number, only non-element nodes exist
                    cellNode = rowNode.prepend_child("c"); // insert a new cell node at row begin. When saving, this will keep whitespace formatting towards next cell node
                else
                    cellNode = rowNode.insert_child_after("c", cellNode);
                setDefaultCellAttributes(cellNode, cellRef.address(), rowNode, columnNumber, colStyles);
            }
        }
        // ===== Otherwise, start from the beginning
        else {
            // ===== At this point, it is guaranteed that there is at least one node_element in the row that is not empty.
            cellNode = rowNode.first_child_of_type(pugi::node_element);

            // ===== It has been verified above that the requested columnNumber is <= the column number of the last node_element, therefore this loop will halt:
            while (XLCellReference(cellNode.attribute("r").value()).column() < columnNumber)
                cellNode = cellNode.next_sibling_of_type(pugi::node_element);
            // ===== If the forwards search failed to locate the requested cell
            if (XLCellReference(cellNode.attribute("r").value()).column() > columnNumber) {
                cellNode = rowNode.insert_child_before("c", cellNode);
                setDefaultCellAttributes(cellNode, cellRef.address(), rowNode, columnNumber, colStyles);
            }
        }
        return cellNode;
    }


    /**
     * @brief find the index of nodeName in nodeOrder
     * @param nodeName search this
     * @param nodeOrder in this
     * @return index of nodeName in nodeOrder
     * @return -1 if nodeName is not an element of nodeOrder
     * @note this function uses a vector of std::string_view because std::string is not constexpr-capable, and in
     *        a future c++20 build, a std::span could be used to have the node order in any OpenXLSX class be a constexpr
     */
    constexpr const int SORT_INDEX_NOT_FOUND = -1;
    inline int findStringInVector( std::string const & nodeName, std::vector< std::string_view > const & nodeOrder )
    {
        for (int i = 0; static_cast< size_t >( i ) < nodeOrder.size(); ++i)
            if (nodeName == nodeOrder[ i ]) return i;
        return SORT_INDEX_NOT_FOUND;
    }

    /**
     * @brief copy all leading pc_data nodes from fromNode to toNode
     * @param parent parent node that can perform sibling insertions
     * @param fromNode node whose preceeding whitespaces shall be duplicated
     * @param toNode node before which the duplicated whitespaces shall be inserted
     * @return N/A
     */
    inline void copyLeadingWhitespaces( XMLNode & parent, XMLNode fromNode, XMLNode toNode )
    {
        fromNode = fromNode.previous_sibling(); // move to preceeding whitespace node, if any
        // loop from back to front, inserting in the same order before toNode
        while (fromNode.type() == pugi::node_pcdata) { // loop ends on pugi::node_element or node_null
            toNode = parent.insert_child_before(pugi::node_pcdata, toNode);   // prepend as toNode a new pcdata node
            toNode.set_value(fromNode.value());                               //  with the value of fromNode
            fromNode = fromNode.previous_sibling();
        }
    }

    /**
     * @brief ensure that node with nodeName exists in parent and return it
     * @param parent parent node that can perform sibling insertions
     * @param nodename name of the node to be (created &) returned
     * @param nodeOrder optional vector of a predefined element node sequence required by MS Office
     * @param force_ns optional force nodeName namespace
     * @return the requested XMLNode or an empty node if the insert operation failed
     * @note 2024-12-19: appendAndGetNode will attempt to perform an ordered insert per nodeOrder if provided
     *       Once sufficiently tested, this functionality might be generalized (e.g. in XLXmlParser OpenXLSX_xml_node)
     */
    inline XMLNode appendAndGetNode(XMLNode & parent, std::string const & nodeName, std::vector< std::string_view > const & nodeOrder = {}, bool force_ns = false )
    {
        if (parent.empty()) return XMLNode{};

        XMLNode nextNode = parent.first_child_of_type(pugi::node_element);
        if (nextNode.empty()) return parent.prepend_child(nodeName.c_str(), force_ns); // nothing to sort, whitespaces "belong" to parent closing tag

        XMLNode node{}; // empty until successfully created;

        int nodeSortIndex = (nodeOrder.size() > 1 ? findStringInVector(nodeName, nodeOrder) : SORT_INDEX_NOT_FOUND);
        if (nodeSortIndex != SORT_INDEX_NOT_FOUND) { // can't sort anything if nodeOrder contains less than 2 entries or does not contain nodeName
            // ===== Find first node to follow nodeName per nodeOrder
            while (not nextNode.empty() && findStringInVector(nextNode.name(), nodeOrder) < nodeSortIndex)
                nextNode = nextNode.next_sibling_of_type(pugi::node_element);
            // ===== Evaluate search result
            if (not nextNode.empty()) {   // found nodeName or a node before which nodeName should be inserted
                if( nextNode.name() == nodeName )  // if nodeName was found
                    node = nextNode;                  // use existing node
                else {                             // else: a node was found before which nodeName must be inserted
                    node = parent.insert_child_before(nodeName.c_str(), nextNode, force_ns); // insert before nextNode without whitespaces
                    copyLeadingWhitespaces(parent, node, nextNode);
                }
            }
            // else: no node was found before which nodeName should be inserted: proceed as usual
        }
        else // no possibility to perform ordered insert - attempt to locate existing node:
            node = parent.child(nodeName.c_str());

        if( node.empty() ) {  // neither nodeName, nor a node following per nodeOrder was found
            // ===== There is no reference to perform an ordered insert for nodeName
            nextNode = parent.last_child_of_type(pugi::node_element);               // at least one element node must exist, tested at begin of function
            node = parent.insert_child_after(nodeName.c_str(), nextNode, force_ns); // append as the last element node, but before final whitespaces
            copyLeadingWhitespaces( parent, nextNode, node );                       // duplicate the prefix whitespaces of nextNode to node
        }

        return node;
    }

    inline XMLAttribute appendAndGetAttribute(XMLNode & node, std::string const & attrName, std::string const & attrDefaultVal)
    {
        if (node.empty()) return XMLAttribute{};
        XMLAttribute attr = node.attribute(attrName.c_str());
        if (attr.empty()) {
            attr = node.append_attribute(attrName.c_str());
            attr.set_value(attrDefaultVal.c_str());
        }
        return attr;
    }

    inline XMLAttribute appendAndSetAttribute(XMLNode & node, std::string const & attrName, std::string const & attrVal)
    {
        if (node.empty()) return XMLAttribute{};
        XMLAttribute attr = node.attribute(attrName.c_str());
        if (attr.empty())
            attr = node.append_attribute(attrName.c_str());
        attr.set_value(attrVal.c_str()); // silently fails on empty attribute, which is intended here
        return attr;
    }

    /**
     * @brief ensure that node with nodeName exists in parent, has an attribute with attrName and return that attribute
     * @param parent parent node that can perform sibling insertions
     * @param nodename name of the node under which attribute attrName shall exist
     * @param attrName name of the attribute to get for node nodeName
     * @param attrDefaultVal value to assign to the attribute if it has to be created
     * @param nodeOrder optional vector of a predefined element node sequence required by MS Office, passed through to appendAndGetNode
     * @returns the requested XMLAttribute or an empty node if the operation failed
     */
    inline XMLAttribute appendAndGetNodeAttribute(XMLNode & parent, std::string const & nodeName, std::string const & attrName, std::string const & attrDefaultVal,
    /**/                                   std::vector< std::string_view > const & nodeOrder = {})
    {
        if (parent.empty()) return XMLAttribute{};
        XMLNode node = appendAndGetNode(parent, nodeName, nodeOrder);
        return appendAndGetAttribute(node, attrName, attrDefaultVal);
    }

    /**
     * @brief ensure that node with nodeName exists in parent, has an attribute with attrName, set attribute value and return that attribute
     * @param parent parent node that can perform sibling insertions
     * @param nodename name of the node under which attribute attrName shall exist
     * @param attrName name of the attribute to set for node nodeName
     * @param attrVal value to assign to the attribute
     * @param removeAttributes if true, all other attributes of the node with nodeName will be deleted
     * @param nodeOrder optional vector of a predefined element node sequence required by MS Office, passed through to appendAndGetNode
     * @returns the XMLAttribute that was modified or an empty node if the operation failed
     */
    inline XMLAttribute appendAndSetNodeAttribute(XMLNode & parent, std::string const & nodeName, std::string const & attrName, std::string const & attrVal,
    /**/                                   bool removeAttributes = XLKeepAttributes, std::vector< std::string_view > const & nodeOrder = {})
    {
        if (parent.empty()) return XMLAttribute{};
        XMLNode node = appendAndGetNode(parent, nodeName, nodeOrder);
        if (removeAttributes) node.remove_attributes();
        return appendAndSetAttribute(node, attrName, attrVal);
    }

    /**
     * @brief special bool attribute getter function for tags that should have a val="true" or val="false" attribute,
     *       but when omitted shall default to "true"
     * @param parent node under which tagName shall be found
     * @param tagName name of the boolean tag to evaluate
     * @param attrName (default: "val") name of boolean attribute that shall default to true
     * @returns true if parent & tagName exist, and attribute with attrName is either omitted or as_bool() returns true. Otherwise return false
     * @note this will create and explicitly set attrName if omitted
     */
    inline bool getBoolAttributeWhenOmittedMeansTrue( XMLNode & parent, std::string const & tagName, std::string const & attrName = "val" )
    {
        if (parent.empty()) return false;     // can't do anything
        XMLNode tagNode = parent.child(tagName.c_str());
        if( tagNode.empty() ) return false;   // if tag does not exist: return false
        XMLAttribute valAttr = tagNode.attribute(attrName.c_str());
        if( valAttr.empty() ) {               // if no attribute with attrName exists: default to true
            appendAndSetAttribute(tagNode, attrName, "true" ); // explicitly create & set attribute
            return true;
        }
        // if execution gets here: attribute with attrName exists
        return valAttr.as_bool(); // return attribute value
    }

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLUTILITIES_HPP

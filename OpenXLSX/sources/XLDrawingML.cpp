/*
 * XLDrawingML.cpp
 * 
 * DrawingML (Drawing Markup Language) implementation for OpenXLSX
 */

// ===== External Includes ===== //
#include <cstdint>
#include <iostream>
#include <string>

// ===== OpenXLSX Includes ===== //
#include "XLDrawingML.hpp"
#include "XLImage.hpp"
#include "XLXmlData.hpp"
#include "XLCellReference.hpp"
#include "utilities/XLUtilities.hpp"

namespace OpenXLSX
{
    // ===== Constructors & Destructors ===== //
    XLDrawingML::XLDrawingML(XLXmlData* xmlData) : XLXmlFile(xmlData) 
    {
    }

    void XLDrawingML::addImage(const XLEmbeddedImage& embImage){
        const XLImageAnchor& imageAnchor = embImage.getImageAnchor();
        XMLNode rootNode = xmlDocument().document_element();
        (void)XLDrawingML::createXMLNode(rootNode, imageAnchor);
        }

    // ===== Public Methods ===== //

    size_t XLDrawingML::imageCount() const
    {
        if (!valid()) {
            return 0;
        }
        
        XMLNode rootNode = xmlDocument().document_element();
        size_t count = 0;
        
        for (XMLNode child : rootNode.children("xdr:oneCellAnchor")) {
            count++;
        }
        for (XMLNode child : rootNode.children("xdr:twoCellAnchor")) {
            count++;
        }
        
        return count;
    }

    // ===== Private Methods ===== //

    uint32_t XLDrawingML::calculateCellPosition(uint32_t row, uint16_t column) const
    {
        // Calculate cell position in Excel units
        // This is a simplified calculation - Excel uses more complex formulas
        return (row - 1) * 15 + (column - 1) * 64;
    }

    int XLDrawingML::verifyData(std::string* dbgMsg) const
    {
        int errorCount = 0;

        // Check if XML is valid
        if (!valid()) {
            appendDbgMsg(dbgMsg, "DrawingML XML is invalid");
            errorCount++;
            return errorCount;
        }

        try {
            // Check XML structure
            XMLNode rootNode = xmlDocument().document_element();
            if (rootNode.empty()) {
                appendDbgMsg(dbgMsg, "DrawingML root node is empty");
                errorCount++;
            }

            // Count images in XML and verify structure
            size_t xmlImageCount = 0;
            for (auto anchor : rootNode.children()) {
                // Check anchor type using enum conversion
                XLImageAnchorType anchorType = XLImageAnchorUtils::stringToAnchorType(anchor.name());
                if (anchorType != XLImageAnchorType::Unknown) {
                    xmlImageCount++;
                    
                    // Check for required elements
                    if (anchor.child("xdr:pic").empty()) {
                        appendDbgMsg(dbgMsg, "image anchor missing xdr:pic element");
                        errorCount++;
                    }
                    
                    if (anchor.child("xdr:clientData").empty()) {
                        appendDbgMsg(dbgMsg, "image anchor missing xdr:clientData element");
                        errorCount++;
                    }
                }
            }

            // Verify image count consistency
            size_t reportedCount = imageCount();
            if (xmlImageCount != reportedCount) {
                appendDbgMsg(dbgMsg, "XML image count (" + std::to_string(xmlImageCount) + 
                          ") does not match reported count (" + std::to_string(reportedCount) + ")");
                errorCount++;
            }

        } catch (const std::exception&) {
            appendDbgMsg(dbgMsg, "error parsing DrawingML XML");
            errorCount++;
        }

        return errorCount;
    }

    XMLNode XLDrawingML::getRootNode() const
    {
        if (!valid()) {
            return XMLNode(); // Return empty node if invalid
        }
        
        try {
            return xmlDocument().document_element();
        } catch (const std::exception&) {
            return XMLNode(); // Return empty node on error
        }
    }

    // ========== XLImageAnchor Implementation ========== //

    XLImageAnchor::XLImageAnchor(const std::string& imageId, const std::string& relationshipId,
                                 uint32_t row, uint16_t col, int32_t rowOffset, int32_t colOffset,
                                 uint32_t displayWidth, uint32_t displayHeight)
        : anchorType(XLImageAnchorType::OneCellAnchor)
        , imageId(imageId)
        , relationshipId(relationshipId)
        , fromRow(row)          // Store as-is (0-based from XML)
        , fromCol(col)          // Store as-is (0-based from XML)
        , fromRowOffset(rowOffset)
        , fromColOffset(colOffset)
        , toRow(row)            // For oneCellAnchor, to = from
        , toCol(col)            // For oneCellAnchor, to = from
        , toRowOffset(0)        // No offset for 'to' in oneCellAnchor
        , toColOffset(0)        // No offset for 'to' in oneCellAnchor
        , displayWidthEMUs(displayWidth)
        , displayHeightEMUs(displayHeight)
        , actualWidthEMUs(displayWidth)
        , actualHeightEMUs(displayHeight)
    {
    }

    XLImageAnchor::XLImageAnchor(const std::string& imageId, const std::string& relationshipId,
                                 uint32_t fromRow, uint16_t fromCol, uint32_t toRow, uint16_t toCol,
                                 int32_t fromRowOffset, int32_t fromColOffset,
                                 int32_t toRowOffset, int32_t toColOffset,
                                 uint32_t displayWidth, uint32_t displayHeight)
        : anchorType(XLImageAnchorType::TwoCellAnchor)
        , imageId(imageId)
        , relationshipId(relationshipId)
        , fromRow(fromRow)      // Store as-is (0-based from XML)
        , fromCol(fromCol)      // Store as-is (0-based from XML)
        , fromRowOffset(fromRowOffset)
        , fromColOffset(fromColOffset)
        , toRow(toRow)          // Store as-is (0-based from XML)
        , toCol(toCol)          // Store as-is (0-based from XML)
        , toRowOffset(toRowOffset)
        , toColOffset(toColOffset)
        , displayWidthEMUs(displayWidth)
        , displayHeightEMUs(displayHeight)
        , actualWidthEMUs(displayWidth)
        , actualHeightEMUs(displayHeight)
    {
    }

    void XLImageAnchor::reset(){
        anchorType = XLImageAnchorType::Unknown;
        imageId.erase();
        relationshipId.erase();
        fromRow = 0;
        fromCol = 0;
        fromRowOffset = 0;
        fromColOffset = 0;
        toRow = 0;
        toCol = 0;
        toRowOffset = 0;
        toColOffset = 0;
        displayWidthEMUs = 0;
        displayHeightEMUs = 0;
        actualWidthEMUs = 0;
        actualHeightEMUs = 0;
    }

    void XLImageAnchor::initOneCell( const XLCellReference& cellRef, int32_t rowOffset,
        int32_t colOffset){
        reset();
        anchorType = XLImageAnchorType::OneCellAnchor;
        setFromCellReference(cellRef);
        fromRowOffset = rowOffset;
        fromColOffset = colOffset;
    }

    void XLImageAnchor::initTwoCell( const XLCellReference& fromCellRef, 
        const XLCellReference& toCellRef, int32_t fromROffset,
        int32_t fromCOffset, int32_t toROffset, int32_t toCOffset){
        reset();
        anchorType = XLImageAnchorType::TwoCellAnchor;
        setFromCellReference(fromCellRef);
        setToCellReference(toCellRef);
        fromRowOffset = fromROffset;
        fromColOffset = fromCOffset;
        toRowOffset = toROffset;
        toColOffset = toCOffset;
    }
        
    XLCellReference XLImageAnchor::getFromCellReference() const{
        const XLCellReference fromCellReference( fromRow + 1, fromCol + 1 );
        return fromCellReference;
    }

    void XLImageAnchor::setFromCellReference( const XLCellReference& fromCellRef ){
        const uint32_t cellRefRow = fromCellRef.row();
        const uint16_t cellRefColumn = fromCellRef.column();
        fromRow = ( cellRefRow > 0 ) ? (cellRefRow - 1) : 0;
        fromCol = ( cellRefColumn > 0 ) ? (cellRefColumn - 1) : 0;
    }

    XLCellReference XLImageAnchor::getToCellReference() const{
        const XLCellReference toCellReference( toRow + 1, toCol + 1 );
        return toCellReference;
    }

    void XLImageAnchor::setToCellReference( const XLCellReference& toCellRef ){
        const uint32_t cellRefRow = toCellRef.row();
        const uint16_t cellRefColumn = toCellRef.column();
        toRow = ( cellRefRow > 0 ) ? (cellRefRow - 1) : 0;
        toCol = ( cellRefColumn > 0 ) ? (cellRefColumn - 1) : 0;
    }

    void XLImageAnchor::setDisplaySizeWithAspectRatio(
            const uint32_t& widthPixels, const uint32_t& heightPixels,
            const uint32_t& maxWidthEmus, const uint32_t& maxHeightEmus ){
        uint32_t targetWidth = XLImageUtils::pixelsToEmus(widthPixels);
        uint32_t targetHeight = XLImageUtils::pixelsToEmus(heightPixels);

        if (widthPixels > 0 || heightPixels > 0) {
            double aspectRatio = XLImageUtils::calculateAspectRatio(widthPixels, heightPixels);
            assert(aspectRatio > 0.0);

            const double maxWidth = static_cast<double>(maxWidthEmus);
            const double maxHeight = static_cast<double>(maxHeightEmus);
            const double widthA = maxHeight * aspectRatio;

            if (widthA <= maxWidth) {
                /*  
                +------------+--------+
                |            |        |
                |            |        |
                |            |        |
                |            |        |maxHeight
                |            |        |
                |            |        |
                |            |        |
                |   widthA   |        |
                +------------+--------+
                       maxWidth 
                */
                targetWidth = static_cast<uint32_t>(round(widthA));
                targetHeight = maxHeightEmus;
            }
            else {
                /* 
                +---------------------+
                |                     |
                |                     |
                +---------------------+
                |                     |maxHeight
                |                     |
                |              heightB|
                |                     |
                |                     |
                +---------------------+
                       maxWidth  
                */
                const double heightB = maxWidth / aspectRatio;
                assert(heightB <= (1.00000001 * maxHeight));
                targetWidth = maxWidthEmus;
                targetHeight = static_cast<uint32_t>(round(heightB));
            }

            assert(fabs(aspectRatio - (static_cast<double>(targetWidth) / 
                    static_cast<double>(targetHeight))) < 1e-4);
        }

        displayWidthEMUs = targetWidth;
        displayHeightEMUs = targetHeight;
    }

    void XLImageAnchor::setDisplaySizeWithAspectRatio(
        const std::vector<uint8_t>& imageData, const XLMimeType& mimeType,
        const uint32_t& maxWidthEmus, const uint32_t& maxHeightEmus ){

        // Convert string mimeType to enum, auto-detect if unknown
        XLMimeType mimeTypeEnum = mimeType;
        if (mimeTypeEnum == XLMimeType::Unknown) {
            mimeTypeEnum = XLImageUtils::detectMimeTypeEnum(imageData);
        }

        const std::pair<uint32_t, uint32_t> imageDims =
            XLImageUtils::getImageDimensions(imageData, mimeType);

        setDisplaySizeWithAspectRatio( imageDims.first, imageDims.second,
            maxWidthEmus, maxHeightEmus );
    }

    void XLImageAnchor::setDisplaySizeWithAspectRatio(
        const std::string& imageFileName, const XLMimeType& mimeType,
        const uint32_t& maxWidthEmus, const uint32_t& maxHeightEmus ){

        std::ifstream file(imageFileName, std::ios::binary);
        if (file.is_open()) {
            // Read the entire file into a vector
            file.seekg(0, std::ios::end);
            size_t fileSize = file.tellg();
            file.seekg(0, std::ios::beg);

            std::vector<uint8_t> imageData(fileSize);
            file.read(reinterpret_cast<char*>(imageData.data()), fileSize);
            file.close();

            setDisplaySizeWithAspectRatio( imageData, mimeType, maxWidthEmus,
                maxHeightEmus );
        }
    }


    std::string XLImageAnchor::getAnchorCellReference() const
    {
        // Convert 0-based coordinates to 1-based for cell reference
        return XLCellReference(fromRow + 1, fromCol + 1).address();
    }

    bool XLImageAnchor::isTwoCell() const
    {
        return anchorType == XLImageAnchorType::TwoCellAnchor;
    }

    std::pair<uint32_t, uint32_t> XLImageAnchor::getDisplayDimensions() const
    {
        return std::make_pair(displayWidthEMUs, displayHeightEMUs);
    }

    std::pair<uint32_t, uint32_t> XLImageAnchor::getActualDimensions() const
    {
        return std::make_pair(actualWidthEMUs, actualHeightEMUs);
    }

    void XLImageAnchor::setActualDimensions(uint32_t width, uint32_t height)
    {
        actualWidthEMUs = width;
        actualHeightEMUs = height;
    }

    std::pair<uint32_t, uint32_t> XLImageAnchor::getDisplayDimensionsPixels() const
    {
        return std::make_pair(
            XLImageUtils::emusToPixels(displayWidthEMUs),
            XLImageUtils::emusToPixels(displayHeightEMUs)
        );
    }

    std::pair<uint32_t, uint32_t> XLImageAnchor::getActualDimensionsPixels() const
    {
        return std::make_pair(
            XLImageUtils::emusToPixels(actualWidthEMUs),
            XLImageUtils::emusToPixels(actualHeightEMUs)
        );
    }

    // ========== XLDrawingML Static Utility Functions ========== //

    XMLNode XLDrawingML::createXMLNode(XMLNode rootNode, const XLImageAnchor& imgAnchorInfo)
    {
        if (rootNode.empty()) {
            return XMLNode();
        }

        // Create the appropriate anchor element
        XMLNode anchor;
        if (imgAnchorInfo.isTwoCell()) {
            anchor = rootNode.append_child("xdr:twoCellAnchor");
            anchor.append_attribute("editAs").set_value("oneCell");
        } else {
            anchor = rootNode.append_child("xdr:oneCellAnchor");
        }

        // Create the 'from' element
        XMLNode from = anchor.append_child("xdr:from");
        from.append_child("xdr:col").append_child(pugi::node_pcdata).set_value(std::to_string(imgAnchorInfo.fromCol).c_str());
        from.append_child("xdr:colOff").append_child(pugi::node_pcdata).set_value(std::to_string(imgAnchorInfo.fromColOffset).c_str());
        from.append_child("xdr:row").append_child(pugi::node_pcdata).set_value(std::to_string(imgAnchorInfo.fromRow).c_str());
        from.append_child("xdr:rowOff").append_child(pugi::node_pcdata).set_value(std::to_string(imgAnchorInfo.fromRowOffset).c_str());

        // Create the 'to' element (for twoCellAnchor)
        if (imgAnchorInfo.isTwoCell()) {
            XMLNode to = anchor.append_child("xdr:to");
            to.append_child("xdr:col").append_child(pugi::node_pcdata).set_value(std::to_string(imgAnchorInfo.toCol).c_str());
            to.append_child("xdr:colOff").append_child(pugi::node_pcdata).set_value(std::to_string(imgAnchorInfo.toColOffset).c_str());
            to.append_child("xdr:row").append_child(pugi::node_pcdata).set_value(std::to_string(imgAnchorInfo.toRow).c_str());
            to.append_child("xdr:rowOff").append_child(pugi::node_pcdata).set_value(std::to_string(imgAnchorInfo.toRowOffset).c_str());
        }

        // Create the 'ext' element (required for both anchor types)
        XMLNode ext = anchor.append_child("xdr:ext");
        ext.append_attribute("cx").set_value(std::to_string(imgAnchorInfo.displayWidthEMUs).c_str());
        ext.append_attribute("cy").set_value(std::to_string(imgAnchorInfo.displayHeightEMUs).c_str());

        // Create the picture element
        XMLNode pic = anchor.append_child("xdr:pic");

        // Create non-visual picture properties
        XMLNode nvPicPr = pic.append_child("xdr:nvPicPr");
        XMLNode cNvPr = nvPicPr.append_child("xdr:cNvPr");
        cNvPr.append_attribute("id").set_value(imgAnchorInfo.imageId.c_str());
        cNvPr.append_attribute("name").set_value(("Picture " + imgAnchorInfo.imageId).c_str());

        XMLNode cNvPicPr = nvPicPr.append_child("xdr:cNvPicPr");
        XMLNode picLocks = cNvPicPr.append_child("a:picLocks");
        picLocks.append_attribute("noChangeAspect").set_value("1");
        picLocks.append_attribute("noChangeArrowheads").set_value("1");

        // Create blip fill
        XMLNode blipFill = pic.append_child("xdr:blipFill");
        XMLNode blip = blipFill.append_child("a:blip");
        blip.append_attribute("r:embed").set_value(imgAnchorInfo.relationshipId.c_str());

        XMLNode srcRect = blipFill.append_child("a:srcRect");
        XMLNode stretch = blipFill.append_child("a:stretch");
        stretch.append_child("a:fillRect");

        // Create shape properties
        XMLNode spPr = pic.append_child("xdr:spPr");
        XMLNode xfrm = spPr.append_child("a:xfrm");
        XMLNode off = xfrm.append_child("a:off");
        off.append_attribute("x").set_value("0");
        off.append_attribute("y").set_value("0");

        XMLNode xfrmExt = xfrm.append_child("a:ext");
        xfrmExt.append_attribute("cx").set_value(std::to_string(imgAnchorInfo.actualWidthEMUs).c_str());
        xfrmExt.append_attribute("cy").set_value(std::to_string(imgAnchorInfo.actualHeightEMUs).c_str());

        XMLNode prstGeom = spPr.append_child("a:prstGeom");
        prstGeom.append_attribute("prst").set_value("rect");
        prstGeom.append_child("a:avLst");

        // Use noFill to match other methods (no borders)
        spPr.append_child("a:noFill");

        // Add clientData element (required for Excel compatibility)
        anchor.append_child("xdr:clientData");

        return anchor;
    }

    bool XLDrawingML::deleteXMLNode(XMLNode rootNode, const std::string& relationshipId)
    {
        if (rootNode.empty() || relationshipId.empty()) {
            return false;
        }

        // Find and remove the image node
        for (auto anchor : rootNode.children()) {
            // Check both oneCellAnchor and twoCellAnchor using enum conversion
            XLImageAnchorType anchorType = XLImageAnchorUtils::stringToAnchorType(anchor.name());
            if (anchorType != XLImageAnchorType::Unknown) {
                
                // Look for the pic element with the matching relationship ID
                XMLNode pic = anchor.child("xdr:pic");
                if (pic.empty()) {
                    pic = anchor.child("pic");
                }
                
                if (!pic.empty()) {
                    // Check blipFill -> blip -> r:embed
                    XMLNode blipFill = pic.child("xdr:blipFill");
                    if (blipFill.empty()) {
                        blipFill = pic.child("a:blipFill");
                    }
                    if (blipFill.empty()) {
                        blipFill = pic.child("blipFill");
                    }
                    
                    if (!blipFill.empty()) {
                        XMLNode blip = blipFill.child("a:blip");
                        if (blip.empty()) {
                            blip = blipFill.child("blip");
                        }
                        
                        if (!blip.empty()) {
                            std::string embedId = blip.attribute("r:embed").value();
                            if (embedId == relationshipId) {
                                // Found the matching image, remove the entire anchor
                                rootNode.remove_child(anchor);
                                return true;
                            }
                        }
                    }
                }
            }
        }
        
        return false; // Image not found
    }

    bool XLDrawingML::parseXMLNode(const pugi::xml_node& anchorXMLNode, XLImageAnchor* imgAnchorInfo)
    {
        if (anchorXMLNode.empty() || !imgAnchorInfo) {
            return false;
        }

        try {
            // Determine anchor type
            std::string nodeName = anchorXMLNode.name();
            // Determine anchor type using utility function
            imgAnchorInfo->anchorType = XLImageAnchorUtils::stringToAnchorType(nodeName);
            if (imgAnchorInfo->anchorType == XLImageAnchorType::Unknown) {
                return false; // Unknown anchor type
            }

            // Extract cell position from 'from' element
            XMLNode from = anchorXMLNode.child("xdr:from");
            if (from.empty()) {
                from = anchorXMLNode.child("from");
            }
            if (from.empty()) {
                return false;
            }

            XMLNode colNode = from.child("xdr:col");
            if (colNode.empty()) {
                colNode = from.child("col");
            }
            XMLNode rowNode = from.child("xdr:row");
            if (rowNode.empty()) {
                rowNode = from.child("row");
            }
            
            if (colNode.empty() || rowNode.empty()) {
                return false;
            }

            imgAnchorInfo->fromCol = colNode.text().as_uint();
            imgAnchorInfo->fromRow = rowNode.text().as_uint();

            // Extract offsets
            XMLNode colOffNode = from.child("xdr:colOff");
            if (colOffNode.empty()) {
                colOffNode = from.child("colOff");
            }
            XMLNode rowOffNode = from.child("xdr:rowOff");
            if (rowOffNode.empty()) {
                rowOffNode = from.child("rowOff");
            }

            imgAnchorInfo->fromColOffset = colOffNode.empty() ? 0 : colOffNode.text().as_int();
            imgAnchorInfo->fromRowOffset = rowOffNode.empty() ? 0 : rowOffNode.text().as_int();

            // Extract 'to' element for twoCellAnchor
            if (imgAnchorInfo->isTwoCell()) {
                XMLNode to = anchorXMLNode.child("xdr:to");
                if (to.empty()) {
                    to = anchorXMLNode.child("to");
                }
                if (!to.empty()) {
                    XMLNode toColNode = to.child("xdr:col");
                    if (toColNode.empty()) {
                        toColNode = to.child("col");
                    }
                    XMLNode toRowNode = to.child("xdr:row");
                    if (toRowNode.empty()) {
                        toRowNode = to.child("row");
                    }

                    if (!toColNode.empty() && !toRowNode.empty()) {
                        imgAnchorInfo->toCol = toColNode.text().as_uint();
                        imgAnchorInfo->toRow = toRowNode.text().as_uint();
                    }

                    // Extract 'to' offsets
                    XMLNode toColOffNode = to.child("xdr:colOff");
                    if (toColOffNode.empty()) {
                        toColOffNode = to.child("colOff");
                    }
                    XMLNode toRowOffNode = to.child("xdr:rowOff");
                    if (toRowOffNode.empty()) {
                        toRowOffNode = to.child("rowOff");
                    }

                    imgAnchorInfo->toColOffset = toColOffNode.empty() ? 0 : toColOffNode.text().as_int();
                    imgAnchorInfo->toRowOffset = toRowOffNode.empty() ? 0 : toRowOffNode.text().as_int();
                }
            } else {
                // For oneCellAnchor, to = from
                imgAnchorInfo->toCol = imgAnchorInfo->fromCol;
                imgAnchorInfo->toRow = imgAnchorInfo->fromRow;
                imgAnchorInfo->toColOffset = 0;
                imgAnchorInfo->toRowOffset = 0;
            }

            // Extract dimensions from xdr:ext element
            XMLNode ext = anchorXMLNode.child("xdr:ext");
            if (ext.empty()) {
                ext = anchorXMLNode.child("ext");
            }
            if (!ext.empty()) {
                imgAnchorInfo->displayWidthEMUs = ext.attribute("cx").as_uint();
                imgAnchorInfo->displayHeightEMUs = ext.attribute("cy").as_uint();
            }

            // Extract image information from pic element
            XMLNode pic = anchorXMLNode.child("xdr:pic");
            if (pic.empty()) {
                pic = anchorXMLNode.child("pic");
            }
            if (!pic.empty()) {
                // Extract image ID
                XMLNode nvPicPr = pic.child("xdr:nvPicPr");
                if (nvPicPr.empty()) {
                    nvPicPr = pic.child("nvPicPr");
                }
                if (!nvPicPr.empty()) {
                    XMLNode cNvPr = nvPicPr.child("xdr:cNvPr");
                    if (cNvPr.empty()) {
                        cNvPr = nvPicPr.child("cNvPr");
                    }
                    if (!cNvPr.empty()) {
                        imgAnchorInfo->imageId = cNvPr.attribute("id").value();
                    }
                }

                // Extract relationship ID
                XMLNode blipFill = pic.child("xdr:blipFill");
                if (blipFill.empty()) {
                    blipFill = pic.child("a:blipFill");
                }
                if (blipFill.empty()) {
                    blipFill = pic.child("blipFill");
                }
                if (!blipFill.empty()) {
                    XMLNode blip = blipFill.child("a:blip");
                    if (blip.empty()) {
                        blip = blipFill.child("blip");
                    }
                    if (!blip.empty()) {
                        imgAnchorInfo->relationshipId = blip.attribute("r:embed").value();
                    }
                }

                // Extract actual dimensions from shape properties
                XMLNode spPr = pic.child("xdr:spPr");
                if (spPr.empty()) {
                    spPr = pic.child("spPr");
                }
                if (!spPr.empty()) {
                    XMLNode xfrm = spPr.child("a:xfrm");
                    if (!xfrm.empty()) {
                        XMLNode xfrmExt = xfrm.child("a:ext");
                        if (!xfrmExt.empty()) {
                            imgAnchorInfo->actualWidthEMUs = xfrmExt.attribute("cx").as_uint();
                            imgAnchorInfo->actualHeightEMUs = xfrmExt.attribute("cy").as_uint();
                        }
                    }
                }
            }

            // If actual dimensions weren't found, use display dimensions
            if (imgAnchorInfo->actualWidthEMUs == 0) {
                imgAnchorInfo->actualWidthEMUs = imgAnchorInfo->displayWidthEMUs;
            }
            if (imgAnchorInfo->actualHeightEMUs == 0) {
                imgAnchorInfo->actualHeightEMUs = imgAnchorInfo->displayHeightEMUs;
            }

            return true;
        }
        catch (const std::exception&) {
            return false;
        }
    }
}

// ========== XLImageAnchorUtils Implementation ========== //

OpenXLSX::XLImageAnchorType OpenXLSX::XLImageAnchorUtils::stringToAnchorType(const std::string& anchorTypeStr)
{
    if (anchorTypeStr.find("oneCellAnchor") != std::string::npos) return XLImageAnchorType::OneCellAnchor;
    if (anchorTypeStr.find("twoCellAnchor") != std::string::npos) return XLImageAnchorType::TwoCellAnchor;
    return XLImageAnchorType::Unknown;
}

std::string OpenXLSX::XLImageAnchorUtils::anchorTypeToString(XLImageAnchorType anchorType)
{
    switch (anchorType) {
        case XLImageAnchorType::OneCellAnchor: return "oneCellAnchor";
        case XLImageAnchorType::TwoCellAnchor: return "twoCellAnchor";
        case XLImageAnchorType::Unknown:
        default: return "";
    }
}

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
#include "utilities/XLUtilities.hpp"

namespace OpenXLSX
{
    // ===== Constructors & Destructors ===== //
    XLDrawingML::XLDrawingML(XLXmlData* xmlData) : XLXmlFile(xmlData) 
    {
    }

    // ===== Public Methods ===== //
    void XLDrawingML::addImage(const XLImage& image, uint32_t row, uint16_t column, const std::string& relationshipId)
    {
        // Get the root element
        XMLNode rootNode = xmlDocument().document_element();
        
        // Create a oneCellAnchor element for single-cell images
        XMLNode anchor = rootNode.append_child("xdr:oneCellAnchor");
        
        // Create the 'from' element
        XMLNode from = anchor.append_child("xdr:from");
        from.append_child("xdr:col").append_child(pugi::node_pcdata).set_value(std::to_string(column - 1).c_str());
        from.append_child("xdr:colOff").append_child(pugi::node_pcdata).set_value("0");
        from.append_child("xdr:row").append_child(pugi::node_pcdata).set_value(std::to_string(row - 1).c_str());
        from.append_child("xdr:rowOff").append_child(pugi::node_pcdata).set_value("0");
        
        // Create the 'ext' element - this is crucial for oneCellAnchor!
        XMLNode ext = anchor.append_child("xdr:ext");
        ext.append_attribute("cx").set_value(std::to_string(image.displayWidth()).c_str());
        ext.append_attribute("cy").set_value(std::to_string(image.displayHeight()).c_str());
        
        // Create the picture element
        XMLNode pic = anchor.append_child("xdr:pic");
        
        // Create non-visual picture properties
        XMLNode nvPicPr = pic.append_child("xdr:nvPicPr");
        XMLNode cNvPr = nvPicPr.append_child("xdr:cNvPr");
        cNvPr.append_attribute("id").set_value(std::to_string(imageCount() + 2).c_str()); // Start from 2
        cNvPr.append_attribute("name").set_value(("Picture " + std::to_string(imageCount() + 1)).c_str());
        
        // Add extension list with creation ID for better compatibility
        // XMLNode cNvPrExtLst = cNvPr.append_child("a:extLst");
        // XMLNode cNvPrExt = cNvPrExtLst.append_child("a:ext");
        // cNvPrExt.append_attribute("uri").set_value("{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}");
        // XMLNode creationId = cNvPrExt.append_child("a16:creationId");
        // std::string creationIdValue = "{00000000-0008-0000-0000-00000" + std::to_string(imageCount() + 1) + "00000}";
        // creationId.append_attribute("id").set_value(creationIdValue.c_str());
        
        XMLNode cNvPicPr = nvPicPr.append_child("xdr:cNvPicPr");
        XMLNode picLocks = cNvPicPr.append_child("a:picLocks");
        picLocks.append_attribute("noChangeAspect").set_value("1");
        picLocks.append_attribute("noChangeArrowheads").set_value("1");
        
        // Create blip fill
        XMLNode blipFill = pic.append_child("xdr:blipFill");
        XMLNode blip = blipFill.append_child("a:blip");
        blip.append_attribute("r:embed").set_value(relationshipId.c_str());
        
        // Add extension list for better compatibility
        // XMLNode extLst = blip.append_child("a:extLst");
        // XMLNode blipExt = extLst.append_child("a:ext");
        // blipExt.append_attribute("uri").set_value("{28A0092B-C50C-407E-A947-70E740481C1C}");
        // XMLNode useLocalDpi = blipExt.append_child("a14:useLocalDpi");
        // useLocalDpi.append_attribute("val").set_value("0");
        
        XMLNode srcRect = blipFill.append_child("a:srcRect");
        XMLNode stretch = blipFill.append_child("a:stretch");
        stretch.append_child("a:fillRect");
        
        // Create shape properties
        XMLNode spPr = pic.append_child("xdr:spPr");
        spPr.append_attribute("bwMode").set_value("auto");
        
        XMLNode xfrm = spPr.append_child("a:xfrm");
        XMLNode off = xfrm.append_child("a:off");
        off.append_attribute("x").set_value("0");
        off.append_attribute("y").set_value("0");
        
        // Calculate image size that preserves aspect ratio within the bounding box
        uint32_t imageWidth = image.displayWidth();
        uint32_t imageHeight = image.displayHeight();
        
        // Get the image's natural dimensions
        uint32_t naturalWidth = image.widthPixels();
        uint32_t naturalHeight = image.heightPixels();
        
        // Calculate scaling factor to fit within bounding box while preserving aspect ratio
        double scaleX = static_cast<double>(imageWidth) / naturalWidth;
        double scaleY = static_cast<double>(imageHeight) / naturalHeight;
        double scale = std::min(scaleX, scaleY);  // Use smaller scale to fit within box
        
        // Calculate actual image size (preserving aspect ratio)
        uint32_t actualWidth = static_cast<uint32_t>(naturalWidth * scale);
        uint32_t actualHeight = static_cast<uint32_t>(naturalHeight * scale);
        
        XMLNode xfrmExt = xfrm.append_child("a:ext");
        xfrmExt.append_attribute("cx").set_value(std::to_string(actualWidth).c_str());
        xfrmExt.append_attribute("cy").set_value(std::to_string(actualHeight).c_str());
        
        XMLNode prstGeom = spPr.append_child("a:prstGeom");
        prstGeom.append_attribute("prst").set_value("rect");
        prstGeom.append_child("a:avLst");
        
        spPr.append_child("a:noFill");
        
        // Create client data - simple empty element as per Google example
        anchor.append_child("xdr:clientData");
    }

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
    uint32_t XLDrawingML::emusToExcelUnits(uint32_t emus) const
    {
        // Convert EMUs to Excel units (1 EMU = 1/9525 Excel units)
        return emus / 9525;
    }

    uint32_t XLDrawingML::pointsToEmus(double points) const
    {
        // Convert points to EMUs (1 point = 12700 EMUs)
        return static_cast<uint32_t>(points * 12700);
    }

    uint32_t XLDrawingML::calculateCellPosition(uint32_t row, uint16_t column) const
    {
        // Calculate cell position in Excel units
        // This is a simplified calculation - Excel uses more complex formulas
        return (row - 1) * 15 + (column - 1) * 64;
    }

    /**
     * @details Add an image to the drawing with precise positioning
     */
    void XLDrawingML::addImageWithOffset(const XLImage& image, uint32_t row, uint16_t column, 
                                         const std::string& relationshipId,
                                         int32_t rowOffset, int32_t colOffset)
    {
        // Get the root element
        XMLNode rootNode = xmlDocument().document_element();
        
        // Create a twoCellAnchor element with proper oneCell editing (Excel standard)
        XMLNode anchor = rootNode.append_child("xdr:twoCellAnchor");
        anchor.append_attribute("editAs").set_value("oneCell");
        
        // Create the 'from' element with offset
        XMLNode from = anchor.append_child("xdr:from");
        from.append_child("xdr:col").append_child(pugi::node_pcdata).set_value(std::to_string(column - 1).c_str());
        from.append_child("xdr:colOff").append_child(pugi::node_pcdata).set_value(std::to_string(colOffset).c_str());
        from.append_child("xdr:row").append_child(pugi::node_pcdata).set_value(std::to_string(row - 1).c_str());
        from.append_child("xdr:rowOff").append_child(pugi::node_pcdata).set_value(std::to_string(rowOffset).c_str());
        
        // Create the 'to' element (same as from for oneCell anchoring)
        XMLNode to = anchor.append_child("xdr:to");
        to.append_child("xdr:col").append_child(pugi::node_pcdata).set_value(std::to_string(column).c_str());
        to.append_child("xdr:colOff").append_child(pugi::node_pcdata).set_value("0");
        to.append_child("xdr:row").append_child(pugi::node_pcdata).set_value(std::to_string(row).c_str());
        to.append_child("xdr:rowOff").append_child(pugi::node_pcdata).set_value("0");
        
        // Create the 'ext' element (required for twoCellAnchor)
        XMLNode ext = anchor.append_child("xdr:ext");
        ext.append_attribute("cx").set_value(std::to_string(image.displayWidth()).c_str());
        ext.append_attribute("cy").set_value(std::to_string(image.displayHeight()).c_str());
        
        // Create the picture element (same as original addImage method)
        XMLNode pic = anchor.append_child("xdr:pic");
        
        // Create non-visual picture properties
        XMLNode nvPicPr = pic.append_child("xdr:nvPicPr");
        XMLNode cNvPr = nvPicPr.append_child("xdr:cNvPr");
        cNvPr.append_attribute("id").set_value(std::to_string(imageCount() + 2).c_str());
        cNvPr.append_attribute("name").set_value(("Picture " + std::to_string(imageCount() + 1)).c_str());
        
        XMLNode cNvPicPr = nvPicPr.append_child("xdr:cNvPicPr");
        XMLNode picLocks = cNvPicPr.append_child("a:picLocks");
        picLocks.append_attribute("noChangeAspect").set_value("1");
        picLocks.append_attribute("noChangeArrowheads").set_value("1");
        
        // Create blip fill
        XMLNode blipFill = pic.append_child("xdr:blipFill");
        XMLNode blip = blipFill.append_child("a:blip");
        blip.append_attribute("r:embed").set_value(relationshipId.c_str());
        
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
        xfrmExt.append_attribute("cx").set_value(std::to_string(image.displayWidth()).c_str());
        xfrmExt.append_attribute("cy").set_value(std::to_string(image.displayHeight()).c_str());
        
        XMLNode prstGeom = spPr.append_child("a:prstGeom");
        prstGeom.append_attribute("prst").set_value("rect");
        prstGeom.append_child("a:avLst");
        
        // Use noFill to match other methods (no borders)
        spPr.append_child("a:noFill");
        
        // Add clientData element (required for Excel compatibility)
        anchor.append_child("xdr:clientData");
    }

    /**
     * @details Add an image to the drawing with two-cell anchoring
     */
    void XLDrawingML::addImageTwoCellAnchor(const XLImage& image,
                                            uint32_t fromRow, uint16_t fromCol,
                                            uint32_t toRow, uint16_t toCol,
                                            const std::string& relationshipId,
                                            int32_t fromRowOffset, int32_t fromColOffset,
                                            int32_t toRowOffset, int32_t toColOffset)
    {
        // Get the root element
        XMLNode rootNode = xmlDocument().document_element();
        
        // Create a twoCellAnchor element
        XMLNode anchor = rootNode.append_child("xdr:twoCellAnchor");
        anchor.append_attribute("editAs").set_value("oneCell");
        
        // Create the 'from' element
        XMLNode from = anchor.append_child("xdr:from");
        from.append_child("xdr:col").append_child(pugi::node_pcdata).set_value(std::to_string(fromCol - 1).c_str());
        from.append_child("xdr:colOff").append_child(pugi::node_pcdata).set_value(std::to_string(fromColOffset).c_str());
        from.append_child("xdr:row").append_child(pugi::node_pcdata).set_value(std::to_string(fromRow - 1).c_str());
        from.append_child("xdr:rowOff").append_child(pugi::node_pcdata).set_value(std::to_string(fromRowOffset).c_str());
        
        // Create the 'to' element
        XMLNode to = anchor.append_child("xdr:to");
        to.append_child("xdr:col").append_child(pugi::node_pcdata).set_value(std::to_string(toCol).c_str());
        to.append_child("xdr:colOff").append_child(pugi::node_pcdata).set_value(std::to_string(toColOffset).c_str());
        to.append_child("xdr:row").append_child(pugi::node_pcdata).set_value(std::to_string(toRow).c_str());
        to.append_child("xdr:rowOff").append_child(pugi::node_pcdata).set_value(std::to_string(toRowOffset).c_str());
        
        // Create the 'ext' element (required for twoCellAnchor)
        XMLNode ext = anchor.append_child("xdr:ext");
        ext.append_attribute("cx").set_value(std::to_string(image.displayWidth()).c_str());
        ext.append_attribute("cy").set_value(std::to_string(image.displayHeight()).c_str());
        
        // Create the picture element (same as original addImage method)
        XMLNode pic = anchor.append_child("xdr:pic");
        
        // Create non-visual picture properties
        XMLNode nvPicPr = pic.append_child("xdr:nvPicPr");
        XMLNode cNvPr = nvPicPr.append_child("xdr:cNvPr");
        cNvPr.append_attribute("id").set_value(std::to_string(imageCount() + 2).c_str());
        cNvPr.append_attribute("name").set_value(("Picture " + std::to_string(imageCount() + 1)).c_str());
        
        XMLNode cNvPicPr = nvPicPr.append_child("xdr:cNvPicPr");
        XMLNode picLocks = cNvPicPr.append_child("a:picLocks");
        picLocks.append_attribute("noChangeAspect").set_value("1");
        picLocks.append_attribute("noChangeArrowheads").set_value("1");
        
        // Create blip fill
        XMLNode blipFill = pic.append_child("xdr:blipFill");
        XMLNode blip = blipFill.append_child("a:blip");
        blip.append_attribute("r:embed").set_value(relationshipId.c_str());
        
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
        xfrmExt.append_attribute("cx").set_value(std::to_string(image.displayWidth()).c_str());
        xfrmExt.append_attribute("cy").set_value(std::to_string(image.displayHeight()).c_str());
        
        XMLNode prstGeom = spPr.append_child("a:prstGeom");
        prstGeom.append_attribute("prst").set_value("rect");
        prstGeom.append_child("a:avLst");
        
        // Use noFill to match oneCellAnchor behavior (no borders)
        spPr.append_child("a:noFill");
        
        // Add clientData element (required for Excel compatibility)
        anchor.append_child("xdr:clientData");
    }
}

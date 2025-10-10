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
#include <algorithm>
#include <set>
#include <sstream>
#include <map>

// ===== OpenXLSX Includes ===== //
#include "XLSheet.hpp"
#include "XLCellReference.hpp"
#include "XLDocument.hpp"

namespace OpenXLSX
{
    //----------------------------------------------------------------------------------------------------------------------
    //           Image Query Methods Implementation
    //----------------------------------------------------------------------------------------------------------------------

    /**
     * @brief Get all image information in the worksheet
     * @return Const reference to vector of XLImageInfo objects
     */
    const std::vector<XLImageInfo>& XLWorksheet::getImageInfos() const
    {
        // Auto-refresh registry if empty
        if (m_imageRegistry.empty()) {
            refreshImageRegistry();
        }
        
        return m_imageRegistry;
    }

    /**
     * @brief Get specific image information by its ID
     * @param imageId The unique image identifier
     * @return Const reference to XLImageInfo object, or empty XLImageInfo if not found
     */
    const XLImageInfo& XLWorksheet::getImageInfoByImageID(const std::string& imageId) const
    {
        // Auto-refresh registry if empty
        if (m_imageRegistry.empty()) {
            refreshImageRegistry();
        }

        // Linear search for the image ID
        for (const auto& imgInfo : m_imageRegistry) {
            if (imgInfo.imageId == imageId) {
                return imgInfo;
            }
        }

        // Return empty image info if not found
        return m_emptyImageInfo;
    }

    /**
     * @brief Get specific image information by its relationship ID
     * @param relationshipId The relationship ID in drawing.xml.rels
     * @return Const reference to XLImageInfo object, or empty XLImageInfo if not found
     */
    const XLImageInfo& XLWorksheet::getImageInfoByRelationshipId(const std::string& relationshipId) const
    {
        // Auto-refresh registry if empty
        if (m_imageRegistry.empty()) {
            refreshImageRegistry();
        }

        // Linear search for the relationship ID
        for (const auto& imgInfo : m_imageRegistry) {
            if (imgInfo.relationshipId == relationshipId) {
                return imgInfo;
            }
        }

        // Return empty image info if not found
        return m_emptyImageInfo;
    }

    /**
     * @brief Get all image information at a specific cell
     * @param cellRef The cell reference (e.g., "A1", "B5")
     * @return Const reference to vector of XLImageInfo objects at that cell
     */
    const std::vector<XLImageInfo>& XLWorksheet::getImageInfosAtCell(const std::string& cellRef) const
    {
        // Auto-refresh registry if empty
        if (m_imageRegistry.empty()) {
            refreshImageRegistry();
        }

        // Clear the result vector
        m_emptyImageInfoVector.clear();

        // Filter images by cell reference
        for (const auto& imgInfo : m_imageRegistry) {
            if (imgInfo.anchorCell == cellRef) {
                m_emptyImageInfoVector.push_back(imgInfo);
            }
        }

        return m_emptyImageInfoVector;
    }

    /**
     * @brief Get all image information in a specific range
     * @param cellRange The cell range (e.g., "A1:B5", "C10:D20")
     * @return Const reference to vector of XLImageInfo objects in that range
     */
    const std::vector<XLImageInfo>& XLWorksheet::getImageInfosInRange(const std::string& cellRange) const
    {
        // Auto-refresh registry if empty
        if (m_imageRegistry.empty()) {
            refreshImageRegistry();
        }

        // Clear the result vector
        m_emptyImageInfoVector.clear();

        // Parse the range
        size_t colonPos = cellRange.find(':');
        if (colonPos == std::string::npos) {
            // Single cell range
            return getImageInfosAtCell(cellRange);
        }

        std::string topLeft = cellRange.substr(0, colonPos);
        std::string bottomRight = cellRange.substr(colonPos + 1);

        try {
            XLCellReference topLeftRef(topLeft);
            XLCellReference bottomRightRef(bottomRight);

            // Filter images by range
            for (const auto& imgInfo : m_imageRegistry) {
                if (!imgInfo.anchorCell.empty()) {
                    XLCellReference imgRef(imgInfo.anchorCell);
                    
                    // Check if image is within range
                    if (imgRef.row() >= topLeftRef.row() && imgRef.row() <= bottomRightRef.row() &&
                        imgRef.column() >= topLeftRef.column() && imgRef.column() <= bottomRightRef.column()) {
                        m_emptyImageInfoVector.push_back(imgInfo);
                    }
                }
            }
        } catch (const std::exception&) {
            // Invalid range format, return empty vector
        }

        return m_emptyImageInfoVector;
    }

    /**
     * @brief Get iterator to the beginning of the image information collection
     * @return Const iterator to the first image info
     */
    std::vector<XLImageInfo>::const_iterator XLWorksheet::imageInfosBegin() const
    {
        // Auto-refresh registry if empty
        if (m_imageRegistry.empty()) {
            refreshImageRegistry();
        }
        return m_imageRegistry.begin();
    }

    /**
     * @brief Get iterator to the end of the image information collection
     * @return Const iterator to the end of the image information collection
     */
    std::vector<XLImageInfo>::const_iterator XLWorksheet::imageInfosEnd() const
    {
        // Auto-refresh registry if empty
        if (m_imageRegistry.empty()) {
            refreshImageRegistry();
        }
        return m_imageRegistry.end();
    }

    //----------------------------------------------------------------------------------------------------------------------
    //           Image Modification Methods Implementation
    //----------------------------------------------------------------------------------------------------------------------

    /**
     * @brief Remove an image by its ID
     * @param imageId The unique image identifier to remove
     * @return True if the image was found and removed, false otherwise
     */
    bool XLWorksheet::removeImageByImageID(const std::string& imageId)
    {
        // Auto-refresh registry if empty
        if (m_imageRegistry.empty()) {
            refreshImageRegistry();
        }

        // Find the image in the registry
        auto it = std::find_if(m_imageRegistry.begin(), m_imageRegistry.end(),
            [&imageId](const XLImageInfo& img) { return img.imageId == imageId; });

        if (it == m_imageRegistry.end()) {
            return false; // Image not found
        }

        // Store the relationship ID before removing from registry
        std::string relationshipId = it->relationshipId;

        // Get image path BEFORE removing relationships (we need the relationship to find the file path)
        std::string imagePath = getImagePathFromRelationship(relationshipId);
        
        // Remove from DrawingML XML (in-memory XML document)
        bool xmlRemoved = removeImageFromDrawingXML(relationshipId);
        
        // Remove from relationships file (in-memory XML document)
        bool relsRemoved = removeImageFromRelationships(relationshipId);
        
        // Remove image file from archive (delete binary file entry) - use the path we got earlier
        bool fileRemoved = removeImageFileByPath(imagePath);

        // Remove from registry (in-memory cache)
        m_imageRegistry.erase(it);

        // Return true if at least XML removal succeeded
        return xmlRemoved;
    }

    /**
     * @brief Remove an image by its relationship ID
     * @param relationshipId The relationship ID to remove
     * @return True if the image was found and removed, false otherwise
     */
    bool XLWorksheet::removeImageByRelationshipId(const std::string& relationshipId)
    {
        // Auto-refresh registry if empty
        if (m_imageRegistry.empty()) {
            refreshImageRegistry();
        }

        // Find the image in the registry
        auto it = std::find_if(m_imageRegistry.begin(), m_imageRegistry.end(),
            [&relationshipId](const XLImageInfo& img) { return img.relationshipId == relationshipId; });

        if (it == m_imageRegistry.end()) {
            return false; // Image not found
        }

        // Remove from DrawingML XML (in-memory XML document)
        bool xmlRemoved = removeImageFromDrawingXML(relationshipId);
        
        // Remove from relationships file (in-memory XML document)
        bool relsRemoved = removeImageFromRelationships(relationshipId);
        
        // Remove image file from archive (delete binary file entry)
        bool fileRemoved = removeImageFileFromArchive(relationshipId);

        // Remove from registry (in-memory cache)
        m_imageRegistry.erase(it);

        // Return true if at least XML removal succeeded
        return xmlRemoved;
    }

    /**
     * @brief Remove all images from the worksheet
     */
    void XLWorksheet::clearImages()
    {
        // Clear the registry
        m_imageRegistry.clear();

        // TODO: Implement XML clearing from DrawingML
        // This would involve:
        // 1. Clearing all image nodes from drawing.xml
        // 2. Clearing the relationships file
        // 3. Removing all image files from the archive
    }

    //----------------------------------------------------------------------------------------------------------------------
    //           Image Registry Management Implementation
    //----------------------------------------------------------------------------------------------------------------------

    /**
     * @brief Refresh the image registry from DrawingML XML
     */
    void XLWorksheet::refreshImageRegistry() const
    {
        // Clear existing registry
        m_imageRegistry.clear();

        // Check if DrawingML is valid
        if (!m_drawingML.valid()) {
            return; // No drawing data available
        }

        try {
            // Get the DrawingML XML content by reading directly from the archive
            std::string drawingFilename = "xl/drawings/drawing" + std::to_string(sheetXmlNumber()) + ".xml";
            
            std::string xmlContent = parentDoc().readRelationshipsFile(drawingFilename);
            if (xmlContent.empty()) {
                return; // No XML data available
            }

            // Parse the XML document
            pugi::xml_document doc;
            if (!doc.load_string(xmlContent.c_str())) {
                return; // Failed to parse XML
            }

            // Get the root element (xdr:wsDr)
            auto wsDr = doc.document_element();
            std::string rootName = wsDr.name();
            
            // Check for both "wsDr" and "xdr:wsDr" (with namespace prefix)
            if (rootName != "wsDr" && rootName != "xdr:wsDr") {
                return; // Invalid root element
            }

            // Read the relationships file to get image mappings
            std::map<std::string, std::string> relationshipMap;
            readDrawingRelationships(relationshipMap);

            // Process each anchor element
            for (auto anchor : wsDr.children()) {
                std::string anchorName = anchor.name();
                
                if (matchesElementName(anchorName, "oneCellAnchor")) {
                    processOneCellAnchor(anchor, relationshipMap);
                }
                else if (matchesElementName(anchorName, "twoCellAnchor")) {
                    processTwoCellAnchor(anchor, relationshipMap);
                }
            }
        }
        catch (const std::exception& ) {
            // If parsing fails, leave registry empty
            m_imageRegistry.clear();
        }
    }

    /**
     * @brief Read the drawing relationships file to map relationship IDs to image files
     * @param relationshipMap Output map of relationship ID to image file path
     */
    void XLWorksheet::readDrawingRelationships(std::map<std::string, std::string>& relationshipMap) const
    {
        try {
            // Get the drawing filename
            std::string drawingFilename = "xl/drawings/drawing" + std::to_string(sheetXmlNumber()) + ".xml";
            std::string relsFilename = "xl/drawings/_rels/drawing" + std::to_string(sheetXmlNumber()) + ".xml.rels";

            // Check if relationships file exists
            if (!parentDoc().hasRelationshipsFile(relsFilename)) {
                return; // No relationships file
            }

            // Read the relationships file
            std::string relsContent = parentDoc().readRelationshipsFile(relsFilename);
            if (relsContent.empty()) {
                return; // Empty relationships file
            }

            // Parse the relationships XML
            pugi::xml_document relsDoc;
            if (!relsDoc.load_string(relsContent.c_str())) {
                return; // Failed to parse relationships
            }

            // Extract relationship mappings
            for (auto rel : relsDoc.document_element().children("Relationship")) {
                std::string id = rel.attribute("Id").value();
                std::string target = rel.attribute("Target").value();
                std::string type = rel.attribute("Type").value();

                // Only process image relationships
                if (type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image") {
                    relationshipMap[id] = target;
                }
            }
        }
        catch (const std::exception& ) {
            // If reading relationships fails, leave map empty
        }
    }

    /**
     * @brief Process a oneCellAnchor element and add to registry
     * @param anchor The oneCellAnchor XML node
     * @param relationshipMap Map of relationship IDs to image files
     */
    void XLWorksheet::processOneCellAnchor(const pugi::xml_node& anchor, const std::map<std::string, std::string>& relationshipMap) const
    {
        try {
            // Extract cell position
            auto from = anchor.child("xdr:from");
            if (from.empty()) {
                from = anchor.child("from"); // Try without namespace prefix
            }
            if (from.empty()) {
                return;
            }

            auto colNode = from.child("xdr:col");
            if (colNode.empty()) {
                colNode = from.child("col"); // Try without namespace prefix
            }
            auto rowNode = from.child("xdr:row");
            if (rowNode.empty()) {
                rowNode = from.child("row"); // Try without namespace prefix
            }
            
            int col = colNode.text().as_int();
            int row = rowNode.text().as_int();
            std::string cellRef = XLCellReference(row + 1, col + 1).address(); // Convert to 1-based

            // Extract dimensions
            auto ext = anchor.child("xdr:ext");
            if (ext.empty()) {
                ext = anchor.child("ext"); // Try without namespace prefix
            }
            if (ext.empty()) {
                return;
            }

            uint32_t widthEMUs = ext.attribute("cx").as_uint();
            uint32_t heightEMUs = ext.attribute("cy").as_uint();

            // Extract image information
            auto pic = anchor.child("xdr:pic");
            if (pic.empty()) {
                pic = anchor.child("pic"); // Try without namespace prefix
            }
            if (pic.empty()) {
                return;
            }

            auto blipFill = pic.child("xdr:blipFill");
            if (blipFill.empty()) {
                blipFill = pic.child("a:blipFill"); // Try with a namespace prefix
            }
            if (blipFill.empty()) {
                blipFill = pic.child("blipFill"); // Try without namespace prefix
            }
            if (blipFill.empty()) {
                return;
            }

            auto blip = blipFill.child("a:blip");
            if (blip.empty()) {
                blip = blipFill.child("blip"); // Try without namespace prefix
            }
            if (blip.empty()) {
                return;
            }

            std::string relationshipId = blip.attribute("r:embed").value();
            if (relationshipId.empty()) {
                return;
            }

            // Get image file path from relationship
            auto it = relationshipMap.find(relationshipId);
            if (it == relationshipMap.end()) {
                return;
            }

            std::string imagePath = it->second;
            
            // Extract image ID from filename (e.g., "../media/image_img1.png" -> "img1")
            std::string imageId = extractImageIdFromPath(imagePath);

            // Create XLImageInfo
            XLImageInfo imgInfo;
            imgInfo.imageId = imageId;
            imgInfo.relationshipId = relationshipId;
            imgInfo.anchorCell = cellRef;
            imgInfo.anchorType = "oneCellAnchor";
            imgInfo.displayWidthEMUs = widthEMUs;
            imgInfo.displayHeightEMUs = heightEMUs;

            // Convert EMUs to pixels (approximate)
            imgInfo.widthPixels = emusToPixels(widthEMUs);
            imgInfo.heightPixels = emusToPixels(heightEMUs);

            // Add to registry
            m_imageRegistry.push_back(imgInfo);
        }
        catch (const std::exception& ) {
            // If processing fails, skip this anchor
        }
    }

    /**
     * @brief Process a twoCellAnchor element and add to registry
     * @param anchor The twoCellAnchor XML node
     * @param relationshipMap Map of relationship IDs to image files
     */
    void XLWorksheet::processTwoCellAnchor(const pugi::xml_node& anchor, const std::map<std::string, std::string>& relationshipMap) const
    {
        try {
            // Extract cell position (use 'from' cell as primary anchor)
            auto from = anchor.child("xdr:from");
            if (from.empty()) {
                from = anchor.child("from"); // Try without namespace prefix
            }
            if (from.empty()) {
                return;
            }

            auto colNode = from.child("xdr:col");
            if (colNode.empty()) {
                colNode = from.child("col"); // Try without namespace prefix
            }
            auto rowNode = from.child("xdr:row");
            if (rowNode.empty()) {
                rowNode = from.child("row"); // Try without namespace prefix
            }
            
            int col = colNode.text().as_int();
            int row = rowNode.text().as_int();
            std::string cellRef = XLCellReference(row + 1, col + 1).address(); // Convert to 1-based

            // Extract dimensions
            auto ext = anchor.child("xdr:ext");
            if (ext.empty()) {
                ext = anchor.child("ext"); // Try without namespace prefix
            }
            if (ext.empty()) {
                return;
            }

            uint32_t widthEMUs = ext.attribute("cx").as_uint();
            uint32_t heightEMUs = ext.attribute("cy").as_uint();

            // Extract image information
            auto pic = anchor.child("xdr:pic");
            if (pic.empty()) {
                pic = anchor.child("pic"); // Try without namespace prefix
            }
            if (pic.empty()) {
                return;
            }

            auto blipFill = pic.child("xdr:blipFill");
            if (blipFill.empty()) {
                blipFill = pic.child("a:blipFill"); // Try with a namespace prefix
            }
            if (blipFill.empty()) {
                blipFill = pic.child("blipFill"); // Try without namespace prefix
            }
            if (blipFill.empty()) {
                return;
            }

            auto blip = blipFill.child("a:blip");
            if (blip.empty()) {
                blip = blipFill.child("blip"); // Try without namespace prefix
            }
            if (blip.empty()) {
                return;
            }

            std::string relationshipId = blip.attribute("r:embed").value();
            if (relationshipId.empty()) {
                return;
            }

            // Get image file path from relationship
            auto it = relationshipMap.find(relationshipId);
            if (it == relationshipMap.end()) {
                return;
            }

            std::string imagePath = it->second;
            
            // Extract image ID from filename
            std::string imageId = extractImageIdFromPath(imagePath);

            // Create XLImageInfo
            XLImageInfo imgInfo;
            imgInfo.imageId = imageId;
            imgInfo.relationshipId = relationshipId;
            imgInfo.anchorCell = cellRef;
            imgInfo.anchorType = "twoCellAnchor";
            imgInfo.displayWidthEMUs = widthEMUs;
            imgInfo.displayHeightEMUs = heightEMUs;

            // Convert EMUs to pixels (approximate)
            imgInfo.widthPixels = emusToPixels(widthEMUs);
            imgInfo.heightPixels = emusToPixels(heightEMUs);

            // Add to registry
            m_imageRegistry.push_back(imgInfo);
        }
        catch (const std::exception& ) {
            // If processing fails, skip this anchor
        }
    }

    /**
     * @brief Extract image ID from image file path
     * @param imagePath The image file path (e.g., "../media/image_img1.png")
     * @return The image ID (e.g., "img1")
     */
    std::string XLWorksheet::extractImageIdFromPath(const std::string& imagePath) const
    {
        // Extract filename from path
        size_t lastSlash = imagePath.find_last_of('/');
        std::string filename = (lastSlash != std::string::npos) ? imagePath.substr(lastSlash + 1) : imagePath;

        // Remove "image_" prefix and file extension
        if (filename.substr(0, 6) == "image_") {
            std::string withoutPrefix = filename.substr(6);
            size_t dotPos = withoutPrefix.find_last_of('.');
            if (dotPos != std::string::npos) {
                return withoutPrefix.substr(0, dotPos);
            }
        }

        return filename; // Fallback to full filename
    }

    /**
     * @brief Convert EMUs to pixels (approximate conversion)
     * @param emus EMUs value
     * @return Approximate pixel value
     */
    uint32_t XLWorksheet::emusToPixels(uint32_t emus) const
    {
        // Standard conversion: 1 inch = 914400 EMUs, 1 inch = 96 pixels (at 96 DPI)
        // So: 1 pixel = 914400 / 96 = 9525 EMUs
        return static_cast<uint32_t>(emus / 9525);
    }

    /**
     * @brief Validate image registry for data integrity
     * @return True if registry is valid (no duplicate IDs), false otherwise
     */
    bool XLWorksheet::validateImageRegistry() const
    {
        std::set<std::string> imageIds;
        std::set<std::string> relationshipIds;

        for (const auto& imgInfo : m_imageRegistry) {
            // Check for duplicate image IDs
            if (!imgInfo.imageId.empty()) {
                if (imageIds.find(imgInfo.imageId) != imageIds.end()) {
                    return false; // Duplicate image ID found
                }
                imageIds.insert(imgInfo.imageId);
            }

            // Check for duplicate relationship IDs
            if (!imgInfo.relationshipId.empty()) {
                if (relationshipIds.find(imgInfo.relationshipId) != relationshipIds.end()) {
                    return false; // Duplicate relationship ID found
                }
                relationshipIds.insert(imgInfo.relationshipId);
            }
        }

        return true; // No duplicates found
    }

    /**
     * @brief Check if an image ID already exists
     * @param imageId The image ID to check
     * @return True if the ID exists, false otherwise
     */
    bool XLWorksheet::hasImageWithId(const std::string& imageId) const
    {
        // Auto-refresh registry if empty
        if (m_imageRegistry.empty()) {
            refreshImageRegistry();
        }

        return std::any_of(m_imageRegistry.begin(), m_imageRegistry.end(),
            [&imageId](const XLImageInfo& img) { return img.imageId == imageId; });
    }

    /**
     * @brief Check if a relationship ID already exists
     * @param relationshipId The relationship ID to check
     * @return True if the ID exists, false otherwise
     */
    bool XLWorksheet::hasImageWithRelationshipId(const std::string& relationshipId) const
    {
        // Auto-refresh registry if empty
        if (m_imageRegistry.empty()) {
            refreshImageRegistry();
        }

        return std::any_of(m_imageRegistry.begin(), m_imageRegistry.end(),
            [&relationshipId](const XLImageInfo& img) { return img.relationshipId == relationshipId; });
    }

    /**
     * @brief Get image path from relationship ID
     * @param relationshipId The relationship ID to look up
     * @return The image path, or empty string if not found
     */
    std::string XLWorksheet::getImagePathFromRelationship(const std::string& relationshipId) const
    {
        try {
            std::string relsFilename = "xl/drawings/_rels/drawing" + std::to_string(sheetXmlNumber()) + ".xml.rels";
            
            if (!parentDoc().hasRelationshipsFile(relsFilename)) {
                return "";
            }

            std::string relsContent = parentDoc().readRelationshipsFile(relsFilename);
            if (relsContent.empty()) {
                return "";
            }

            pugi::xml_document relsDoc;
            if (!relsDoc.load_string(relsContent.c_str())) {
                return "";
            }

            for (auto rel : relsDoc.document_element().children("Relationship")) {
                if (rel.attribute("Id").value() == relationshipId) {
                    return rel.attribute("Target").value();
                }
            }
            
            return "";
        }
        catch (const std::exception&) {
            return "";
        }
    }

    /**
     * @brief Remove image file by direct path
     * @param imagePath The relative image path (e.g., "../media/image_img3.png")
     * @return True if the file was found and removed
     */
    bool XLWorksheet::removeImageFileByPath(const std::string& imagePath) const
    {
        if (imagePath.empty()) {
            return false;
        }

        try {
            // Convert relative path to absolute path within the archive
            std::string fullImagePath = "xl/media/" + imagePath.substr(imagePath.find_last_of('/') + 1);
            
            bool result = const_cast<XLDocument&>(parentDoc()).deleteEntry(fullImagePath);
            return result;
        }
        catch (const std::exception&) {
            return false;
        }
    }

    /**
     * @brief Remove image node from DrawingML XML (in-memory document)
     * @param relationshipId The relationship ID of the image to remove
     * @return True if the image node was found and removed
     * 
     * @note This function modifies the existing XLXmlData object for the drawing file.
     *       The changes are automatically persisted when the document is saved.
     */
    bool XLWorksheet::removeImageFromDrawingXML(const std::string& relationshipId) const
    {
        try {
            // Use standard sequential numbering approach (consistent with rest of codebase)
            std::string drawingFilename = "xl/drawings/drawing" + std::to_string(sheetXmlNumber()) + ".xml";
            // Get the existing XLXmlData object for the drawing file
            // This is the proper way - use the existing XML data structure rather than bypassing it
            XLXmlData* xmlData = const_cast<XLDocument&>(parentDoc()).getDrawingXmlData(drawingFilename);
            if (!xmlData) {
                return false;
            }
            
            // Get the underlying XMLDocument (pugi::xml_document)
            XMLDocument& xmlDoc = *xmlData->getXmlDocument();
            if (!xmlDoc.document_element()) {
                return false;
            }

            // Get the root element (xdr:wsDr)
            auto wsDr = xmlDoc.document_element();
            std::string rootName = wsDr.name();
            
            // Check for both "wsDr" and "xdr:wsDr" (with namespace prefix)
            if (rootName != "wsDr" && rootName != "xdr:wsDr") {
                return false;
            }

            // Find and remove the image node
            for (auto anchor : wsDr.children()) {
                // Check both oneCellAnchor and twoCellAnchor
                if (matchesElementName(anchor.name(), "oneCellAnchor") || 
                    matchesElementName(anchor.name(), "twoCellAnchor")) {
                    
                    // Look for the pic element with the matching relationship ID
                    auto pic = anchor.child("xdr:pic");
                    if (pic.empty()) {
                        pic = anchor.child("pic");
                    }
                    
                    if (!pic.empty()) {
                        // Check blipFill -> blip -> r:embed
                        auto blipFill = pic.child("xdr:blipFill");
                        if (blipFill.empty()) {
                            blipFill = pic.child("a:blipFill");
                        }
                        if (blipFill.empty()) {
                            blipFill = pic.child("blipFill");
                        }
                        
                        if (!blipFill.empty()) {
                            auto blip = blipFill.child("a:blip");
                            if (blip.empty()) {
                                blip = blipFill.child("blip");
                            }
                            
                            if (!blip.empty()) {
                                std::string embedId = blip.attribute("r:embed").value();
                                if (embedId == relationshipId) {
                                    // Found the matching image, remove the entire anchor
                                    wsDr.remove_child(anchor);
                                    
                                    // The XMLDocument is automatically updated - no need to save/restore
                                    // The XLXmlData will persist the changes when the document is saved
                                    return true;
                                }
                            }
                        }
                    }
                }
            }
            
            return false; // Image not found
        }
        catch (const std::exception&) {
            return false;
        }
    }

    /**
     * @brief Remove image relationship from relationships file (in-memory document)
     * @param relationshipId The relationship ID to remove
     * @return True if the relationship was found and removed
     */
    bool XLWorksheet::removeImageFromRelationships(const std::string& relationshipId) const
    {
        try {
            // Use standard sequential numbering approach (consistent with rest of codebase)
            std::string relsFilename = "xl/drawings/_rels/drawing" + std::to_string(sheetXmlNumber()) + ".xml.rels";
            
            // Check if relationships file exists
            if (!parentDoc().hasRelationshipsFile(relsFilename)) {
                return false;
            }

            // Read the relationships file from in-memory document
            std::string relsContent = parentDoc().readRelationshipsFile(relsFilename);
            if (relsContent.empty()) {
                return false;
            }

            // Parse the relationships XML
            pugi::xml_document relsDoc;
            if (!relsDoc.load_string(relsContent.c_str())) {
                return false;
            }

            // Find and remove the relationship
            for (auto rel : relsDoc.document_element().children("Relationship")) {
                std::string relId = rel.attribute("Id").value();
                if (relId == relationshipId) {
                    relsDoc.document_element().remove_child(rel);
                    
                    // Update the in-memory XML document
                    std::ostringstream oss;
                    relsDoc.save(oss);
                    const_cast<XLDocument&>(parentDoc()).addRelationshipsFile(relsFilename, oss.str());
                    
                    return true;
                }
            }
            
            return false; // Relationship not found
        }
        catch (const std::exception&) {
            return false;
        }
    }

    /**
     * @brief Remove image file from archive (delete binary file entry)
     * @param relationshipId The relationship ID of the image file to remove
     * @return True if the file was found and removed from archive
     */
    bool XLWorksheet::removeImageFileFromArchive(const std::string& relationshipId) const
    {
        try {
            // Use standard sequential numbering approach (consistent with rest of codebase)
            std::string relsFilename = "xl/drawings/_rels/drawing" + std::to_string(sheetXmlNumber()) + ".xml.rels";
            
            if (!parentDoc().hasRelationshipsFile(relsFilename)) {
                return false;
            }

            std::string relsContent = parentDoc().readRelationshipsFile(relsFilename);
            if (relsContent.empty()) {
                return false;
            }

            pugi::xml_document relsDoc;
            if (!relsDoc.load_string(relsContent.c_str())) {
                return false;
            }

            std::string imagePath = "";
            for (auto rel : relsDoc.document_element().children("Relationship")) {
                std::string relId = rel.attribute("Id").value();
                if (relId == relationshipId) {
                    imagePath = rel.attribute("Target").value();
                    break;
                }
            }

            if (!imagePath.empty()) {
                // Convert relative path to absolute path within the archive
                std::string fullImagePath = "xl/media/" + imagePath.substr(imagePath.find_last_of('/') + 1);
                
                // Remove the image file from the archive
                // This deletes the binary file entry from the in-memory archive
                try {
                    bool result = const_cast<XLDocument&>(parentDoc()).deleteEntry(fullImagePath);
                    return result;
                } catch (const std::exception&) {
                    return false;
                }
            } else {
                return false;
            }
            
            return false;
        }
        catch (const std::exception&) {
            return false;
        }
    }

    /**
     * @brief Robust XML element name matching that handles namespace prefixes
     * @param elementName The actual element name (may include namespace prefix)
     * @param targetName The target element name to match
     * @return True if the element name matches the target (with or without namespace)
     * 
     * @example
     * Exact match: "oneCellAnchor" matches "oneCellAnchor"
     * Namespace prefix: "xdr:oneCellAnchor" matches "oneCellAnchor"
     * Different namespace: "a:oneCellAnchor" matches "oneCellAnchor"
     */
    bool XLWorksheet::matchesElementName(const std::string& elementName, const std::string& targetName) const
    {
        // Exact match
        if (elementName == targetName) {
            return true;
        }
        
        // Check if element name ends with target name (handles namespace prefixes)
        if (elementName.length() > targetName.length() + 1) {
            size_t pos = elementName.find_last_of(':');
            if (pos != std::string::npos && pos + 1 < elementName.length()) {
                std::string localName = elementName.substr(pos + 1);
                if (localName == targetName) {
                    return true;
                }
            }
        }
        
        // Check if element name starts with target name followed by colon
        if (elementName.length() == targetName.length() + 1 && elementName[targetName.length()] == ':') {
            if (elementName.substr(0, targetName.length()) == targetName) {
                return true;
            }
        }
        
        return false;
    }

} // namespace OpenXLSX

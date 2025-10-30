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
#include "XLImage.hpp"
#include "XLSheet.hpp"
#include "XLCellReference.hpp"
#include "XLDocument.hpp"
#include "XLDrawingML.hpp"

namespace OpenXLSX
{
    //----------------------------------------------------------------------------------------------------------------------
    //           Image Query Methods Implementation
    //----------------------------------------------------------------------------------------------------------------------

    /**
     * @brief Get specific embedded image by its image ID
     */
    const XLEmbeddedImage& XLWorksheet::getEmbImageByImageID(const std::string& imageId) const
    {
        const std::vector<XLEmbeddedImage>& embImages = getEmbImages();
        const XLEmbeddedImage *result = nullptr;
        std::vector<XLEmbeddedImage>::const_iterator embImgItr = embImages.begin();
        for( ; (embImages.end() != embImgItr) && (nullptr == result); ++embImgItr ){
            const XLEmbeddedImage& embImage = *embImgItr;
            if( embImage.id() == imageId ){
                result = &embImage;
            }
        }

#if (defined(_DEBUG) || !defined(NDEBUG))
        /* check for duplicate ID */
        for( ; embImages.end() != embImgItr; ++embImgItr ){
            const XLEmbeddedImage& embImage = *embImgItr;
            if( embImage.id() == imageId ){
                std::cerr << "WARNING: Worksheet '" << name() 
                    << "' has duplicate imageId = '" << imageId << std::endl;
            }
        }
#endif
        // Return empty embedded image if not found
        static const XLEmbeddedImage emptyEmbImage{};
        if( nullptr == result ){
            result = &emptyEmbImage;
            }
            
        return *result;
    }

    /**
     * @brief Get specific embedded image by its relationship ID
     */
    const XLEmbeddedImage& XLWorksheet::getEmbImageByRelationshipId(const std::string& relationshipId) const
    {
        const std::vector<XLEmbeddedImage>& embImages = getEmbImages();
        const XLEmbeddedImage *result = nullptr;
        std::vector<XLEmbeddedImage>::const_iterator embImgItr = embImages.begin();
        for( ; (embImages.end() != embImgItr) && (nullptr == result); ++embImgItr ){
            const XLEmbeddedImage& embImage = *embImgItr;
            if( embImage.relationshipId() == relationshipId ){
                result = &embImage;
            }
        }

#if (defined(_DEBUG) || !defined(NDEBUG))
        /* check for duplicate ID */
        for( ; embImages.end() != embImgItr; ++embImgItr ){
            const XLEmbeddedImage& embImage = *embImgItr;
            if( embImage.relationshipId() == relationshipId ){
                std::cerr << "WARNING: Worksheet '" << name() 
                    << "' has duplicate relationshipId = '" << relationshipId << std::endl;
            }
        }
#endif
        // Return empty embedded image if not found
        static const XLEmbeddedImage emptyEmbImage{};
        if( nullptr == result ){
            result = &emptyEmbImage;
            }
            
        return *result;
    }

    /**
     * @brief Get all embedded images at a specific cell
     */
    std::vector<XLEmbeddedImage> XLWorksheet::getEmbImagesAtCell(const std::string& cellRef) const
    {
        std::vector<XLEmbeddedImage> result;

        // Filter embedded images by cell reference
        for (const auto& embImage : getEmbImages()) {
            if (embImage.anchorCell() == cellRef) {
                result.push_back(embImage);
            }
        }

        return result;
    }

    /**
     * @brief Get all embedded images in a specific range
     */
    std::vector<XLEmbeddedImage> XLWorksheet::getEmbImagesInRange(const std::string& cellRange) const
    {
        std::vector<XLEmbeddedImage> result;

        try {
            // Parse the range
            size_t colonPos = cellRange.find(':');
            if (colonPos == std::string::npos) {
                // Single cell range
                return getEmbImagesAtCell(cellRange);
            } else {
                // Range format
                std::string topLeftStr = cellRange.substr(0, colonPos);
                std::string bottomRightStr = cellRange.substr(colonPos + 1);
                
                XLCellReference topLeftRef(topLeftStr);
                XLCellReference bottomRightRef(bottomRightStr);
                
                // Filter embedded images by range
                for (const auto& embImage : getEmbImages()) {
                    XLCellReference imgRef(embImage.anchorCell());
                    
                    // Check if image is within range
                    if (imgRef.row() >= topLeftRef.row() && imgRef.row() <= bottomRightRef.row() &&
                        imgRef.column() >= topLeftRef.column() && imgRef.column() <= bottomRightRef.column()) {
                        result.push_back(embImage);
                    }
                }
            }
        } catch (const std::exception&) {
            // Invalid range format, return empty vector
        }

        return result;
    }

    /**
     * @brief Get iterator to the end of the embedded images collection
     * @return Const iterator to the end of the embedded images collection
     */
    std::vector<XLEmbeddedImage>::const_iterator XLWorksheet::embImagesEnd() const
    {
        return getEmbImages().end();
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
        // Find the image in embedded images
        auto it = std::find_if(getEmbImages().begin(), getEmbImages().end(),
            [&imageId](const XLEmbeddedImage& img) { return img.id() == imageId; });

        if (it == getEmbImages().end()) {
            return false; // Image not found
        }

        // Store the relationship ID before removing
        std::string relationshipId = it->relationshipId();

        // Remove from DrawingML XML (in-memory XML document)
        bool xmlRemoved = removeImageFromDrawingXML(relationshipId);
        
        // Remove from relationships file (in-memory XML document)
        bool relsRemoved = removeImageFromRelationships(relationshipId);
        

        // Remove from m_embImages vector (in-memory cache)
        getEmbImages().erase(it);

        // Note: XLDocument.getImageManager()->prune() will delete binary image data files from archive as needed.

        // Return true if all critical operations succeeded
        return xmlRemoved && relsRemoved;
    }

    /**
     * @brief Remove an image by its relationship ID
     * @param relationshipId The relationship ID to remove
     * @return True if the image was found and removed, false otherwise
     */
    bool XLWorksheet::removeImageByRelationshipId(const std::string& relationshipId)
    {
        // Find the image in embedded images
        auto it = std::find_if(getEmbImages().begin(), getEmbImages().end(),
            [&relationshipId](const XLEmbeddedImage& img) { return img.relationshipId() == relationshipId; });

        if (it == getEmbImages().end()) {
            return false; // Image not found
        }

        // Remove from DrawingML XML (in-memory XML document)
        bool xmlRemoved = removeImageFromDrawingXML(relationshipId);
        
        // Remove from relationships file (in-memory XML document)
        bool relsRemoved = removeImageFromRelationships(relationshipId);
        

        // Remove from m_embImages vector (in-memory cache)
        getEmbImages().erase(it);

        // Note: XLDocument.getImageManager()->prune() will delete binary image data files from archive as needed.

        // Return true if all critical operations succeeded
        return xmlRemoved && relsRemoved;
    }

    /**
     * @brief Remove all images from the worksheet
     * @return Number of images removed
     */
    void XLWorksheet::clearImages()
    {
        // Ensure m_embImages vector is cleared (safety check)
        getEmbImages().clear();
        
        // TODO: Add verification checks to ensure complete cleanup:
        // 1. Check that DrawingML XML has no image nodes
        // 2. Check that relationships file has no image relationships
        // 3. Check that no image files remain in the archive
        // 5. Verify that getImageIDs() returns empty vector
        // 6. Verify that imageCount() returns 0
        
        return;
    }

    //----------------------------------------------------------------------------------------------------------------------
    //           Image Registry Management Implementation
    //----------------------------------------------------------------------------------------------------------------------

    /**
     * @brief Read the drawing relationships file to map relationship IDs to image files
     * @param relationshipMap Output map of relationship ID to image file path
     */
    void XLWorksheet::readDrawingRelationships(std::map<std::string, std::string>& relationshipMap) const
    {
        try {
            // Get the drawing filename
            std::string drawingFilename = getDrawingFilename();
            std::string relsFilename = getDrawingRelsFilename();

            // Check if relationships file exists
            if (!parentDoc().hasRelationshipsFile(relsFilename)) {
                return; // No relationships file
            }

            // Alternate implementation using XLXmlData access
            // Get relationships XML data from in-memory cache
            XLXmlData* relsXmlData = const_cast<XLDocument&>(parentDoc()).getRelationshipsXmlData(relsFilename);
            if (!relsXmlData || !relsXmlData->valid()) {
                return; // No relationships data available
            }

            // Parse the relationships XML from in-memory data
            XMLDocument* relsDoc = relsXmlData->getXmlDocument();
            if (!relsDoc || !relsDoc->document_element()) {
                return; // Invalid XML data
            }

            // Extract relationship mappings
            int relationshipCount = 0;
            for (auto rel : relsDoc->document_element().children("Relationship")) {
                std::string id = rel.attribute("Id").value();
                std::string target = rel.attribute("Target").value();
                std::string type = rel.attribute("Type").value();

                // Only process image relationships
                if (type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image") {
                    relationshipMap[id] = target;
                    relationshipCount++;
                }
            }
        }
        catch (const std::exception& ) {
            // If reading relationships fails, leave map empty
        }
    }

    /**
     * @brief Populate m_embImages vector from existing XML data
     */
    void XLWorksheet::populateImagesFromXML()
    {
        // Only populate if m_embImages is empty to avoid duplicates
        if (!getEmbImages().empty()) {
            return;
        }

        // Check if DrawingML is valid (was initialized because images exist)
        if (!m_drawingML.valid()) {
            return; // No images to populate
        }

        try {
            // Use standard sequential numbering approach (consistent with rest of codebase)
            std::string drawingFilename = getDrawingFilename();
            // Get the existing XLXmlData object for the drawing file
            // This is the proper way - use the existing XML data structure rather than bypassing it
            XLXmlData* xmlData = const_cast<XLDocument&>(parentDoc()).getDrawingXmlData(drawingFilename);
            if (!xmlData) {
                return;
            }

            XMLNode wsDr;
            std::string rootName;
            // Get the underlying XMLDocument (pugi::xml_document)
            XMLDocument& xmlDoc = *xmlData->getXmlDocument();
            if (!xmlDoc.document_element()) {
                return;
            }

            // Get the root element (xdr:wsDr)
            wsDr = xmlDoc.document_element();
            rootName = wsDr.name();
            
            // Check for both "wsDr" and "xdr:wsDr" (with namespace prefix)
            if (rootName != "wsDr" && rootName != "xdr:wsDr") {
                return; // Invalid root element
            }

            // Read the relationships file to get image mappings
            std::map<std::string, std::string> relationshipMap;    
            readDrawingRelationships(relationshipMap);

            // Process each anchor element directly
            for (auto anchor : wsDr.children()) {
                std::string anchorName = anchor.name();
                // Check anchor type using enum conversion
                XLImageAnchorType anchorType = XLImageAnchorUtils::stringToAnchorType(anchorName);
                if (anchorType != XLImageAnchorType::Unknown) {
                    // Parse anchor using XLDrawingML utility function
                    XLImageAnchor anchorInfo;
                    if (XLDrawingML::parseXMLNode(anchor, &anchorInfo)) {
                        // Process this anchor directly into XLEmbeddedImage
                        processAnchorToEmbeddedImage(anchorInfo, relationshipMap);
                    }
                }
            }
        }
        catch (const std::exception&) {
            // If any error occurs, just return without populating
            return;
        }
    }

    /**
     * @brief Process a single XLImageAnchor into an XLEmbeddedImage and add to m_embImages
     * @param anchorInfo The parsed anchor information
     * @param relationshipMap Map of relationship IDs to image files
     */
    void XLWorksheet::processAnchorToEmbeddedImage(const XLImageAnchor& anchorInfo, const std::map<std::string, std::string>& relationshipMap)
    {
        try {
            // Find the image file path using the relationship ID
            auto it = relationshipMap.find(anchorInfo.relationshipId);
            if (it == relationshipMap.end()) {
                return; // Skip if relationship not found
            }

            std::string packagePath = it->second;
            
            // Convert relative path to absolute path if needed
            packagePath = getAbsoluteImagePath(packagePath);

            // Load binary data from archive
            std::string binaryData = parentDoc().readFile(packagePath);
            if (binaryData.empty()) {
                return; // Skip if binary data not found
            }
            
            // Convert string to vector<uint8_t>
            std::vector<uint8_t> imageData(binaryData.begin(), binaryData.end());
            
            // Determine MIME type from file extension ot imageData
            std::string extension = packagePath.substr(packagePath.find_last_of('.'));
            XLMimeType mimeType = XLImageUtils::mimeTypeFromExtension(extension);
            if (XLMimeType::Unknown == mimeType) {
                mimeType = XLImageUtils::detectMimeTypeEnum(imageData);
            }
        
            // Convert enum back to XLContentType for findOrAddImage
            XLContentType contentType = XLImageUtils::mimeTypeToContentType(mimeType);
           
            // Use XLImageManager to find or add the image with existing package path
            XLImageShPtr image = parentDoc().getImageManager().findOrAddImage(packagePath,
                imageData, contentType);
        
            if (!image.get()) {
                return;
            }

            // Add to embedded image register
            XLEmbeddedImage embImage;
            embImage.setImage(image);
            embImage.setImageAnchor( anchorInfo );
            getEmbImages().push_back(embImage);
        }
        catch (const std::exception&) {
            // If creating XLEmbeddedImage fails, skip this entry
            // This ensures we don't crash on malformed data
        }
    }

    /**
     * @brief Extract image ID from XML pic element
     * @param pic The pic XML node containing the image information
     * @return The image ID from the cNvPr id attribute
     */
    std::string XLWorksheet::extractImageIdFromXML(const pugi::xml_node& pic) const
    {
        try {
            // Navigate to cNvPr element: pic -> nvPicPr -> cNvPr
            auto nvPicPr = pic.child("xdr:nvPicPr");
            if (nvPicPr.empty()) {
                nvPicPr = pic.child("nvPicPr"); // Try without namespace
            }
            
            if (!nvPicPr.empty()) {
                auto cNvPr = nvPicPr.child("xdr:cNvPr");
                if (cNvPr.empty()) {
                    cNvPr = nvPicPr.child("cNvPr"); // Try without namespace
                }
                
                   if (!cNvPr.empty()) {
                       auto idAttr = cNvPr.attribute("id");
                       if (!idAttr.empty()) {
                           std::string imageId = idAttr.value();
                           return imageId;
                       }
                   }
            }
        } catch (const std::exception&) {
            // If XML parsing fails, fall back to empty string
        }
        
        return ""; // Fallback to empty string
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

        for (const auto& embImage : getEmbImages()) {
            // Check for duplicate image IDs
            if (!embImage.id().empty()) {
                if (imageIds.find(embImage.id()) != imageIds.end()) {
                    return false; // Duplicate image ID found
                }
                imageIds.insert(embImage.id());
            }

            // Check for duplicate relationship IDs
            if (!embImage.relationshipId().empty()) {
                if (relationshipIds.find(embImage.relationshipId()) != relationshipIds.end()) {
                    return false; // Duplicate relationship ID found
                }
                relationshipIds.insert(embImage.relationshipId());
            }
        }

        return true; // No duplicates found
    }

    /**
     * @brief Check if an image ID already exists
     * @param imageId The image ID to check
     * @return True if the ID exists, false otherwise
     */
    bool XLWorksheet::hasEmbImageWithId(const std::string& imageId) const
    {
        return std::any_of(getEmbImages().begin(), getEmbImages().end(),
            [&imageId](const XLEmbeddedImage& img) { return img.id() == imageId; });
    }

    /**
     * @brief Check if a relationship ID already exists
     * @param relationshipId The relationship ID to check
     * @return True if the ID exists, false otherwise
     */
    bool XLWorksheet::hasImageWithRelationshipId(const std::string& relationshipId) const
    {
        return std::any_of(getEmbImages().begin(), getEmbImages().end(),
            [&relationshipId](const XLEmbeddedImage& img) { return img.relationshipId() == relationshipId; });
    }

    /**
     * @brief Get image path from relationship ID
     * @param relationshipId The relationship ID to look up
     * @return The image path, or empty string if not found
     */
    std::string XLWorksheet::getImagePathFromRelationship(const std::string& relationshipId) const
    {
        try {
            std::string relsFilename = getDrawingRelsFilename();
            
            if (!parentDoc().hasRelationshipsFile(relsFilename)) {
                return "";
            }
            // Alternate implementation using XLXmlData access
            // Get relationships XML data from in-memory cache
            XLXmlData* relsXmlData = const_cast<XLDocument&>(parentDoc()).getRelationshipsXmlData(relsFilename);
            if (!relsXmlData || !relsXmlData->valid()) {
            return "";
            }

            // Parse the relationships XML from in-memory data
            XMLDocument* relsDoc = relsXmlData->getXmlDocument();
            if (!relsDoc || !relsDoc->document_element()) {
                return "";
            }

            // Find the relationship with the specified ID
            for (auto rel : relsDoc->document_element().children("Relationship")) {
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
 * @brief Add an embedded image, assuming image is fully created and imageAnchor has all user-specified data
 * @details This function generates unique IDs, completes the anchor initialization, and adds the image to the worksheet
 * @param imageAnchor Anchor configuration (may be partially initialized)
 * @param image Shared pointer to the XLImage object
 * @return Image ID string
 */
std::string XLWorksheet::embedImage(const XLImageAnchor& imageAnchor, XLImageShPtr image)
{
    std::string imageId;
    if (image.get()) {
        // Get unique IDs
        imageId = generateNextImageId();
        const std::string relationshipId = getNextUnusedRelationshipId();
        
        // Fully initialize image anchor
        XLImageAnchor anchor = imageAnchor;
        anchor.imageId = imageId;
        anchor.relationshipId = relationshipId;
        if (XLImageAnchorType::Unknown == anchor.anchorType) {
            anchor.anchorType = ((0 == anchor.toCol) && (0 == anchor.toRow)) ?
                               XLImageAnchorType::OneCellAnchor : XLImageAnchorType::TwoCellAnchor;
        }
        if ((0 == anchor.displayWidthEMUs) || (0 == anchor.displayHeightEMUs)) {
            const uint32_t widthPixels = image->widthPixels();
            const uint32_t heightPixels = image->heightPixels();
            anchor.displayWidthEMUs = XLImageUtils::pixelsToEMUs(widthPixels);
            anchor.displayHeightEMUs = XLImageUtils::pixelsToEMUs(heightPixels);
        }
        if ((0 == anchor.actualWidthEMUs) || (0 == anchor.actualHeightEMUs)) {
            anchor.actualWidthEMUs = anchor.displayWidthEMUs;
            anchor.actualHeightEMUs = anchor.displayHeightEMUs;
        }
        
        // Construct embedded image
        XLEmbeddedImage embImage;
        embImage.setImageAnchor(anchor);
        embImage.setImage(image);
        
        // Add anchor to drawing file XML (drawing3.xml)
        if (!m_drawingML.valid()) {
            createEmptyDrawingML();
        }
        m_drawingML.addImage(embImage);
        
        // Add relationship to relationship file XML (drawing3.xml.rels)
        addImageRelationship(embImage);
        
        // Add to register
        getEmbImages().push_back(embImage);
    }
    return imageId;
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
            std::string drawingFilename = getDrawingFilename();
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

            // Delegate to XLDrawingML utility function
            return XLDrawingML::deleteXMLNode(wsDr, relationshipId);
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
            std::string relsFilename = getDrawingRelsFilename();
            
            // Check if relationships file exists
            if (!parentDoc().hasRelationshipsFile(relsFilename)) {
                return false;
            }

            // Alternate implementation using XLXmlData access
            // Get relationships XML data from in-memory cache
            XLXmlData* relsXmlData = const_cast<XLDocument&>(parentDoc()).getRelationshipsXmlData(relsFilename);
            if (!relsXmlData || !relsXmlData->valid()) {
                return false;
            }

            // Parse the relationships XML from in-memory data
            XMLDocument* relsDoc = relsXmlData->getXmlDocument();
            if (!relsDoc || !relsDoc->document_element()) {
                return false;
            }

            // Find and remove the relationship
            bool found = false;
            for (auto rel : relsDoc->document_element().children("Relationship")) {
                std::string relId = rel.attribute("Id").value();
                if (relId == relationshipId) {
                    relsDoc->document_element().remove_child(rel);
                    found = true;
                    break;
                }
            }
            
            if (found) {
                // Persist the changes to XLXmlData using setRawData()
                // This avoids archive access and lets XLXmlData handle persistence
                std::ostringstream oss;
                relsDoc->save(oss);
                relsXmlData->setRawData(oss.str());
                return true;
            }
            
            return false; // Relationship not found
        }
        catch (const std::exception&) {
            return false;
        }
    }

    /**
     * @brief Check if a binary file is referenced by any relationships in this worksheet
     * @param imagePath The relative image path (e.g., "../media/image_3.gif")
     * @return True if the file is referenced by at least one relationship in this worksheet
*     TODO: imagePath should be imagePackagePath
     */
    bool XLWorksheet::isBinaryFileReferenced(const std::string& imagePath) const
    {
        try {
            // Get the drawing relationships filename
            std::string relsFilename = getDrawingRelsFilename();
            
            // Check if relationships file exists in archive
            if (!parentDoc().hasRelationshipsFile(relsFilename)) {
                return false; // No relationships file
            }

            // Alternate implementation using XLXmlData access
            // Get relationships XML data from in-memory cache
            XLXmlData* relsXmlData = const_cast<XLDocument&>(parentDoc()).getRelationshipsXmlData(relsFilename);
            if (!relsXmlData || !relsXmlData->valid()) {
                return false; // No relationships data available
            }

            // Parse the relationships XML from in-memory data
            XMLDocument* relsDoc = relsXmlData->getXmlDocument();
            if (!relsDoc || !relsDoc->document_element()) {
                return false; // Invalid XML data
            }

            // Check if any relationship references the image path
            for (auto rel : relsDoc->document_element().children("Relationship")) {
                std::string target = rel.attribute("Target").value();
                if (target == imagePath) {
                    return true; // Found at least one reference
                }
            }
            
            return false; // No references found
            
        } catch (const std::exception&) {
            // If any error occurs, assume no references to be safe
            return false;
        }
    }

} // namespace OpenXLSX

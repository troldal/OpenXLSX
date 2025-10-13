/*

   ____                               ____      ___ ____       ____  ____      ___
  6MMMMb                              `MM(      )M' `MM'      6MMMMb\`MM(      )M'
 8P    Y8                              `MM.     d'   MM      6M'    ` `MM.     d'
6M      Mb __ ____     ____  ___  __    `MM.   d'    MM      MM        `MM.   d'
MM      MM `M6MMMMb   6MMMMb `MM 6MMb    `MM. d'     MM      YM.        `MM. d'
MM      MM  MM    MM MM    MM MM'   MM     `MMd       MM       YMMMMb     `MMd
MM      MM  MM    MM MMMMMMMM MM    MM    d'`MM.     MM            `Mb     dMM.
YM      M9  MM    MM MM       MM    MM   d'  `MM.    MM            MM    d'`MM.
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
  DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY
  DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
  (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
  LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
  ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
  (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
  SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

 */

#ifndef OPENXLSX_XLIMAGE_HPP
#define OPENXLSX_XLIMAGE_HPP

// ===== External Includes ===== //
#include <cstdint>    // uint8_t, uint16_t, uint32_t
#include <string>     // std::string
#include <vector>     // std::vector

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLContentTypes.hpp"

/*
 * ================================================================================
 * IMAGE EMBEDDING RELATIONSHIP CREATION IN OPENXLSX
 * ================================================================================
 * 
 * When images are embedded in Excel worksheets, OpenXLSX creates TWO types of
 * relationships that work together to link worksheets → drawings → images.
 * Each relationship type uses different ID management strategies appropriate
 * to their scope and purpose.
 * 
 * ┌─────────────────────────────────────────────────────────────────────────────┐
 * │  RELATIONSHIP CREATION FLOW                                                 │
 * ├─────────────────────────────────────────────────────────────────────────────┤
 * │                                                                             │
 * │  1. Worksheet.addImage() called                                            │
 * │     ↓                                                                       │
 * │  2. drawingML() ensures drawing.xml exists                                  │
 * │     ↓                                                                       │
 * │  3. Creates Worksheet → Drawing relationship                                │
 * │     ↓                                                                       │
 * │  4. addImageToDrawingML() adds image to drawing.xml                        │
 * │     ↓                                                                       │
 * │  5. createDrawingRelationshipsFile() creates Drawing → Image relationships │
 * │                                                                             │
 * └─────────────────────────────────────────────────────────────────────────────┘
 * 
 * RELATIONSHIP TYPE 1: WORKSHEET-LEVEL RELATIONSHIPS
 * ──────────────────────────────────────────────────────────────────────────────
 * 
 * Location: Lines 1735-1750 in XLWorksheet::drawingML()
 * 
 * Code:
 *   std::string drawingRelativePath = getPathARelativeToPathB(drawingFilename, getXmlPath());
 *   XLRelationshipItem drawingRelationship;
 *   if (!m_relationships.targetExists(drawingRelativePath))
 *       drawingRelationship = m_relationships.addRelationship(XLRelationshipType::Drawing, drawingRelativePath);
 *   else
 *       drawingRelationship = m_relationships.relationshipByTarget(drawingRelativePath);
 *   
 *   XMLNode drawing = appendAndGetNode(docElement, "drawing", m_nodeOrder);
 *   appendAndSetAttribute(drawing, "r:id", drawingRelationship.id());
 * 
 * Creates:
 *   File: xl/worksheets/_rels/sheet1.xml.rels
 *   Content: <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
 * 
 * Purpose: Links worksheet to drawing.xml file
 * Uses: XLRelationships::addRelationship() with intelligent ID management
 * 
 * ID Management Strategy:
 *   - Uses GetNewRelsID() function with intelligent scanning
 *   - Checks existing relationships to avoid conflicts
 *   - Handles complex scenarios with multiple relationship types
 *   - Ensures uniqueness within the worksheet's relationship file
 * 
 * RELATIONSHIP TYPE 2: DRAWING-LEVEL RELATIONSHIPS
 * ──────────────────────────────────────────────────────────────────────────────
 * 
 * Location: Lines 2016-2045 in XLWorksheet::createDrawingRelationshipsFile()
 * 
 * Code:
 *   int relationshipId = 1;
 *   for (const auto& image : m_images) {
 *       std::string imageFilename = "xl/media/image_" + image.id() + image.extension();
 *       std::string imageRelativePath = "../media/image_" + image.id() + image.extension();
 *       
 *       relsXml += "<Relationship Id=\"rId" + std::to_string(relationshipId) + 
 *                  "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" " +
 *                  "Target=\"" + imageRelativePath + "\"/>\n";
 *       relationshipId++;
 *   }
 * 
 * Creates:
 *   File: xl/drawings/_rels/drawing1.xml.rels
 *   Content: <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image_img1.png"/>
 * 
 * Purpose: Links drawing.xml to image files
 * Uses: Simple sequential counter (rId1, rId2, rId3...)
 * 
 * ID Management Strategy:
 *   - Uses simple sequential numbering starting from 1
 *   - Regenerates entire relationship file each time
 *   - No conflict checking needed (file is recreated)
 *   - Fast and predictable for drawing-specific relationships
 * 
 * WHY TWO DIFFERENT STRATEGIES?
 * ──────────────────────────────────────────────────────────────────────────────
 * 
 * 1. SCOPE DIFFERENCES:
 *    - Worksheet relationships: Complex, shared namespace with other worksheet parts
 *    - Drawing relationships: Simple, isolated to drawing objects only
 * 
 * 2. FREQUENCY OF CHANGES:
 *    - Worksheet relationships: Infrequent, stable (drawing, comments, tables)
 *    - Drawing relationships: Frequent, dynamic (images added/removed often)
 * 
 * 3. COMPLEXITY REQUIREMENTS:
 *    - Worksheet relationships: Must avoid conflicts with existing relationships
 *    - Drawing relationships: Can use simple sequential approach
 * 
 * 4. PERFORMANCE CONSIDERATIONS:
 *    - Worksheet relationships: Intelligent scanning for uniqueness
 *    - Drawing relationships: Fast regeneration of entire file
 * 
 * RELATIONSHIP FILE STRUCTURE:
 * ──────────────────────────────────────────────────────────────────────────────
 * 
 * xl/worksheets/_rels/sheet1.xml.rels:
 *   <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 *     <Relationship Id="rId1" Type="..." Target="../drawings/drawing1.xml"/>
 *     <Relationship Id="rId2" Type="..." Target="../comments/comments1.xml"/>
 *     <Relationship Id="rId3" Type="..." Target="../tables/table1.xml"/>
 *   </Relationships>
 * 
 * xl/drawings/_rels/drawing1.xml.rels:
 *   <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 *     <Relationship Id="rId1" Type="..." Target="../media/image_img1.png"/>
 *     <Relationship Id="rId2" Type="..." Target="../media/image_img2.jpg"/>
 *     <Relationship Id="rId3" Type="..." Target="../media/image_img3.gif"/>
 *   </Relationships>
 * 
 * KEY INSIGHT:
 * Both relationship types can use the same ID (e.g., "rId1") because they exist
 * in separate relationship files with independent namespaces. This is the
 * fundamental principle of Excel's relationship scoping system.
 * 
 * ================================================================================
 */

namespace OpenXLSX
{
    // ========== Global Helper Functions ========== //
    
    /**
     * @brief Helper function to append debug messages for debugging
     * @param dbgMsg Optional pointer to append debug message (up to 16384 characters)
     * @param msg The message to append
     */
    inline void appendDbgMsg(std::string* dbgMsg, const std::string& msg)
    {
        if (dbgMsg && dbgMsg->length() < 16000) { // Leave some buffer
            if (!dbgMsg->empty()){
                dbgMsg->append("\n  ");
            }
            dbgMsg->append(msg);
        }
    }

    // ========== Excel Constants ========== //
    
    /**
     * @brief Standard Excel cell height in EMUs (15 points × 12700 EMUs/point)
     */
    constexpr uint32_t EXCEL_CELL_HEIGHT_EMUS = 190500;
    
    /**
     * @brief Excel cell vertical spacing in EMUs (includes cell height + divider)
     * This is approximately 198000 EMUs (cell height × 26/25)
     */
    constexpr uint32_t EXCEL_CELL_Y_SPACING_EMUS = 198000;
    
    /**
     * @brief Standard EMU-to-pixel conversion ratio (approximately)
     * Note: This may vary based on screen resolution
     */
    constexpr uint32_t EMU_TO_PIXEL_RATIO = 9525;

    /**
     * @brief Structure containing image metadata for query operations
     * 
     * This structure holds all the essential information about an embedded image
     * that can be queried from a worksheet. It provides a unified interface
     * for accessing image properties without needing to parse XML directly.
     * 
     * The structure is designed to be efficient for both storage and access,
     * with all data pre-parsed and readily available for common operations.
     */
    struct OPENXLSX_EXPORT XLImageInfo
    {
        std::string imageId;           /**< Unique image identifier (e.g., "img1", "img2") */
        std::string relationshipId;    /**< Relationship ID in drawing.xml.rels (e.g., "rId1", "rId2") */
        std::string anchorCell;        /**< Primary anchor cell reference (e.g., "A1", "B5") */
        std::string anchorType;        /**< Anchor type: "oneCellAnchor" or "twoCellAnchor" */
        uint32_t widthPixels{0};       /**< Original image width in pixels */
        uint32_t heightPixels{0};      /**< Original image height in pixels */
        uint32_t displayWidthEMUs{0};  /**< Display width in Excel units (EMUs) */
        uint32_t displayHeightEMUs{0}; /**< Display height in Excel units (EMUs) */
        
        /**
         * @brief Default constructor
         */
        XLImageInfo() = default;
        
        /**
         * @brief Constructor with all parameters
         * @param id Image ID
         * @param relId Relationship ID
         * @param cell Anchor cell reference
         * @param type Anchor type
         * @param wPx Width in pixels
         * @param hPx Height in pixels
         * @param wEmu Display width in EMUs
         * @param hEmu Display height in EMUs
         */
        XLImageInfo(const std::string& id, const std::string& relId, const std::string& cell,
                   const std::string& type, uint32_t wPx, uint32_t hPx, uint32_t wEmu, uint32_t hEmu)
            : imageId(id), relationshipId(relId), anchorCell(cell), anchorType(type),
              widthPixels(wPx), heightPixels(hPx), displayWidthEMUs(wEmu), displayHeightEMUs(hEmu) {}
        
        /**
         * @brief Check if this image info is valid (has non-empty imageId)
         * @return True if valid, false otherwise
         */
        bool isValid() const { return !imageId.empty(); }
        
        /**
         * @brief Check if this image info is empty (default constructed)
         * @return True if empty, false otherwise
         */
        bool isEmpty() const { return imageId.empty(); }

        /**
         * @brief Compare two image info structures for debugging
         * @param other The other XLImageInfo to compare with
         * @param diffMsg Optional pointer to append difference message (up to 16384 characters)
         * @return 0 if identical, <0 if this precedes other, >0 if this follows other
         */
        int compare(const XLImageInfo& other, std::string* diffMsg = nullptr) const;
    };

    /**
     * @brief The XLImage class represents an image that can be embedded in a worksheet
     */
    class OPENXLSX_EXPORT XLImage
    {
    public:
        /**
         * @brief Default constructor
         */
        XLImage() = default;

        /**
         * @brief Constructor from file path
         * @param imagePath Path to the image file
         */
        explicit XLImage(const std::string& imagePath);

        /**
         * @brief Constructor from binary data
         * @param imageData Binary image data
         * @param mimeType MIME type of the image
         */
        XLImage(const std::vector<uint8_t>& imageData, const std::string& mimeType);

        /**
         * @brief Copy constructor
         * @param other The XLImage to copy
         */
        XLImage(const XLImage& other) = default;

        /**
         * @brief Move constructor
         * @param other The XLImage to move
         */
        XLImage(XLImage&& other) noexcept = default;

        /**
         * @brief Destructor
         */
        ~XLImage() = default;

        /**
         * @brief Copy assignment operator
         * @param other The XLImage to copy
         * @return Reference to this XLImage
         */
        XLImage& operator=(const XLImage& other) = default;

        /**
         * @brief Move assignment operator
         * @param other The XLImage to move
         * @return Reference to this XLImage
         */
        XLImage& operator=(XLImage&& other) noexcept = default;

        /**
         * @brief Load image from file
         * @param imagePath Path to the image file
         * @return True if successful, false otherwise
         */
        bool loadFromFile(const std::string& imagePath);

        /**
         * @brief Load image from file
         * @param imagePath Path to the image file
         * @param imageId The unique ID for this image
         * @return True if successful, false otherwise
         */
        bool loadFromFile(const std::string& imagePath, const std::string& imageId);

        /**
         * @brief Load image from binary data
         * @param imageData Binary image data
         * @param mimeType MIME type of the image
         * @return True if successful, false otherwise
         */
        bool loadFromData(const std::vector<uint8_t>& imageData, const std::string& mimeType);

        /**
         * @brief Load image from binary data
         * @param imageData Binary image data
         * @param mimeType MIME type of the image
         * @param imageId The unique ID for this image
         * @return True if successful, false otherwise
         */
        bool loadFromData(const std::vector<uint8_t>& imageData, const std::string& mimeType, const std::string& imageId);

        /**
         * @brief Set the image ID
         * @param imageId The unique ID for this image
         */
        void setId(const std::string& imageId);

        /**
         * @brief Get the image data
         * @return Reference to the image data vector
         */
        const std::vector<uint8_t>& data() const;

        /**
         * @brief Get the MIME type
         * @return The MIME type string
         */
        const std::string& mimeType() const;

        /**
         * @brief Get the file extension
         * @return The file extension (e.g., ".png", ".jpg")
         */
        const std::string& extension() const;

        /**
         * @brief Get the unique image ID
         * @return The unique image ID
         */
        const std::string& id() const;

        /**
         * @brief Get the image width in pixels
         * @return The width in pixels
         */
        uint32_t widthPixels() const;

        /**
         * @brief Get the image height in pixels
         * @return The height in pixels
         */
        uint32_t heightPixels() const;

        /**
         * @brief Set the display width in Excel units (EMUs)
         * @param width The display width in EMUs
         */
        void setDisplayWidth(uint32_t width);

        /**
         * @brief Set the display height in Excel units (EMUs)
         * @param height The display height in EMUs
         */
        void setDisplayHeight(uint32_t height);

        /**
         * @brief Get the display width in Excel units (EMUs)
         * @return The display width in EMUs
         */
        uint32_t displayWidth() const;

        /**
         * @brief Get the display height in Excel units (EMUs)
         * @return The display height in EMUs
         */
        uint32_t displayHeight() const;

        /**
         * @brief Set display size in pixels (converts to EMUs automatically)
         * @param widthPixels The display width in pixels
         * @param heightPixels The display height in pixels
         */
        void setDisplaySizePixels(uint32_t widthPixels, uint32_t heightPixels);

        /**
         * @brief Set display size maintaining aspect ratio
         * @param maxWidthEmus Maximum width in EMUs (0 = no limit)
         * @param maxHeightEmus Maximum height in EMUs (0 = no limit)
         */
        void setDisplaySizeWithAspectRatio(uint32_t maxWidthEmus = 0, uint32_t maxHeightEmus = 0);

        /**
         * @brief Set display size maintaining aspect ratio
         * @param maxWidthPixels Maximum width in pixels (0 = no limit)
         * @param maxHeightPixels Maximum height in pixels (0 = no limit)
         */
        void setDisplaySizePixelsWithAspectRatio(uint32_t maxWidthPixels = 0, uint32_t maxHeightPixels = 0);


        /**
         * @brief Convert pixels to EMUs (Excel units)
         * @param pixels Number of pixels
         * @return Equivalent EMUs
         */
        static uint32_t pixelsToEmus(uint32_t pixels);

        /**
         * @brief Convert EMUs to pixels
         * @param emus Number of EMUs
         * @return Equivalent pixels
         */
        static uint32_t emusToPixels(uint32_t emus);

        /**
         * @brief Check if the image is valid
         * @return True if the image has valid data
         */
        bool isValid() const;

        /**
         * @brief Get the content type for this image
         * @return The XLContentType enum value
         */
        XLContentType contentType() const;

        /**
         * @brief Compare two images for debugging
         * @param other The other XLImage to compare with
         * @param diffMsg Optional pointer to append difference message (up to 16384 characters)
         * @return 0 if identical, <0 if this precedes other, >0 if this follows other
         */
        int compare(const XLImage& other, std::string* diffMsg = nullptr) const;

        /**
         * @brief Verify internal data integrity and class invariants
         * @param dbgMsg Optional pointer to append debug message explaining any errors found
         * @return Number of errors found (0 = EXIT_SUCCESS, data integrity OK)
         */
        int verifyData(std::string* dbgMsg = nullptr) const;

    private:
        std::vector<uint8_t> m_imageData;     /**< Binary image data */
        std::string m_mimeType;               /**< MIME type of the image */
        std::string m_extension;              /**< File extension */
        std::string m_id;                     /**< Unique image ID */
        uint32_t m_widthPixels{0};            /**< Image width in pixels */
        uint32_t m_heightPixels{0};           /**< Image height in pixels */
        uint32_t m_displayWidth{0};           /**< Display width in EMUs */
        uint32_t m_displayHeight{0};          /**< Display height in EMUs */

        /**
         * @brief Generate a unique image ID
         * @return A unique image ID string
         */
        static std::string generateId(uint32_t imageNumber);

        /**
         * @brief Determine MIME type from file extension
         * @param extension File extension
         * @return MIME type string
         */
        static std::string mimeTypeFromExtension(const std::string& extension);

        /**
         * @brief Determine file extension from MIME type
         * @param mimeType MIME type string
         * @return File extension string
         */
        static std::string extensionFromMimeType(const std::string& mimeType);

        /**
         * @brief Convert pixels to EMUs (Excel Measurement Units)
         * @param pixels Number of pixels
         * @return Number of EMUs
         */
        static uint32_t pixelsToEMUs(uint32_t pixels);

        /**
         * @brief Get image dimensions from binary data
         * @param data Binary image data
         * @param mimeType MIME type of the image
         * @return Pair of (width, height) in pixels
         */
        static std::pair<uint32_t, uint32_t> getImageDimensions(const std::vector<uint8_t>& data, const std::string& mimeType);
    };

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLIMAGE_HPP

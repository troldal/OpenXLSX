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
#include <map>        // std::map
#include <string>     // std::string
#include <vector>     // std::vector
#include <utility>    // std::pair

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLContentTypes.hpp"
#include "XLDrawingML.hpp"

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
    // Forward declarations
    class XLDocument;

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
     * @brief MIME type enumeration for image formats
     * @details This enum provides type-safe representation of image MIME types,
     *          replacing string-based MIME types for better performance and type safety.
     */
    enum class XLMimeType : uint8_t {
        PNG,        /**< PNG image format */
        JPEG,       /**< JPEG image format */
        BMP,        /**< BMP image format */
        GIF,        /**< GIF image format */
        Unknown     /**< Unknown or unsupported image format */
    };

    /**
     * @brief Class representing shared binary data and metadata for an image file
     * @details This class stores metadata about an image and provides access to shared binary data
     *          through the parent document. It uses a hash-based approach to avoid storing
     *          duplicate binary data in memory.
     */
    class OPENXLSX_EXPORT XLImage {
    public:
        /**
         * @brief Default constructor
         */
        XLImage() = default;

        /**
         * @brief Constructor with all parameters
         * @param parentDoc Pointer to the parent XLDocument
         * @param imageDataHash Hash of the image binary data
         * @param mimeType MIME type of the image
         * @param extension File extension of the image
         * @param filePackagePath Package path in archive
         * @param widthPixels Width in pixels
         * @param heightPixels Height in pixels
         */
        XLImage(XLDocument* parentDoc, size_t imageDataHash, XLMimeType mimeType, 
                const std::string& extension, const std::string& filePackagePath, 
                uint32_t widthPixels, uint32_t heightPixels);

        /**
         * @brief Get the image binary data
         * @return Vector containing the image binary data
         */
        std::vector<uint8_t> getImageData() const;

        /**
         * @brief Verify the image data integrity and metadata consistency
         * @param dbgMsg Optional pointer to append debug message
         * @return Number of issues found (0 = no issues)
         */
        int verifyData(std::string* dbgMsg = nullptr) const;

        // Getters
        XLDocument* parentDoc() const { return m_parentDoc; }
        size_t imageDataHash() const { return m_imageDataHash; }
        XLMimeType mimeType() const { return m_mimeType; }
        const std::string& extension() const { return m_extension; }
        const std::string& filePackagePath() const { return m_filePackagePath; }
        uint32_t widthPixels() const { return m_widthPixels; }
        uint32_t heightPixels() const { return m_heightPixels; }

        // Setters
        void setParentDoc(XLDocument* parentDoc) { m_parentDoc = parentDoc; }
        void setImageDataHash(size_t imageDataHash) { m_imageDataHash = imageDataHash; }
        void setMimeType(XLMimeType mimeType) { m_mimeType = mimeType; }
        void setExtension(const std::string& extension) { m_extension = extension; }
        void setFilePackagePath(const std::string& filePackagePath) { m_filePackagePath = filePackagePath; }
        void setWidthPixels(uint32_t widthPixels) { m_widthPixels = widthPixels; }
        void setHeightPixels(uint32_t heightPixels) { m_heightPixels = heightPixels; }

    private:
        XLDocument* m_parentDoc = nullptr;        /**< Pointer to parent document */
        size_t m_imageDataHash = 0;               /**< Hash of the image binary data */
        XLMimeType m_mimeType = XLMimeType::Unknown; /**< MIME type of the image */
        std::string m_extension;                  /**< File extension of the image */
        std::string m_filePackagePath;           /**< Package path in archive */
        uint32_t m_widthPixels = 0;               /**< Width in pixels */
        uint32_t m_heightPixels = 0;             /**< Height in pixels */
    };

    /**
     * @brief Shared pointer type for XLImage
     */
    using XLImageShPtr = std::shared_ptr<XLImage>;

    /**
     * @brief Vector type for embedded images
     */
    using XLEmbImgVec = std::vector<XLEmbeddedImage>;

    /**
     * @brief Shared pointer type for embedded image vector
     */
    using XLEmbImgVecShPtr = std::shared_ptr<XLEmbImgVec>;

    /**
     * @brief Class managing a collection of XLImage objects and caching XLEmbImgVec objects 
     * @details This class manages a collection of shared pointers to XLImage objects,
     *          providing functionality to find, add, and prune images based on usage.
     *          In addition, it keeps XLEmbImgVec containers for XLWorksheet's when
     *          the XLWorksheet's go out of scope.
     */
    class OPENXLSX_EXPORT XLImageManager {    
    public:
        /**
         * @brief Constructor
         * @param parentDoc Pointer to the parent XLDocument
         */
        explicit XLImageManager(XLDocument* parentDoc);

        /**
         * @brief Destructor
         */
        ~XLImageManager() = default;

        /**
         * @brief Remove unused images from the collection and archive
         * @details For each XLImageShPtr image in images, if image.get() == nullptr or 
         *          image->use_count() <= 1, then remove image from images and remove 
         *          image from archive by calling parentDoc->deleteEntry()
         */
        void prune();

        /**
         * @brief Find or add an image to the collection
         * @param packagePath The image package path (e.g., "xl/media/image1.png") or empty string to generate new
         * @param imageData The image binary data
         * @param contentType The content type of the image
         * @return Shared pointer to the XLImage object
         * @details If image already exists (findImageByImageData), return image from images.
         *          Otherwise, call parentDoc->addImageEntry() and add image to images.
         */
        XLImageShPtr findOrAddImage(const std::string& packagePath, 
                                   const std::vector<uint8_t>& imageData, 
                                   XLContentType contentType);

        /**
         * @brief Find an image by its binary data hash
         * @param imageData The image binary data to search for
         * @return Shared pointer to the XLImage object, or nullptr if not found
         */
        XLImageShPtr findImageByImageData(const std::vector<uint8_t>& imageData) const;


        /**
         * @brief Generate a unique package filename for an image
         * @param extension The file extension (e.g., ".png", ".jpg")
         * @return A unique package filename (e.g., "xl/media/image1.png")
         * @details Generates sequential filenames like Excel does, but ensures uniqueness
         *          across all extensions (image1.jpg, image2.png, image3.gif, etc.)
         *          Uses O(n log n) efficiency with sorted vector and binary search.
         *          Relies on XLImageManager being the authoritative source for all image paths.
         */
        std::string generateUniquePackageFilename(const std::string& extension) const;


        /**
         * @brief Find an image by its file package path
         * @param filePackagePath The file package path to search for
         * @return Shared pointer to the XLImage object, or nullptr if not found
         */
        XLImageShPtr findImageByFilePackagePath(const std::string& filePackagePath) const;

        /**
         * @brief Get the number of images in the collection
         * @return Number of images
         */
        size_t getImageCount() const;

        /**
         * @brief Get iterator to the beginning of the images collection
         * @return Const iterator to the beginning
         */
        std::vector<XLImageShPtr>::const_iterator begin() const;

        /**
         * @brief Get iterator to the end of the images collection
         * @return Const iterator to the end
         */
        std::vector<XLImageShPtr>::const_iterator end() const;

        /**
         * @brief Remove embedded image vector for worksheet
         * @param sheetName The name of the worksheet
         */
        void eraseEmbImgsForSheetName(const std::string& sheetName);

        /**
         * @brief Find or create embedded image vector for worksheet
         * @param sheetName The name of the worksheet
         * @return Shared pointer to the embedded image vector
         */
        XLEmbImgVecShPtr getEmbImgsForSheetName(const std::string& sheetName);

        /**
         * @brief Find embedded image vector for worksheet, return null pointer if not found
         * @param sheetName The name of the worksheet
         * @return Shared pointer to the embedded image vector, or nullptr if not found
         */
        XLEmbImgVecShPtr getEmbImgsForSheetName(const std::string& sheetName) const;

        /**
         * @brief Verify data integrity for all images
         * @param dbgMsg Optional pointer to append debug message
         * @return Number of issues found (0 = no issues)
         * @details Call verifyData() for each non-null image, check each image has same parentDoc,
         *          check each image has unique imageData
         */
        int verifyData(std::string* dbgMsg = nullptr) const;

        /**
         * @brief Debug function to print reference counts and brief info for all XLImageShPtr objects
         * @param title Optional title for the debug output
         */
        void printReferenceCounts(const std::string& title = "XLImageManager Reference Counts") const;

    private:
        XLDocument* m_parentDoc;                 /**< Pointer to parent document */
        std::vector<XLImageShPtr> m_images;      /**< Collection of image shared pointers */
        std::map<std::string, XLEmbImgVecShPtr> m_sheetNmEmbImgVMap; /* cache embedded images for each worksheet */
    };

    /**
     * @brief The XLEmbeddedImage class represents an image that can be embedded in a worksheet
     */
    class OPENXLSX_EXPORT XLEmbeddedImage
    {
    public:
        /**
         * @brief Default constructor
         */
        XLEmbeddedImage() = default;

        /**
         * @brief Copy constructor
         * @param other The XLEmbeddedImage to copy
         */
        XLEmbeddedImage(const XLEmbeddedImage& other) = default;

        /**
         * @brief Move constructor
         * @param other The XLEmbeddedImage to move
         */
        XLEmbeddedImage(XLEmbeddedImage&& other) noexcept = default;

        /**
         * @brief Destructor
         */
        ~XLEmbeddedImage() = default;

        /**
         * @brief Copy assignment operator
         * @param other The XLEmbeddedImage to copy
         * @return Reference to this XLEmbeddedImage
         */
        XLEmbeddedImage& operator=(const XLEmbeddedImage& other) = default;

        /**
         * @brief Move assignment operator
         * @param other The XLEmbeddedImage to move
         * @return Reference to this XLEmbeddedImage
         */
        XLEmbeddedImage& operator=(XLEmbeddedImage&& other) noexcept = default;

        /**
         * @brief Set the image ID
         * @param imageId The unique ID for this image
         */
        void setId(const std::string& imageId);

        /**
         * @brief Get the image data
         * @return Reference to the image data vector
         */
        std::vector<uint8_t> data() const;

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
         * @brief Get the image MIME type
         * @return The MIME type as enum
         */
        XLMimeType mimeType() const;

        /**
         * @brief Set the display width in Excel units (EMUs)
         * @param width The display width in EMUs
         */
        void setDisplayWidthEMUs(uint32_t width);

        /**
         * @brief Set the display height in Excel units (EMUs)
         * @param height The display height in EMUs
         */
        void setDisplayHeightEMUs(uint32_t height);

        /**
         * @brief Get the display width in Excel units (EMUs)
         * @return The display width in EMUs
         */
        uint32_t displayWidthEMUs() const;

        /**
         * @brief Get the display height in Excel units (EMUs)
         * @return The display height in EMUs
         */
        uint32_t displayHeightEMUs() const;

        /**
         * @brief Set display size in pixels (converts to EMUs automatically)
         * @param widthPixels The display width in pixels
         * @param heightPixels The display height in pixels
         */
        void setDisplaySizePixels(uint32_t widthPixels, uint32_t heightPixels);

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
         * @param other The other XLEmbeddedImage to compare with
         * @param diffMsg Optional pointer to append difference message (up to 16384 characters)
         * @return 0 if identical, <0 if this precedes other, >0 if this follows other
         */
        int compare(const XLEmbeddedImage& other, std::string* diffMsg = nullptr) const;

        /**
         * @brief Verify internal data integrity and class invariants
         * @param dbgMsg Optional pointer to append debug message explaining any errors found
         * @return Number of errors found (0 = EXIT_SUCCESS, data integrity OK)
         */
        int verifyData(std::string* dbgMsg = nullptr) const;

        // Registry/relationship setter functions (for XLImageInfo compatibility)
        /**
         * @brief Set the relationship ID
         * @param relationshipId The relationship ID in drawing.xml.rels
         */
        void setRelationshipId(const std::string& relationshipId);

        /**
         * @brief Set the anchor cell reference
         * @param anchorCell The primary anchor cell reference (e.g., "A1", "B5")
         */
        void setAnchorCell(const std::string& anchorCell);


        /**
         * @brief Set the anchor type
         * @param anchorType The anchor type enum
         */
        void setAnchorType(XLImageAnchorType anchorType);

        /**
         * @brief Convert row/column coordinates to Excel cell reference
         * @param row Row number (0-based)
         * @param col Column number (0-based)
         * @return Excel cell reference (e.g., "A1", "B5")
         */
        static std::string rowColToCellRef(uint32_t row, uint16_t col);

        /**
         * @brief Convert Excel cell reference to row/column coordinates
         * @param cellRef Excel cell reference (e.g., "A1", "B5")
         * @return Pair of (row, col) coordinates (0-based)
         */
        static std::pair<uint32_t, uint16_t> cellRefToRowCol(const std::string& cellRef);

        // Registry/relationship getter functions (for XLImageInfo compatibility)
        /**
         * @brief Get the relationship ID
         * @return The relationship ID string
         */
        const std::string& relationshipId() const;

        /**
         * @brief Get the anchor cell reference
         * @return The anchor cell reference string
         */
        std::string anchorCell() const;

        /**
         * @brief Get the anchor type
         * @return anchor type
         */
        const XLImageAnchorType& anchorType() const;

        /**
         * @brief Get the shared pointer to the image data
         * @return Shared pointer to the XLImage
         */
        XLImageShPtr getImage() const;

        /**
         * @brief Set the shared pointer to the image data
         * @param image Shared pointer to the XLImage
         */
        void setImage(XLImageShPtr image);

        const XLImageAnchor& getImageAnchor() const{ return m_imageAnchor; }

        void setImageAnchor(const XLImageAnchor& a){ m_imageAnchor = a; }

        /**
         * @brief Check if this image is empty (default constructed)
         * @return True if empty, false otherwise
         */
        bool isEmpty() const;

    private:
        XLImageShPtr m_image;                   /**< Shared pointer to the image data */
        XLImageAnchor m_imageAnchor;           /**< Image anchor information (positioning, IDs, display dimensions) */

        /**
         * @brief Generate a unique image ID
         * @return A unique image ID string
         */
        static std::string generateId(uint32_t imageNumber);
    };

    /**
     * @brief Utility class for image operations and test data
     */
    class OPENXLSX_EXPORT XLImageUtils
    {
    public:
        // Static const sample binary data for testing
        static const std::vector<uint8_t> png31x15Data;
        static const std::vector<uint8_t> jpeg32x18Data;
        static const std::vector<uint8_t> bmp33x16Data;
        static const std::vector<uint8_t> gif26x16Data;

        /**
         * @brief Detect MIME type from binary image data
         * @param data Binary image data
         * @return MIME type string (e.g., "image/png")
         */
        static std::string detectMimeType(const std::vector<uint8_t>& data);

        /**
         * @brief Get file extension from MIME type
         * @param mime MIME type string
         * @return File extension (e.g., ".png")
         */
        static std::string extensionFromMime(const std::string& mime);

        /**
         * @brief Determine MIME type from file extension
         * @param extension File extension
         * @return MIME type
         */
        static XLMimeType mimeTypeFromExtension(const std::string& extension);

        /**
         * @brief Determine file extension from MIME type
         * @param mimeType MIME type string
         * @return File extension string
         */
        static std::string extensionFromMimeType(XLMimeType mimeType);

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
        static std::pair<uint32_t, uint32_t> getImageDimensions(const std::vector<uint8_t>& data, XLMimeType mimeType);

        /**
         * @brief Convert pixels to EMUs (Excel units)
         * @param pixels Number of pixels
         * @return Equivalent EMUs
         @deprecated  -- duplicate of pixelsToEMUs()
         */
        static uint32_t pixelsToEmus(uint32_t pixels);

        /**
         * @brief Convert EMUs to pixels
         * @param emus Number of EMUs
         * @return Equivalent pixels
         */
        static uint32_t emusToPixels(uint32_t emus);

        /**
         * @brief Convert EMUs to Excel units
         * @param emus EMU value
         * @return Excel unit value
         */
        static uint32_t emusToExcelUnits(uint32_t emus);

        /**
         * @brief Convert points to EMUs
         * @param points Point value
         * @return EMU value
         */
        static uint32_t pointsToEmus(double points);

        /**
         * @brief Validate image data integrity
         * @param data Binary image data
         * @param mimeType MIME type of the image
         * @return True if image data is valid
         */
        static bool validateImageData(const std::vector<uint8_t>& data, const XLMimeType& mimeType);

        /**
         * @brief Calculate aspect ratio from dimensions
         * @param width Width in pixels
         * @param height Height in pixels
         * @return Aspect ratio (width/height)
         */
        static double calculateAspectRatio(uint32_t width, uint32_t height);

        /**
         * @brief Get image dimensions from binary data (auto-detect MIME type)
         * @param data Binary image data
         * @return Pair of (width, height) in pixels, or {0,0} if detection fails
         */
        static std::pair<uint32_t, uint32_t> getImageDimensions(const std::vector<uint8_t>& data);

        /**
         * @brief Check if image dimensions are reasonable
         * @param width Width in pixels
         * @param height Height in pixels
         * @return True if dimensions are within reasonable bounds
         */
        static bool isValidImageSize(uint32_t width, uint32_t height);

        /**
         * @brief Calculate hash value for binary image data
         * @param imageData Binary image data as vector<uint8_t>
         * @return Hash value for the image data
         */
        static size_t imageDataHash(const std::vector<uint8_t>& imageData);

        /**
         * @brief Calculate hash value for image data stored as string
         * @param imageDataStr Binary image data as string
         * @return Hash value for the image data (same as imageDataHash for equivalent data)
         */
        static size_t imageDataStrHash(const std::string& imageDataStr);

        /**
         * @brief Match XML element name with target name (handles namespace prefixes)
         * @param elementName The XML element name (e.g., "xdr:oneCellAnchor")
         * @param targetName The target name to match (e.g., "oneCellAnchor")
         * @return True if the element name matches the target name
         */
        static bool matchesElementName(const std::string& elementName, const std::string& targetName);

        /**
         * @brief Convert MIME type to XLContentType enum
         * @param mimeType MIME type enum
         * @return Corresponding XLContentType enum value
         */
        static XLContentType mimeTypeToContentType(const XLMimeType& mimeType);

        /**
         * @brief Convert XLContentType enum to MIME type
         * @param contentType XLContentType enum value
         * @return Corresponding MIME type
         */
        static XLMimeType contentTypeToMimeType(XLContentType contentType);

        /**
         * @brief Convert MIME type string to XLMimeType enum
         * @param mimeType MIME type string (e.g., "image/png")
         * @return Corresponding XLMimeType enum value
         */
        static XLMimeType stringToMimeType(const std::string& mimeType);

        /**
         * @brief Convert XLMimeType enum to MIME type string
         * @param mimeType XLMimeType enum value
         * @return Corresponding MIME type string (e.g., "image/png")
         */
        static std::string mimeTypeToString(XLMimeType mimeType);

        /**
         * @brief Detect MIME type from binary image data and return as enum
         * @param data Binary image data
         * @return Corresponding XLMimeType enum value
         */
        static XLMimeType detectMimeTypeEnum(const std::vector<uint8_t>& data);
    };

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLIMAGE_HPP

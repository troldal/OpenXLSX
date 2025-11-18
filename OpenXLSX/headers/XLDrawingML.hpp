/*
 * XLDrawingML.hpp
 * 
 * DrawingML (Drawing Markup Language) support for OpenXLSX
 * This provides modern Excel-compatible drawing format instead of legacy VML
 */

#ifndef OPENXLSX_XLDRAWINGML_HPP
#define OPENXLSX_XLDRAWINGML_HPP

// ===== External Includes ===== //
#include <cstdint>
#include <string>
#include <vector>
#include <utility>    // std::pair

// ===== OpenXLSX Includes ===== //
#include "XLXmlFile.hpp"
#include "XLCellReference.hpp"

namespace OpenXLSX
{
    // Forward declarations
    class XLEmbeddedImage;
    class XLImageAnchor;
    enum class XLMimeType : uint8_t;

    /**
     * @brief Image anchor type enumeration
     * @details Defines the types of image anchors supported in Excel worksheets
     */
    enum class XLImageAnchorType : uint8_t {
        OneCellAnchor,  /**< Single cell anchor - image positioned relative to one cell */
        TwoCellAnchor,  /**< Two cell anchor - image spans between two cells */
        Unknown         /**< Unknown or unsupported anchor type */
    };

    /**
     * @brief Utility functions for XLImageAnchorType enum
     */
    namespace XLImageAnchorUtils {
        /**
         * @brief Convert anchor type string to XLImageAnchorType enum
         * @param anchorTypeStr Anchor type string (e.g., "oneCellAnchor", "twoCellAnchor")
         * @return Corresponding XLImageAnchorType enum value
         */
        XLImageAnchorType stringToAnchorType(const std::string& anchorTypeStr);

        /**
         * @brief Convert XLImageAnchorType enum to anchor type string
         * @param anchorType XLImageAnchorType enum value
         * @return Corresponding anchor type string (e.g., "oneCellAnchor", "twoCellAnchor")
         */
        std::string anchorTypeToString(XLImageAnchorType anchorType);
    }

    /**
     * @brief Class for managing DrawingML drawings in worksheets
     */
    class OPENXLSX_EXPORT XLDrawingML : public XLXmlFile
    {
    public:
        // ===== Constructors & Destructors ===== //
        XLDrawingML() = default;
        XLDrawingML(XLXmlData* xmlData);
        ~XLDrawingML() = default;

        // ===== Copy & Move Constructors & Assignment Operators ===== //
        XLDrawingML(const XLDrawingML& other) = default;
        XLDrawingML(XLDrawingML&& other) noexcept = default;
        XLDrawingML& operator=(const XLDrawingML& other) = default;
        XLDrawingML& operator=(XLDrawingML&& other) noexcept = default;

        void addImage(const XLEmbeddedImage& embImage);

        // ===== Public Methods ===== //
        /**
         * @brief Get the number of images in the drawing
         * @return Number of images
         */
        size_t imageCount() const;

        /**
         * @brief Verify internal data integrity and class invariants
         * @param dbgMsg Optional pointer to append debug message explaining any errors found
         * @return Number of errors found (0 = EXIT_SUCCESS, data integrity OK)
         */
        int verifyData(std::string* dbgMsg = nullptr) const;

        /**
         * @brief Get the XML root node for verification purposes
         * @return XML root node (empty if invalid)
         */
        XMLNode getRootNode() const;

        // ===== Static Utility Functions ===== //
        
        /**
         * @brief Create XML anchor node from XLImageAnchor data
         * @param rootNode The root XML node to append the anchor to
         * @param imgAnchorInfo The anchor data to convert to XML
         * @return The created anchor XML node
         */
        static XMLNode createXMLNode(XMLNode rootNode, const XLImageAnchor& imgAnchorInfo);
        
        /**
         * @brief Delete XML anchor node by relationship ID
         * @param rootNode The root XML node to search in
         * @param relationshipId The relationship ID to find and delete
         * @return True if the anchor was found and deleted, false otherwise
         */
        static bool deleteXMLNode(XMLNode rootNode, const std::string& relationshipId);
        
        /**
         * @brief Parse XML anchor node into XLImageAnchor data
         * @param anchorXMLNode The XML node containing anchor data
         * @param imgAnchorInfo Pointer to XLImageAnchor to populate
         * @return True if parsing was successful, false otherwise
         */
        static bool parseXMLNode(const pugi::xml_node& anchorXMLNode, XLImageAnchor* imgAnchorInfo);

    private:
        // ===== Private Methods ===== //

        /**
         * @brief Calculate cell position in Excel units
         * @param row Row number (1-based)
         * @param column Column number (1-based)
         * @return Position in Excel units
         */
        uint32_t calculateCellPosition(uint32_t row, uint16_t column) const;

    };

    // ========== XLImageAnchor Class ========== //
    
    /**
     * @brief Represents data from <oneCellAnchor> and <twoCellAnchor> XML elements
     * 
     * This class directly maps to the XML structure found in Excel's DrawingML
     * anchor elements. All data members correspond exactly to XML attributes and
     * text content, with no automatic coordinate conversions.
     */
    class OPENXLSX_EXPORT XLImageAnchor
    {
    public:
        // ===== Anchor Type and Identification ===== //
        XLImageAnchorType anchorType;  /**< OneCellAnchor or TwoCellAnchor */
        std::string imageId;           /**< Worksheet-level image identifier (e.g., "1", "2")
                                            from xdr:cNvPr id attribute. Maps to XLEmbeddedImage::m_id. 
                                            XML: <xdr:cNvPr id="1" name="Picture 1"/> */
        std::string relationshipId;    /**< Relationship ID from a:blip r:embed attribute */
        
        // ===== Cell Positioning (From Element) - Direct XML mapping ===== //
        uint32_t fromRow;             /**< Row number from xdr:row (0-based in XML) */
        uint16_t fromCol;             /**< Column number from xdr:col (0-based in XML) */
        int32_t fromRowOffset;        /**< Row offset from xdr:rowOff (in EMUs) */
        int32_t fromColOffset;        /**< Column offset from xdr:colOff (in EMUs) */
        
        // ===== Cell Positioning (To Element) - for twoCellAnchor ===== //
        uint32_t toRow;               /**< To row number from xdr:row (0-based in XML) */
        uint16_t toCol;               /**< To column number from xdr:col (0-based in XML) */
        int32_t toRowOffset;          /**< To row offset from xdr:rowOff (in EMUs) */
        int32_t toColOffset;          /**< To column offset from xdr:colOff (in EMUs) */
        
        // ===== Display Dimensions (xdr:ext element) ===== //
        uint32_t displayWidthEMUs;    /**< Display width from xdr:ext cx attribute (in EMUs) */
        uint32_t displayHeightEMUs;   /**< Display height from xdr:ext cy attribute (in EMUs) */
        
        // ===== Actual Image Dimensions (a:xfrm/a:ext element) ===== //
        uint32_t actualWidthEMUs;     /**< Actual image width from a:xfrm/a:ext cx (in EMUs) */
        uint32_t actualHeightEMUs;    /**< Actual image height from a:xfrm/a:ext cy (in EMUs) */
        
        // ===== Constructors ===== //
        XLImageAnchor() = default;

        bool operator==( const XLImageAnchor& other ) const = default;
        
        void reset();

        void initOneCell( const XLCellReference& cellRef, int32_t rowOffset = 0,
            int32_t colOffset = 0);

        void initTwoCell( const XLCellReference& fromCellRef, 
            const XLCellReference& toCellRef, int32_t fromROffset = 0,
            int32_t fromCOffset = 0, int32_t toROffset = 0, int32_t toCOffset = 0);

        
        // ===== Utility Methods ===== //
        
        XLCellReference getFromCellReference() const;

        void setFromCellReference( const XLCellReference& fromCellRef );

        XLCellReference getToCellReference() const;

        void setToCellReference( const XLCellReference& toCellRef );

        void setDisplaySizeWithAspectRatio(
            const uint32_t& widthPixels, const uint32_t& heightPixels,
            const uint32_t& maxWidthEmus, const uint32_t& maxHeightEmus );

        void setDisplaySizeWithAspectRatio(
            const std::vector<uint8_t>& imageData, const XLMimeType& mimeType,
            const uint32_t& maxWidthEmus, const uint32_t& maxHeightEmus );

        void setDisplaySizeWithAspectRatio(
            const std::string& imageFileName, const XLMimeType& mimeType,
            const uint32_t& maxWidthEmus, const uint32_t& maxHeightEmus );
        
        /**
         * @brief Check if this is a twoCellAnchor
         * @return True if twoCellAnchor, false if oneCellAnchor
         */
        bool isTwoCell() const;
        
        /**
         * @brief Get display dimensions as a pair
         * @return Pair of (width, height) in EMUs
         */
        std::pair<uint32_t, uint32_t> getDisplayDimensions() const;
        
        /**
         * @brief Get actual image dimensions as a pair
         * @return Pair of (width, height) in EMUs
         */
        std::pair<uint32_t, uint32_t> getActualDimensions() const;
        
        /**
         * @brief Set actual image dimensions
         * @param width Width in EMUs
         * @param height Height in EMUs
         */
        void setActualDimensions(uint32_t width, uint32_t height);
        
        /**
         * @brief Convert EMUs to pixels for display dimensions
         * @return Pair of (width, height) in pixels
         */
        std::pair<uint32_t, uint32_t> getDisplayDimensionsPixels() const;
        
        /**
         * @brief Convert EMUs to pixels for actual dimensions
         * @return Pair of (width, height) in pixels
         */
        std::pair<uint32_t, uint32_t> getActualDimensionsPixels() const;
    };
}

#endif // OPENXLSX_XLDRAWINGML_HPP

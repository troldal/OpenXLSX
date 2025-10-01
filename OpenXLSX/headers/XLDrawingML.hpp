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

// ===== OpenXLSX Includes ===== //
#include "XLXmlFile.hpp"
#include "XLImage.hpp"

namespace OpenXLSX
{
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

        // ===== Public Methods ===== //
        /**
         * @brief Add an image to the drawing
         * @param image The image to add
         * @param row The row number (1-based)
         * @param column The column number (1-based)
         * @param relationshipId The relationship ID for the image
         */
        void addImage(const XLImage& image, uint32_t row, uint16_t column, const std::string& relationshipId);

        /**
         * @brief Add an image to the drawing with precise positioning
         * @param image The image to add
         * @param row The row number (1-based)
         * @param column The column number (1-based)
         * @param relationshipId The relationship ID for the image
         * @param rowOffset Offset from the top-left of the cell in EMUs
         * @param colOffset Offset from the top-left of the cell in EMUs
         */
        void addImageWithOffset(const XLImage& image, uint32_t row, uint16_t column, 
                               const std::string& relationshipId,
                               int32_t rowOffset = 0, int32_t colOffset = 0);

        /**
         * @brief Add an image to the drawing with two-cell anchoring
         * @param image The image to add
         * @param fromRow Starting row number (1-based)
         * @param fromCol Starting column number (1-based)
         * @param toRow Ending row number (1-based)
         * @param toCol Ending column number (1-based)
         * @param relationshipId The relationship ID for the image
         * @param fromRowOffset Offset from top-left of starting cell in EMUs
         * @param fromColOffset Offset from top-left of starting cell in EMUs
         * @param toRowOffset Offset from top-left of ending cell in EMUs
         * @param toColOffset Offset from top-left of ending cell in EMUs
         */
        void addImageTwoCellAnchor(const XLImage& image,
                                  uint32_t fromRow, uint16_t fromCol,
                                  uint32_t toRow, uint16_t toCol,
                                  const std::string& relationshipId,
                                  int32_t fromRowOffset = 0, int32_t fromColOffset = 0,
                                  int32_t toRowOffset = 0, int32_t toColOffset = 0);

        /**
         * @brief Get the number of images in the drawing
         * @return Number of images
         */
        size_t imageCount() const;

    private:
        // ===== Private Methods ===== //
        /**
         * @brief Convert EMUs to Excel units
         * @param emus EMU value
         * @return Excel unit value
         */
        uint32_t emusToExcelUnits(uint32_t emus) const;

        /**
         * @brief Convert points to EMUs
         * @param points Point value
         * @return EMU value
         */
        uint32_t pointsToEmus(double points) const;

        /**
         * @brief Calculate cell position in Excel units
         * @param row Row number (1-based)
         * @param column Column number (1-based)
         * @return Position in Excel units
         */
        uint32_t calculateCellPosition(uint32_t row, uint16_t column) const;
    };
}

#endif // OPENXLSX_XLDRAWINGML_HPP

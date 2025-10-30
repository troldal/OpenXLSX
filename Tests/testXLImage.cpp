//
// Created by OpenXLSX Image Test Suite
//

#include <OpenXLSX.hpp>
#include <catch.hpp>
#include <fstream>

using namespace OpenXLSX;

TEST_CASE("XLEmbeddedImage", "[XLEmbeddedImage]")
{
    SECTION("Basic Image Embedding Test") {
        
        // Create a new workbook
        XLDocument doc;
        doc.create("./testXLEmbeddedImage.xlsx");
        
        // Get the first worksheet and rename it
        XLWorksheet worksheet1 = doc.workbook().worksheet(1);
        worksheet1.setName("Images_Sheet1");
        REQUIRE(worksheet1.name() == "Images_Sheet1");
        
        // Add a second worksheet
        doc.workbook().addWorksheet("Images_Sheet2");
        XLWorksheet worksheet2 = doc.workbook().worksheet("Images_Sheet2");
        REQUIRE(worksheet2.name() == "Images_Sheet2");
        
        // Add some images to the first worksheet
        try {
            // Create a simple PNG image (1x1 pixel red PNG)
            static const std::vector<uint8_t> red1x1PNGData = {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D,
                0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
                0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xDE, 0x00, 0x00, 0x00,
                0x0C, 0x49, 0x44, 0x41, 0x54, 0x08, 0xD7, 0x63, 0xF8, 0x0F, 0x00, 0x00,
                0x01, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
                0x42, 0x60, 0x82
            };
            
            const std::string imageId1 = worksheet1.embedImageFromImageData(
                XLCellReference(2,2), red1x1PNGData, XLMimeType::PNG );
            const std::string imageId2 = worksheet1.embedImageFromImageData(
                XLCellReference(5,3), red1x1PNGData, XLMimeType::PNG );
            const std::string imageId3 = worksheet2.embedImageFromImageData(
                XLCellReference(1,1), red1x1PNGData, XLMimeType::PNG );
            REQUIRE(!imageId1.empty());
            REQUIRE(!imageId3.empty());
            REQUIRE(!imageId3.empty());

            // Verify the document data integrity
            std::string dbgMsg;
            int result = doc.verifyData(&dbgMsg);
            REQUIRE(result == EXIT_SUCCESS);
            
            // Basic verification that images were added
            REQUIRE(worksheet1.imageCount() > 0);
            REQUIRE(worksheet2.imageCount() > 0);
            
        } catch (const std::exception& e) {
            FAIL("Exception during image embedding: " << e.what());
        }
        
        // Save and close the document
        doc.save();
        doc.close();
        
        // Clean up
        std::remove("./testXLEmbeddedImage.xlsx");
    }
}

/*
COMPREHENSIVE XLEmbeddedImage UNIT TEST OUTLINE
=======================================

This outline covers all functions related to embedded images in OpenXLSX:

1. BASIC IMAGE CREATION AND EMBEDDING
   - Test embedImage() with different image formats (PNG, JPEG, GIF, BMP)
   - Test embedImage() with different data sources (vector<uint8_t>, file path)
   - Test embedImage() with different positioning (various row/column combinations)
   - Test embedImage() with different anchor types (oneCellAnchor, twoCellAnchor)
   - Test embedImage() with different sizing strategies (original, aspect ratio, exact dimensions)

2. IMAGE QUERY OPERATIONS
   - Test getImageInfos() - retrieve all images in worksheet
   - Test getImageInfosAtCell() - retrieve images at specific cell
   - Test getImageInfosInRange() - retrieve images in cell range
   - Test getImageInfoByImageID() - retrieve specific image by ID
   - Test getImageInfoByRelationshipId() - retrieve specific image by relationship ID
   - Test imageCount() - verify correct count of images

3. IMAGE MODIFICATION OPERATIONS
   - Test removeImageByImageID() - remove specific image by ID
   - Test removeImageByRelationshipId() - remove specific image by relationship ID
   - Test clearImages() - remove all images from worksheet
   - Test image positioning changes
   - Test image size modifications

4. IMAGE REGISTRY OPERATIONS
   - Test refreshImageRegistry() - populate registry from XML
   - Test populateImagesFromXML() - populate m_images vector from registry
   - Test registry consistency with XML data
   - Test registry updates after image operations

5. DRAWINGML INTEGRATION
   - Test DrawingML XML generation for different anchor types
   - Test DrawingML XML parsing for Excel-generated files
   - Test DrawingML validation and error handling
   - Test DrawingML consistency with image registry

6. RELATIONSHIPS FILE OPERATIONS
   - Test relationship file creation and updates
   - Test relationship ID generation and uniqueness
   - Test relationship file parsing and validation
   - Test relationship consistency with image data

7. BINARY DATA HANDLING
   - Test binary image data storage and retrieval
   - Test MIME type detection and validation
   - Test file extension handling
   - Test binary data integrity verification

8. ERROR HANDLING AND EDGE CASES
   - Test invalid image data handling
   - Test invalid positioning parameters
   - Test duplicate image ID handling
   - Test missing relationship file handling
   - Test corrupted XML data handling

9. PERFORMANCE AND MEMORY
   - Test memory usage with large images
   - Test performance with many images
   - Test memory cleanup after image removal
   - Test efficient registry operations

10. CROSS-WORKSHEET OPERATIONS
    - Test image operations across multiple worksheets
    - Test image ID uniqueness across worksheets
    - Test relationship ID uniqueness across worksheets
    - Test document-level image operations

11. FILE I/O OPERATIONS
    - Test loading images from existing Excel files
    - Test saving images to new Excel files
    - Test round-trip operations (save/load/verify)
    - Test compatibility with Excel-generated files

12. DATA INTEGRITY VERIFICATION
    - Test verifyData() for various image scenarios
    - Test verifyUniqueImageIDs() for duplicate detection
    - Test verifyDrawingMLConsistency() for XML validation
    - Test verifyImagesVectorConsistency() for vector validation
    - Test verifyBinaryImageData() for binary data validation

13. CONST CORRECTNESS
    - Test const member function behavior
    - Test const reference returns
    - Test const_cast usage validation
    - Test mutable member usage

14. THREAD SAFETY
    - Test concurrent image operations
    - Test static const member usage
    - Test registry access from multiple threads
    - Test XML data access from multiple threads

15. API COMPATIBILITY
    - Test backward compatibility with existing code
    - Test API consistency across different image types
    - Test return type consistency
    - Test parameter validation

This comprehensive test suite would ensure robust image embedding functionality
across all supported scenarios and edge cases.
*/

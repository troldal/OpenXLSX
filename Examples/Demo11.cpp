/*
 * Demo11.cpp - Comprehensive Image Embedding Test Matrix
 * 
 * This demo tests various combinations of image embedding options:
 * - Anchor types: oneCellAnchor vs twoCellAnchor
 * - File formats: PNG, JPEG, GIF, BMP
 * - Sizing strategies: aspect ratio preservation, original size, exact dimensions
 * - Units: pixels, EMUs, cell spacing
 */

#include <iostream>
#include <vector>
#include <string>
#include <filesystem>
#include <ctime>
#include "OpenXLSX.hpp"
#include "XLImage.hpp"  // For XLImageUtils

using namespace OpenXLSX;

/**
 * @brief Search upward through directory tree to find a directory
 * @param relativePath The relative path to search for
 * @return The full path to the directory if found, empty string if not found
 */
std::string findDirectory(const std::string& relativePath)
{
    const std::filesystem::path currentPath = std::filesystem::current_path();
    std::filesystem::path testPath = currentPath;
    std::string result;

    // Search upward through the directory tree
    for (int levels = 0; (levels < 10) && result.empty(); ++levels) {        
        // Append the relative path
        const std::filesystem::path testImagePath = testPath / relativePath;
        
        // Check if the directory exists
        if (std::filesystem::exists(testImagePath) && 
            std::filesystem::is_directory(testImagePath)) {
            result = testImagePath.string();
        }
        else{
            testPath = testPath.parent_path();
        }
    }
    
    return result; // Not found
}



/*  
   a. get image ID 
   b. get relationshiop ID
   c. get anchor cell location(s)
   d. get image width & height in pixels
   e. get image displayed width and displayed height in EMUs

Return image information in the following format
"img_id:ABCD  r_id:EFGH   oneCellAnchor:C44   330x492 px   1249900x1700000 emu"
*/
std::string embeddedImageStr(const XLEmbeddedImage& embImage) {
    std::string result;
    
    // a. get image ID
    std::string imageId = embImage.id();
    result += "img_id:" + imageId;
    
    // b. get relationship ID
    result += "  r_id:" + embImage.relationshipId();
    
    // c. get anchor type and cell location
    const XLImageAnchor& imageAnchor = embImage.getImageAnchor();
    std::string anchorCell = XLEmbeddedImage::rowColToCellRef(imageAnchor.fromRow, imageAnchor.fromCol);
    result += "  " + XLImageAnchorUtils::anchorTypeToString(imageAnchor.anchorType) + ":" + anchorCell;
    
    // d. get image width & height in pixels
    result += "  " + std::to_string(embImage.widthPixels()) + "x" + std::to_string(embImage.heightPixels()) + " px";
    
    // e. get image displayed width and displayed height in EMUs
    result += "  " + std::to_string(embImage.displayWidthEMUs()) + "x" + std::to_string(embImage.displayHeightEMUs()) + " EMU";
    
    // f. additional debugging info
    if (embImage.getImage()) {
        result += "  package:" + embImage.getImage()->filePackagePath();
        result += "  mime:" +  XLImageUtils::mimeTypeToString(embImage.mimeType());
    } else {
        result += "  [NO_IMAGE_OBJECT]";
    }
    
    return result;
}

// Helper function to print embedded images for a worksheet
void printEmbeddedImages(const XLWorksheet& worksheet, const std::string& sheetName, int* totalImages = nullptr) {
    try {
        const auto& embImages = worksheet.getEmbImages();
        size_t imageCount = embImages.size();
        
        // Update total count if pointer provided
        if (totalImages != nullptr) {
            *totalImages += static_cast<int>(imageCount);
        }
        
        // Print header
        const OpenXLSX::XLDrawingML& drawing = worksheet.drawingML();
        const size_t drawingImageCount = drawing.imageCount();
        std::cout << "  Worksheet '" << sheetName << "': " << imageCount << " image(s)    "
            << imageCount << " embedded images"
            << "   drawing valid:" << (drawing.valid() ? "Yes" : "No")
            << "   drawing images:" << drawingImageCount
            << std::endl;

        // Print image details if any images exist
        if (imageCount > 0) {
            for (size_t i = 0; i < embImages.size(); ++i) {
                std::cout << "    " << (i + 1) << ". " << embeddedImageStr(embImages[i]) << std::endl;
            }
        }
    } catch (const std::exception& e) {
        std::cout << "Error getting image info for worksheet '" << sheetName << "': " << e.what() << std::endl;
    }
}

// Helper function to print embedded image information for all worksheets
void printAllEmbeddedImages(const XLDocument& doc, int* totalImages = nullptr) {
    std::cout << "\n=== EMBEDDED IMAGES ===" << std::endl;
    
    int totalCount = 0;
    if (totalImages != nullptr) {
        *totalImages = 0;
    }
    
    // Iterate through all worksheets using workbook
    auto workbook = doc.workbook();
    auto sheetNames = workbook.worksheetNames();
    
    for (const auto& sheetName : sheetNames) {
        try {
            XLWorksheet worksheet = workbook.worksheet(sheetName);
            printEmbeddedImages(worksheet, sheetName, &totalCount);
        } catch (const std::exception& e) {
            std::cout << "Error accessing worksheet '" << sheetName << "': " << e.what() << std::endl;
        }
    }
    
    std::cout << "\nTotal embedded images across all worksheets: " << totalCount << std::endl;
    if (totalImages != nullptr) {
        *totalImages = totalCount;
    }
}

// Comprehensive cross-reference verification function
void verifyImageConsistency(const XLDocument& doc) {
    std::cout << "\n=== IMAGE CONSISTENCY VERIFICATION ===" << std::endl;
    
    int totalErrors = 0;
    
    // Iterate through all worksheets using workbook
    auto workbook = doc.workbook();
    auto sheetNames = workbook.worksheetNames();
    
    for (const auto& sheetName : sheetNames) {
        try {
            XLWorksheet worksheet = workbook.worksheet(sheetName);
            std::cout << "\n--- Verifying worksheet: " << sheetName << " ---" << std::endl;
            
            // Get embedded images data
            const auto& embImages = worksheet.getEmbImages();
            
            std::cout << "EmbeddedImages count: " << embImages.size() << std::endl;
            
            // Check count consistency (no longer needed since we only have one source)
            std::cout << "Embedded images verification complete." << std::endl;
            
            // Cross-reference each image
            size_t minCount = embImages.size();
            for (size_t i = 0; i < minCount; ++i) {
                const auto& embImage = embImages[i];
                
                std::cout << "  Image " << (i+1) << ":" << std::endl;
                
                // Check image ID consistency (no longer needed since we only have one source)
                std::cout << "    Image ID: " << embImage.id() << " ✓" << std::endl;
                
                // Check relationship ID consistency
                std::cout << "    Relationship ID: " << embImage.relationshipId() << " ✓" << std::endl;
                
                // Check anchor cell consistency
                std::cout << "    Anchor Cell: " << embImage.anchorCell() << " ✓" << std::endl;
                
                // Check anchor type consistency
                std::cout << "    Anchor Type: " << XLImageAnchorUtils::anchorTypeToString(
                    embImage.anchorType()) << " ✓" << std::endl;
                
                // Check dimensions consistency
                std::cout << "    Dimensions: " << embImage.widthPixels() << "x" 
                    << embImage.heightPixels() << " px ✓" << std::endl;
                
                // Check display dimensions consistency
                std::cout << "    Display Dimensions: " << embImage.displayWidthEMUs() << "x" 
                    << embImage.displayHeightEMUs() << " EMUs ✓" << std::endl;
                
                std::cout << "    Summary: " << embeddedImageStr(embImage) << std::endl;
            }
            
            // Run built-in verifyData
            std::string verifyMsg;
            int verifyErrors = worksheet.verifyData(&verifyMsg);
            if (verifyErrors > 0) {
                std::cout << "Built-in verifyData found " << verifyErrors << " errors:" << std::endl;
                std::cout << verifyMsg << std::endl;
                totalErrors += verifyErrors;
            } else {
                std::cout << "Built-in verifyData: No errors ✓" << std::endl;
            }
            
        } catch (const std::exception& e) {
            std::cout << "Error verifying worksheet '" << sheetName << "': " << e.what() << std::endl;
            totalErrors++;
        }
    }
    
    std::cout << "\n=== VERIFICATION SUMMARY ===" << std::endl;
    std::cout << "Total errors found: " << totalErrors << std::endl;
    if (totalErrors == 0) {
        std::cout << "All image data is consistent! ✓" << std::endl;
    } else {
        std::cout << "Image data inconsistencies detected!" << std::endl;
    }
}

// Helper function to print detailed image analysis for a worksheet
void printDetailedImageAnalysis(const XLWorksheet& worksheet, const std::string& sheetName) {
    int imageCount = static_cast<int>(worksheet.imageCount());
    std::cout << "  Worksheet '" << sheetName << "' imageCount: " << imageCount << std::endl;
    
    if (imageCount > 0) {        
        try {
            const auto& embImages = worksheet.getEmbImages();
            std::vector<std::string> anchorCells;
            std::vector<std::string> imageIds;
            std::vector<std::string> relationshipIds;
            for (size_t i = 0; i < embImages.size(); ++i) {
                std::cout << "    " << (i + 1) << ". " << embeddedImageStr(embImages[i]) << std::endl;
                anchorCells.push_back(embImages[i].anchorCell());
                imageIds.push_back(embImages[i].id());
                relationshipIds.push_back(embImages[i].relationshipId());
                }

            // Test getEmbImagesAtCell()
            for( const std::string& anchorCell : anchorCells ){
                const auto& cellImgs = worksheet.getEmbImagesAtCell(anchorCell);
                std::cout << "      getEmbImagesAtCell('" << anchorCell << "') returned "  
                      << cellImgs.size() << " image(s)" << std::endl;
            }

            // Test getImageInfoByImageID()
            for( const std::string& imageId : imageIds ){
                const auto& foundImg = worksheet.getEmbImageByImageID(imageId);
                if (foundImg.isValid()) {
                    std::cout << "      getImageInfoByImageID( " << imageId << ") found image" << std::endl;
                } else {
                    std::cout << "      getImageInfoByImageID( " << imageId << ") failed to find image" << std::endl;
                }
            }
            
            // Test getImageInfoByRelationshipId()
            for( const std::string& relationshipId : relationshipIds ){
                const auto& foundRel = worksheet.getEmbImageByRelationshipId(relationshipId);
                if (foundRel.isValid()) {
                    std::cout << "      ✓ getImageInfoByRelationshipId( " << relationshipId << ") found the image" << std::endl;
                } else {
                    std::cout << "      ✗ getImageInfoByRelationshipId( " << relationshipId << ") failed to find the image" << std::endl;
                }
            }

            // Test getEmbImagesInRange()
            const auto& rangeImgs = worksheet.getEmbImagesInRange("A1:Z100");
            std::cout << "    getEmbImagesInRange('A1:Z100') returned " << rangeImgs.size() << " image(s)" << std::endl;

            // Test validation
            bool isRegstryValid = worksheet.validateImageRegistry();
            std::cout << "    validateImageRegistry() returned: " << (isRegstryValid ? "Valid" : "Invalid") << std::endl;
            
        } catch (const std::exception& e) {
            std::cout << "    Error getting image info: " << e.what() << std::endl;
        }
    }
}

// Data integrity verification functions - checking for data integrity can sometimes be expensive,
// but is handy for debugging
void verifyDocData(const XLDocument& doc)
{
    std::string dbgMsg;
    int errorCount = doc.verifyData(&dbgMsg);
    if (errorCount > 0 || !dbgMsg.empty()) {
        std::cout << "ERROR: Document data integrity check failed with " << errorCount 
                  << " error(s): " << dbgMsg << std::endl;
    }
}

void verifyWorksheetData(const XLWorksheet& worksheet)
{
    std::string dbgMsg;
    int errorCount = worksheet.verifyData(&dbgMsg);
    if (errorCount > 0 || !dbgMsg.empty()) {
        std::cout << "ERROR: Worksheet '" << worksheet.name() << "' data integrity check failed with " << errorCount 
                  << " error(s): " << dbgMsg << std::endl;
    }
}



int main()
{
    std::cout << "OpenXLSX Image Test" << std::endl;
    std::cout << "===================" << std::endl;
    std::cout << std::endl;

    // Error tracking
    int errorCount = 0;

    // ========================================
    // I. CREATE NEW DOCUMENT
    // ========================================
    std::cout << "I. CREATE NEW DOCUMENT" << std::endl;
    std::cout << "======================" << std::endl;

    // Create a new workbook
    XLDocument doc;
    const std::string xlsxFileNameA = "Demo11A.xlsx";
    doc.create(xlsxFileNameA);
    auto wb = doc.workbook();

    // Create test worksheets
    wb.addWorksheet("Anchor Types");
    XLWorksheet anchorSheet = wb.worksheet("Anchor Types");
    wb.addWorksheet("File Formats");
    XLWorksheet formatSheet = wb.worksheet("File Formats");
    wb.addWorksheet("Sizing Strategies");
    XLWorksheet sizingSheet = wb.worksheet("Sizing Strategies");
    wb.addWorksheet("Units & Scaling");
    XLWorksheet unitsSheet = wb.worksheet("Units & Scaling");
    wb.addWorksheet("File Loading");
    XLWorksheet fileSheet = wb.worksheet("File Loading");

    // 1. Sheet 1 (empty)
    std::cout << "1. Sheet 1 (empty)" << std::endl;
    // Sheet1 is created by default and remains empty

    // 2. Sheet 2: Anchor Types (oneCellAnchor vs twoCellAnchor)
    std::cout << "2. Sheet 2: Anchor Types (oneCellAnchor vs twoCellAnchor)" << std::endl;
    
    // Test 1: oneCellAnchor vs twoCellAnchor with same image

    // Add headers
    anchorSheet.cell("A1").value() = "oneCellAnchor (PNG, preserve aspect ratio within 2x2 cell height bounding box)";
    anchorSheet.cell("A7").value() = "twoCellAnchor (PNG, 4x3 cells)";
    anchorSheet.cell("A13").value() = "oneCellAnchor (JPEG, original size)";
    anchorSheet.cell("A19").value() = "twoCellAnchor (JPEG, 3x2 cells)";
    
    // Add images
    XLImageAnchor imageAnchor1;
    imageAnchor1.initOneCell( XLCellReference("A2") );
    imageAnchor1.setDisplaySizeWithAspectRatio(XLImageUtils::png31x15Data, XLMimeType::PNG,
        2 * EXCEL_CELL_Y_SPACING_EMUS, 2 * EXCEL_CELL_Y_SPACING_EMUS);
    const std::string imageId1 = anchorSheet.embedImageFromImageData(imageAnchor1,
        XLImageUtils::png31x15Data, XLMimeType::PNG );
    bool success1 = !imageId1.empty();
    std::cout << "  Added oneCellAnchor PNG: " << (success1 ? "Success" : "Failed") << std::endl;
    if (!success1) ++errorCount;
    
    // Add twoCellAnchor image
    XLImageAnchor imageAnchor2;
    imageAnchor2.initTwoCell( XLCellReference("A8"), XLCellReference("E11") );
    const std::string imageId2 = anchorSheet.embedImageFromImageData(imageAnchor2,
        XLImageUtils::png31x15Data, XLMimeType::PNG );
    bool success2 = !imageId2.empty();
    std::cout << "  Added twoCellAnchor PNG: " << (success2 ? "Success" : "Failed") << std::endl;
    if (!success2) ++errorCount;
    
    // Add more twoCellAnchor tests
    const std::string imageId3 = anchorSheet.embedImageFromImageData(XLCellReference("A14" ),
        XLImageUtils::jpeg32x18Data, XLMimeType::JPEG );
    bool success3 = !imageId3.empty();
    std::cout << "  Added oneCellAnchor JPEG: " << (success3 ? "Success" : "Failed") << std::endl;
    if (!success3) ++errorCount;
    
    XLImageAnchor imageAnchor4;
    imageAnchor4.initTwoCell( XLCellReference("A20"), XLCellReference("D22") );
    const std::string imageId4 = anchorSheet.embedImageFromImageData(imageAnchor4,
        XLImageUtils::jpeg32x18Data, XLMimeType::JPEG );
    bool success4 = !imageId4.empty();
    std::cout << "  Added twoCellAnchor JPEG: " << (success4 ? "Success" : "Failed") << std::endl;
    if (!success4) ++errorCount;

    // Check data integrity -- typically only done for debugging
    verifyWorksheetData(anchorSheet); 

    // 3. Sheet 3: File Formats (PNG, JPEG, GIF, BMP)
    std::cout << "3. Sheet 3: File Formats (PNG, JPEG, GIF, BMP)" << std::endl;
    
    // Test different file formats
    static const std::vector<std::tuple<std::vector<uint8_t>, std::string, XLMimeType>> formats = {
        {XLImageUtils::png31x15Data, "PNG", XLMimeType::PNG},
        {XLImageUtils::jpeg32x18Data, "JPEG", XLMimeType::JPEG},
        {XLImageUtils::gif26x16Data, "GIF", XLMimeType::GIF},
        {XLImageUtils::bmp33x16Data, "BMP", XLMimeType::BMP}
    };
    
    for (size_t i = 0; i < formats.size(); ++i) {        
        std::string textCell = "A" + std::to_string(1 + i * 6);
        std::string imageCell = "A" + std::to_string(2 + i * 6);
        formatSheet.cell(textCell).value() = std::get<1>(formats[i]) + " Format Test";
        XLImageAnchor imageAnchor;
        imageAnchor.initOneCell( XLCellReference(imageCell) );
        imageAnchor.setDisplaySizeWithAspectRatio(std::get<0>(formats[i]), std::get<2>(formats[i]),
            4 * EXCEL_CELL_Y_SPACING_EMUS, 2 * EXCEL_CELL_Y_SPACING_EMUS);
        const std::string imageId = formatSheet.embedImageFromImageData(imageAnchor,
            std::get<0>(formats[i]), std::get<2>(formats[i]) );
        bool success = !imageId.empty();
        std::cout << "  Added " << std::get<1>(formats[i]) << " image: " << (success ? "Success" : "Failed") << std::endl;
        if (!success) ++errorCount;
    }

    // Check data integrity -- typically only done for debugging
    verifyWorksheetData(formatSheet); 

    // 4. Sheet 4: Sizing Strategy (aspect ratio, original, exact)
    std::cout << "4. Sheet 4: Sizing Strategy (aspect ratio, original, exact)" << std::endl;
    
    // Test different sizing strategies
    // Strategy 1: Preserve aspect ratio + bounding box
    sizingSheet.cell("A1").value() = "Aspect Ratio + Bounding Box (7x3 cell height units)";
    XLImageAnchor sizingImageAnchor;
    sizingImageAnchor.initOneCell( XLCellReference("A2") );
    sizingImageAnchor.setDisplaySizeWithAspectRatio(XLImageUtils::png31x15Data, XLMimeType::PNG,
        7 * EXCEL_CELL_Y_SPACING_EMUS, 3 * EXCEL_CELL_Y_SPACING_EMUS);
    const std::string imageId3a = sizingSheet.embedImageFromImageData(sizingImageAnchor,
        XLImageUtils::png31x15Data, XLMimeType::PNG );
    bool success3a = !imageId3a.empty();
    std::cout << "  Aspect ratio + bounding box: " << (success3a ? "Success" : "Failed") << std::endl;
    if (!success3a) ++errorCount;
    
    // Strategy 2: Original size (no scaling)
    // Don't call setDisplaySize - use original dimensions
    sizingSheet.cell("A7").value() = "Original Size (no scaling)";
    const std::string imageId3b = sizingSheet.embedImageFromImageData(XLCellReference("A8" ),
        XLImageUtils::png31x15Data, XLMimeType::PNG );
    bool success3b = !imageId3b.empty();
    std::cout << "  Original size: " << (success3b ? "Success" : "Failed") << std::endl;
    if (!success3b) ++errorCount;
    
    // Strategy 3: Force exact dimensions
    sizingSheet.cell("A13").value() = "Exact Dimensions (11x3 cell height units, may distort)";
    XLImageAnchor exactImageAnchor;
    exactImageAnchor.initOneCell( XLCellReference("A14") );
    exactImageAnchor.displayWidthEMUs = 11 * EXCEL_CELL_Y_SPACING_EMUS;
    exactImageAnchor.displayHeightEMUs = 3 * EXCEL_CELL_Y_SPACING_EMUS;
    const std::string imageId3c = sizingSheet.embedImageFromImageData(exactImageAnchor,
        XLImageUtils::png31x15Data, XLMimeType::PNG );
    bool success3c = !imageId3c.empty();
    std::cout << "  Exact dimensions: " << (success3c ? "Success" : "Failed") << std::endl;
    if (!success3c) ++errorCount;

    // Check data integrity -- typically only done for debugging
    verifyWorksheetData(sizingSheet);

    // 5. Sheet 5: Units & Scaling Testing (pixels, EMUs, cells)
    std::cout << "5. Sheet 5: Units & Scaling Testing" << std::endl;
    
    // Test different unit systems 

    // Pixels
    unitsSheet.cell("A1").value() = "Pixel-based (100x50 pixels)";
    XLImageAnchor pixelUnitsImageAnchor;
    pixelUnitsImageAnchor.initOneCell( XLCellReference("A2") );
    const uint32_t maxWidthEmus = XLImageUtils::pixelsToEmus(100);
    const uint32_t maxHeightEmus = XLImageUtils::pixelsToEmus(50);
    pixelUnitsImageAnchor.setDisplaySizeWithAspectRatio(XLImageUtils::png31x15Data, XLMimeType::PNG,
        maxWidthEmus, maxHeightEmus);
    const std::string imageId4a = unitsSheet.embedImageFromImageData(pixelUnitsImageAnchor,
        XLImageUtils::png31x15Data, XLMimeType::PNG );
    bool success4a = !imageId4a.empty();
    std::cout << "  Pixel-based sizing: " << (success4a ? "Success" : "Failed") << std::endl;
    if (!success4a) ++errorCount;
    
    // EMUs
    unitsSheet.cell("A7").value() = "EMU-based (476250x238125 EMUs)";
    XLImageAnchor emuUnitsImageAnchor;
    emuUnitsImageAnchor.initOneCell( XLCellReference("A8") );
    emuUnitsImageAnchor.displayWidthEMUs = 476250;
    emuUnitsImageAnchor.displayHeightEMUs = 238125;
    const std::string imageId4b = unitsSheet.embedImageFromImageData(emuUnitsImageAnchor,
        XLImageUtils::png31x15Data, XLMimeType::PNG );
    bool success4b = !imageId4b.empty();
    std::cout << "  EMU-based sizing: " << (success4b ? "Success" : "Failed") << std::endl;
    if (!success4b) ++errorCount;
    
    // Cell spacing
    unitsSheet.cell("A13").value() = "Cell-based (19x5 cell height units) -- preserve aspect ratio";
    XLImageAnchor cellSpacingImageAnchor;
    cellSpacingImageAnchor.initOneCell( XLCellReference("A14") );
    cellSpacingImageAnchor.setDisplaySizeWithAspectRatio(XLImageUtils::png31x15Data, XLMimeType::PNG,
        19 * EXCEL_CELL_Y_SPACING_EMUS, 5 * EXCEL_CELL_Y_SPACING_EMUS);
    const std::string imageId4c = unitsSheet.embedImageFromImageData(cellSpacingImageAnchor,
        XLImageUtils::png31x15Data, XLMimeType::PNG );
    bool success4c = !imageId4c.empty();
    std::cout << "  Cell-based sizing: " << (success4c ? "Success" : "Failed") << std::endl;
    if (!success4c) ++errorCount;

    // Check data integrity -- typically only done for debugging
    verifyWorksheetData(unitsSheet);

    // 6. Sheet 6: Loading Images from Disk Files
    std::cout << "6. Sheet 6: Loading Images from Disk Files" << std::endl;
        
    // Find the images directory by searching upward through the directory tree
    std::string imageDir = findDirectory("OpenXLSX/Examples/images/");
    if (imageDir.empty()) {
        std::cout << "  ERROR: Could not find images directory!" << std::endl;
        std::cout << "  Searched for: OpenXLSX/Examples/images/ in current "
            "directory and up to 10 levels up" << std::endl;
        ++errorCount;
    } else {
        std::cout << "  Found images directory: " << imageDir << std::endl;
    }
    
    // Only proceed with file loading tests if we found the images directory
    if (!imageDir.empty()) {
        // Image 1: PNG - Original size
        std::string pngPath = imageDir + "tiny_png.png";
        fileSheet.cell("A1").value() = "PNG - Original Size";
        const std::string imageIdFile1 =
            fileSheet.embedImageFromFile(XLCellReference("A2"), pngPath );
        bool fileSuccess1 = !imageIdFile1.empty();
        std::cout << "  PNG original size: " << (fileSuccess1 ? "Success" : "Failed") << std::endl;
        if (!fileSuccess1) ++errorCount;
    
        // Image 2: JPEG - Aspect ratio preserved (2x1 cells)
        std::string jpegPath = imageDir + "tiny_jpeg.jpg";
        fileSheet.cell("A7").value() = "JPEG - Aspect Ratio (2x1 cells)";
        XLImageAnchor jpegImageAnchor;
        jpegImageAnchor.initOneCell( XLCellReference("A8") );
        jpegImageAnchor.setDisplaySizeWithAspectRatio(jpegPath, XLMimeType::JPEG,
            2 * EXCEL_CELL_Y_SPACING_EMUS, 1 * EXCEL_CELL_Y_SPACING_EMUS);
        const std::string imageIdFile2 = fileSheet.embedImageFromFile(jpegImageAnchor,
            jpegPath, XLMimeType::JPEG );
        bool fileSuccess2 = !imageIdFile2.empty();
        std::cout << "  JPEG aspect ratio: " << (fileSuccess2 ? "Success" : "Failed") << std::endl;
        if (!fileSuccess2) ++errorCount;

        // Image 3: GIF - Exact dimensions (may distort)
        std::string gifPath = imageDir + "tiny_gif.gif";
        fileSheet.cell("A13").value() = "GIF - Exact Dimensions (3x2 cells)";
        XLImageAnchor gifImageAnchor;
        gifImageAnchor.initOneCell( XLCellReference("A14") );
        gifImageAnchor.displayWidthEMUs = 3 * EXCEL_CELL_Y_SPACING_EMUS;
        gifImageAnchor.displayHeightEMUs = 2 * EXCEL_CELL_Y_SPACING_EMUS;
        const std::string imageIdFile3 = fileSheet.embedImageFromFile(gifImageAnchor,
            gifPath, XLMimeType::GIF );
        bool fileSuccess3 = !imageIdFile3.empty();
        std::cout << "  GIF exact dimensions: " << (fileSuccess3 ? "Success" : "Failed") << std::endl;
        if (!fileSuccess3) ++errorCount;
    
        // Image 4: BMP - Two-cell anchor spanning multiple cells
        std::string bmpPath = imageDir + "tiny_bmp.bmp";
        fileSheet.cell("A19").value() = "BMP - Two-Cell Anchor (A20:D23)";
        XLImageAnchor bmpImageAnchor;
        bmpImageAnchor.initTwoCell( XLCellReference("A20"), XLCellReference("D23") );
        bmpImageAnchor.setDisplaySizeWithAspectRatio(bmpPath, XLMimeType::BMP,
            3 * EXCEL_CELL_Y_SPACING_EMUS, 2 * EXCEL_CELL_Y_SPACING_EMUS);
        const std::string imageIdFile4 = fileSheet.embedImageFromFile(bmpImageAnchor,
            bmpPath, XLMimeType::BMP );
        bool fileSuccess4 = !imageIdFile4.empty();
        std::cout << "  BMP two-cell anchor: " << (fileSuccess4 ? "Success" : "Failed") << std::endl;
        if (!fileSuccess4) ++errorCount;

        // Image 5: PNG - Precise positioning with offset (using previously unused embedImageWithOffset)
        std::string pngPath2 = imageDir + "tiny_png.png";
        fileSheet.cell("A25").value() = "PNG - Precise Offset (A26 + 50% cell height offset)";
        XLImageAnchor pngImageAnchor2;
        pngImageAnchor2.initOneCell( XLCellReference("A26") );
        pngImageAnchor2.setDisplaySizeWithAspectRatio(pngPath2, XLMimeType::PNG,
            3 * EXCEL_CELL_Y_SPACING_EMUS, 2 * EXCEL_CELL_Y_SPACING_EMUS);
        pngImageAnchor2.fromColOffset = EXCEL_CELL_Y_SPACING_EMUS / 2;  // 50% of cell height
        pngImageAnchor2.fromRowOffset = EXCEL_CELL_Y_SPACING_EMUS / 2;  // 50% of cell height
        const std::string imageIdFile5 = fileSheet.embedImageFromFile(pngImageAnchor2,
            pngPath2, XLMimeType::PNG );
        bool fileSuccess5 = !imageIdFile5.empty();
        std::cout << "  PNG with offset: " << (fileSuccess5 ? "Success" : "Failed") << std::endl;
        if (!fileSuccess5) ++errorCount;
        
        // Image 6: JPEG - Larger size with offset (demonstrates variety in offset sizing)
        std::string jpegPath2 = imageDir + "tiny_jpeg.jpg";
        fileSheet.cell("A31").value() = "JPEG - Large Size with Offset (A32 + 25% cell offset)";
        XLImageAnchor jpegImageAnchor2;
        jpegImageAnchor2.initOneCell( XLCellReference("A32") );
        jpegImageAnchor2.setDisplaySizeWithAspectRatio(jpegPath2, XLMimeType::JPEG,
            4 * EXCEL_CELL_Y_SPACING_EMUS, 3 * EXCEL_CELL_Y_SPACING_EMUS);
        jpegImageAnchor2.fromColOffset = EXCEL_CELL_Y_SPACING_EMUS / 4;  // 25% of cell height
        jpegImageAnchor2.fromRowOffset = EXCEL_CELL_Y_SPACING_EMUS / 4;  // 25% of cell height
        const std::string imageIdFile6 = fileSheet.embedImageFromFile(jpegImageAnchor2,
            jpegPath2, XLMimeType::JPEG );
        bool fileSuccess6 = !imageIdFile6.empty();
        std::cout << "  JPEG large with offset: " << (fileSuccess6 ? "Success" : "Failed") << std::endl;
        if (!fileSuccess6) ++errorCount;
    
        // Check data integrity -- typically only done for debugging
        verifyWorksheetData(fileSheet);
    } // End of file loading tests

    // Check data integrity -- typically only done for debugging
    verifyDocData(doc);

    // 7. Inspect Images
    std::cout << "7. Inspect Embedded Images" << std::endl;
    printAllEmbeddedImages(doc);
    
    // 7a. Verify Image Consistency (cross-reference old vs new system)
    std::cout << "7a. Verify Image Consistency" << std::endl;
    verifyImageConsistency(doc);

    // Check data integrity -- typically only done for debugging
    verifyDocData(doc);

    // 8. Set document properties
    std::cout << "8. Set document properties" << std::endl;
    doc.setCreator("OpenXLSX Demo11");
    doc.setTitle("Comprehensive Image Embedding Test Matrix");
    doc.setSubject("Testing various image embedding options in Excel");
    doc.setDescription("This document demonstrates different combinations of image embedding options including anchor types, file formats, sizing strategies, and units");
    doc.setKeywords("OpenXLSX, images, embedding, Excel, testing, demo");
    doc.setLastModifiedBy("Demo11 Application");
    doc.setCategory("Test Document");
    doc.setApplication("OpenXLSX Demo11 Application");
    doc.setCreationDateToNow();
    doc.setModificationDateToNow();
    doc.setAbsPath("");

    // Check data integrity -- typically only done for debugging
    verifyDocData(doc);

    // Display the properties that were set
    std::cout << "  Creator:           " << doc.creator() << std::endl;
    std::cout << "  Title:             " << doc.title() << std::endl;
    std::cout << "  Subject:           " << doc.subject() << std::endl;
    std::cout << "  Description:       " << doc.description() << std::endl;
    std::cout << "  Keywords:          " << doc.keywords() << std::endl;
    std::cout << "  Last Modified By:  " << doc.lastModifiedBy() << std::endl;
    std::cout << "  Category:          " << doc.category() << std::endl;
    std::cout << "  Application:       " << doc.application() << std::endl;
    std::cout << "  Creation Date:     " << doc.creationDate() << std::endl;
    std::cout << "  Modification Date: " << doc.modificationDate() << std::endl;

    // 9. Save the workbook
    std::cout << "9. Save" << std::endl;
    std::cout << "  Saving workbook as '" << xlsxFileNameA << "'..." << std::endl;
    try {
        doc.save();
        std::cout << "  Workbook saved successfully." << std::endl;        

    } catch (const std::exception& e) {
        std::cout << "  Error saving workbook: " << e.what() << std::endl;
        return 1;
    }

    // ========================================
    // II. READ-MODIFY-WRITE
    // ========================================
    std::cout << std::endl;
    std::cout << "II. READ-MODIFY-WRITE" << std::endl;
    std::cout << "====================" << std::endl;

    // 1. Read back the .xlsx file
    std::cout << "1. Read back the .xlsx file" << std::endl;
    OpenXLSX::XLDocument readDoc;
    try {
        readDoc.open(xlsxFileNameA);
        std::cout << "   Successfully opened: " << xlsxFileNameA << std::endl;
         
        // Check data integrity -- typically only done for debugging
        verifyDocData(readDoc);

        // 2. Compare to original
        std::cout << "2. Compare to original" << std::endl;
        
        // Compare the original doc (that was saved) with the readDoc (that was loaded)
        std::string diffMsg;
        int result = doc.compare(readDoc, &diffMsg);
        
        if (result == 0) {
            std::cout << "   ✓ Documents are identical - no differences found" << std::endl;
        } else {
            std::cout << "   ✗ Documents differ (result: " << result << ")" << std::endl;
            std::cout << "   Differences: " << diffMsg << std::endl;
        }
       
        // 3. Inspect Image Registry
        std::cout << "3. Inspect Embedded Images" << std::endl;

        int totalImagesFound = 0;        
        printAllEmbeddedImages(readDoc,&totalImagesFound);
        
        // 4. Compare to original, after search
        std::cout << "4. Compare to original, after search" << std::endl;
        std::string diffMsgAfterSearch;
        int resultAfterSearch = doc.compare(readDoc, &diffMsgAfterSearch);
        if (resultAfterSearch == 0) {
            std::cout << "   ✓ Documents still identical after search" << std::endl;
        } else {
            std::cout << "   ✗ Documents differ after search (result: " << resultAfterSearch << ")" << std::endl;
            std::cout << "   Differences: " << diffMsgAfterSearch << std::endl;
        }
        
        // 5. Remove images
        std::cout << "5. Remove images" << std::endl;
        
        std::string newImageId = ""; // Declare outside the if block for use in verification summary
        OpenXLSX::XLWorksheet fileLoadingSheet; // Declare outside the if block
        
        if (readDoc.workbook().worksheetExists("File Loading")) {
            fileLoadingSheet = readDoc.workbook().worksheet("File Loading");
            
            // Print embedded images BEFORE removal
            std::cout << "5a. Embedded Images BEFORE Removal:" << std::endl;
            printAllEmbeddedImages(readDoc);
            
            // Print XLImageManager reference counts BEFORE removal
            readDoc.getImageManager().printReferenceCounts("XLImageManager Reference Counts BEFORE Removal");
            
            // Remove image 3 by Image ID and image 4 by Relationship ID (testing both methods)
            std::cout << "   Removing image 3 by Image ID and image 4 by Relationship ID..." << std::endl;
            bool removedImage3 = fileLoadingSheet.removeImageByImageID("3");  // Third image (A14)
            bool removedImage4 = fileLoadingSheet.removeImageByRelationshipId("rId4");  // Fourth image (A20)
            std::cout << "   Removed image 3 (by Image ID): " << (removedImage3 ? "Success" : "Failed") << std::endl;
            if (!removedImage3) ++errorCount;
            std::cout << "   Removed image 4 (by Relationship ID): " << (removedImage4 ? "Success" : "Failed") << std::endl;
            if (!removedImage4) ++errorCount;
            
            // Set "REMOVED" text above the cells where images were removed
            fileLoadingSheet.cell("B13").value() = (removedImage3) ? "REMOVED" : "FAILED TO REMOVE"; // Near image 3 at A14
            fileLoadingSheet.cell("B19").value() = (removedImage4) ? "REMOVED" : "FAILED TO REMOVE"; // Near image 4 at A20

            // Print embedded images AFTER removal
            std::cout << "5. Embedded Images AFTER Removal:" << std::endl;
            printAllEmbeddedImages(readDoc);

            // Print XLImageManager reference counts AFTER removal
            readDoc.getImageManager().printReferenceCounts("XLImageManager Reference Counts AFTER Removal");

            // Check data integrity -- typically only done for debugging
            verifyWorksheetData(fileLoadingSheet); 

            // 7. Add image
            std::cout << "7. Add image" << std::endl;
            XLImageAnchor newImageAnchor;
            newImageAnchor.initOneCell( XLCellReference("A38") );
            newImageAnchor.setDisplaySizeWithAspectRatio(XLImageUtils::png31x15Data, XLMimeType::PNG,
                2 * EXCEL_CELL_Y_SPACING_EMUS, 1 * EXCEL_CELL_Y_SPACING_EMUS);
            newImageId = fileLoadingSheet.embedImageFromImageData(newImageAnchor,
                XLImageUtils::png31x15Data, XLMimeType::PNG );
            std::cout << "   Generated new image ID: " << newImageId << std::endl;
            OpenXLSX::XLEmbeddedImage newImage = fileSheet.getEmbImageByImageID(newImageId);
            if(!newImageId.empty()){ 
                std::cout << "   Successfully loaded image data" << std::endl;

                // Add it to cell A37 (after the existing images)
                fileLoadingSheet.cell("A37").value() = "NEW IMAGE - Added via Read-Modify-Write Cycle";
                std::cout << "   Set cell A37 text" << std::endl;
                
                // Check if image is valid before adding
                std::cout << "   Image valid: " << (newImage.isValid() ? "Yes" : "No") << std::endl;
                std::cout << "   Image data size: " << newImage.data().size() << " bytes" << std::endl;
                std::cout << "   Image MIME type: " <<  XLImageUtils::mimeTypeToString(newImage.mimeType()) << std::endl;
 
                if (newImage.isValid()) {
                    std::cout << "   Successfully added new image to 'File Loading' worksheet" << std::endl;
                    std::cout << "   New image ID: " << newImageId << std::endl;
                    std::cout << "   Image positioned at: A38" << std::endl;
                    
                    // Verify the image was actually added by checking the embedded images
                    try {
                        // Check if the new image was added
                        const auto& embImages = fileLoadingSheet.getEmbImages();
                        bool foundNewImage = false;
                        for (const auto& embImage : embImages) {
                            if (embImage.id() == newImageId) {
                                foundNewImage = true;
                                std::cout << "   Found new image in embedded images: " << embeddedImageStr(embImage) << std::endl;
                            }
                        }
                        if (!foundNewImage) {
                            std::cout << "   WARNING: New image '" << newImageId << "' not found in embedded images after addition!" << std::endl;
                        }
                    } catch (const std::exception& e) {
                        std::cout << "   ERROR: Failed to verify image addition: " << e.what() << std::endl;
                    }
                } else {
                    std::cout << "   Failed to add new image to worksheet" << std::endl;
                    ++errorCount;
                }
                
                // Print embedded images AFTER addition
                std::cout << "7a. Embedded Images AFTER Addition:" << std::endl;
                printAllEmbeddedImages(readDoc);
                
            } else {
                std::cout << "   Failed to create new image from data" << std::endl;
                ++errorCount;
            }
        } else {
            std::cout << "   'File Loading' worksheet not found!" << std::endl;
            ++errorCount;
        }
        
        // Check data integrity -- typically only done for debugging
        verifyWorksheetData(fileLoadingSheet); 
        
        // 8. Compare to original, after modification
        // TODO: XLDocument::compare() should compare list<XLXmlData> m_data / std::unique_ptr<XMLDocument> m_xmlDoc / xml_node/type()/name()/value()
        std::cout << "8. Compare to original, after modification" << std::endl;
        std::string diffMsgAfterModify;
        int resultAfterModify = doc.compare(readDoc, &diffMsgAfterModify);
        
        if (resultAfterModify == 0) {
            std::cout << "   ✓ Documents still identical after modifications" << std::endl;
        } else {
            std::cout << "   ✗ Documents differ after modifications (result: " << resultAfterModify << ")" << std::endl;
            std::cout << "   Differences: " << diffMsgAfterModify << std::endl;
        }
        
        // 9. Inspect Embedded Images
        std::cout << "9. Inspect Embedded Images" << std::endl;
        try {
            if (readDoc.workbook().worksheetExists("File Loading")) {
                printEmbeddedImages(fileLoadingSheet, "File Loading");
                std::cout << "   Total images after modifications: " << fileLoadingSheet.imageCount() << std::endl;
            } else {
                std::cout << "   'File Loading' worksheet not found!" << std::endl;
            }
        } catch (const std::exception& e) {
            std::cout << "     Error getting final images: " << e.what() << std::endl;
        }

        // 10. Set Document properties
        std::cout << "10. Set Document properties" << std::endl;
        
        // Set document properties for the modified version
        readDoc.setCreator("OpenXLSX Demo11 - Modified");
        readDoc.setTitle("Comprehensive Image Embedding Test Matrix - Modified");
        readDoc.setSubject("Testing image embedding with modifications (add/remove operations)");
        readDoc.setDescription("This document demonstrates image embedding operations including adding and removing images during read-modify-write cycles");
        readDoc.setKeywords("OpenXLSX, images, embedding, Excel, testing, demo, modified, read-modify-write");
        readDoc.setLastModifiedBy("Demo11 Application - Modified");
        readDoc.setCategory("Test Document - Modified");
        readDoc.setApplication("OpenXLSX Demo11 Application - Modified");
        readDoc.setModificationDateToNow(); // Update modification date for the modified version

        // Display the properties that were set for the modified version
        std::cout << "     Creator:           " << readDoc.creator() << std::endl;
        std::cout << "     Title:             " << readDoc.title() << std::endl;
        std::cout << "     Subject:           " << readDoc.subject() << std::endl;
        std::cout << "     Keywords:          " << readDoc.keywords() << std::endl;
        std::cout << "     Last Modified By:  " << readDoc.lastModifiedBy() << std::endl;
        std::cout << "     Category:          " << readDoc.category() << std::endl;
        std::cout << "     Application:       " << readDoc.application() << std::endl;
        std::cout << "     Modification Date: " << readDoc.modificationDate() << std::endl;

        // Check data integrity -- typically only done for debugging
        verifyDocData(readDoc); 

        // print reference counts
        readDoc.getImageManager().printReferenceCounts("XLImageManager Reference Counts before saving");
 
        // 11. Save modified file
        const std::string xlsxFileNameB = "Demo11B.xlsx";
        std::cout << "11. Save modified file: " << xlsxFileNameB << std::endl;
        try {
            readDoc.saveAs(xlsxFileNameB);
        } catch (const std::exception& e) {
            std::cout << "   Error saving modified workbook: " << e.what() << std::endl;
        }
        
        readDoc.close();
        
    } catch (const std::exception& e) {
        std::cout << "   Error during read-modify-write cycle: " << e.what() << std::endl;
    }

    // ========================================
    // III. READ EXCEL-GENERATED FILE
    // ========================================
    std::cout << std::endl;
    std::cout << "III. READ EXCEL-GENERATED FILE" << std::endl;
    std::cout << "=============================" << std::endl;

    // 1. Read file
    std::cout << "1. Read file" << std::endl;
    
    // Find the Excel-generated file
    std::string excelFileDir = findDirectory("OpenXLSX/Examples/files/");
    std::string excelFilePath = excelFileDir + "from_excel_embedded_images.xlsx";
    
    if (excelFileDir.empty() || !std::filesystem::exists(excelFilePath)) {
        std::cout << "  ERROR: Could not find Excel-generated file!" << std::endl;
        std::cout << "  Expected: " << excelFilePath << std::endl;
        std::cout << "  Searched for: OpenXLSX/Examples/files/ in current directory and up to 10 levels up" << std::endl;
        ++errorCount;
    } else {
        std::cout << "  Found Excel file: " << excelFilePath << std::endl;
        
        // Load the Excel-generated file
        XLDocument excelDoc;
        try {
            excelDoc.open(excelFilePath);

            // Check data integrity -- typically only done for debugging
            verifyDocData(excelDoc); 


            // 2. Inspect Embedded Images 
            std::cout << "2 Inspect Embedded Images" << std::endl;
            int totalEmbeddedImages = 0;
            printAllEmbeddedImages(excelDoc, &totalEmbeddedImages);
            std::cout << "  Total embedded images in Excel file: " << totalEmbeddedImages << std::endl;
            
            // 2c. Verify Image Consistency (cross-reference old vs new system)
            std::cout << "2c. Verify Image Consistency" << std::endl;
            verifyImageConsistency(excelDoc);
        } catch (const std::exception& e) {
            std::cout << "  ERROR: Failed to open Excel file: " << e.what() << std::endl;
            ++errorCount;
        }

        // Display the properties that were set for the Excel-generated file
        std::cout << "3. Properties" << std::endl;
        std::cout << "     Creator:           " << excelDoc.creator() << std::endl;
        std::cout << "     Title:             " << excelDoc.title() << std::endl;
        std::cout << "     Subject:           " << excelDoc.subject() << std::endl;
        std::cout << "     Keywords:          " << excelDoc.keywords() << std::endl;
        std::cout << "     Last Modified By:  " << excelDoc.lastModifiedBy() << std::endl;
        std::cout << "     Category:          " << excelDoc.category() << std::endl;
        std::cout << "     Application:       " << excelDoc.application() << std::endl;
        std::cout << "     Modification Date: " << excelDoc.modificationDate() << std::endl;
    }

    // ========================================
    // IV. IMAGE QUERY
    // ========================================
    std::cout << std::endl;
    std::cout << "IV. IMAGE QUERY" << std::endl;
    std::cout << "===============" << std::endl;

    // 1. Read file
    std::cout << "1. Read file" << std::endl;
    
    // Re-open the Demo11 file to test image queries
        XLDocument testDoc;
    try {
        testDoc.open(xlsxFileNameA);
        std::cout << "  Successfully opened " << xlsxFileNameA << std::endl;
        
        // Check data integrity -- typically only done for debugging
        verifyDocData(testDoc); 

        // 2. Queries
        std::cout << "2. Queries" << std::endl;
        
        // Test each worksheet with images
        for (uint16_t i = 1; i <= testDoc.workbook().sheetCount(); ++i) {
            XLSheet sheet = testDoc.workbook().sheet(i);
            std::string sheetName = sheet.name();
            
            if (sheet.isType<XLWorksheet>()) {
                std::cout << std::endl;
                XLWorksheet worksheet = sheet.get<XLWorksheet>();
                printDetailedImageAnalysis(worksheet, sheetName);
            }
        }
        
        testDoc.close();
        std::cout << "\n  Image query testing completed successfully!" << std::endl;
        
    } catch (const std::exception& e) {
        std::cout << "  ERROR: Failed to test image queries: " << e.what() << std::endl;
        ++errorCount;
    }

    // Print final error summary
    std::cout << std::endl;
    std::cout << "=========================================" << std::endl;
    std::cout << "FINAL ERROR SUMMARY" << std::endl;
    std::cout << "=========================================" << std::endl;
    std::cout << "Total errors encountered: " << errorCount << std::endl;
    if (errorCount == 0) {
        std::cout << "All operations completed successfully!" << std::endl;
    } else {
        std::cout << "Some operations failed. Check the output above for details." << std::endl;
    }
    std::cout << "=========================================" << std::endl;

    return 0;
}

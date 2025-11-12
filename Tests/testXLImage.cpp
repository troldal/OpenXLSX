//
// Created by OpenXLSX Image Test Suite
//

#include <OpenXLSX.hpp>
#include <catch.hpp>
#include <fstream>
#include <random>
#include <algorithm>
#include <map>
#include <cmath>

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
            // Use shared tiny PNG image data from XLImageUtils
            const std::string imageId1 = worksheet1.embedImageFromImageData(
                XLCellReference(2,2), XLImageUtils::red1x1PNGData, XLMimeType::PNG );
            const std::string imageId2 = worksheet1.embedImageFromImageData(
                XLCellReference(5,3), XLImageUtils::red1x1PNGData, XLMimeType::PNG );
            const std::string imageId3 = worksheet2.embedImageFromImageData(
                XLCellReference(1,1), XLImageUtils::red1x1PNGData, XLMimeType::PNG );
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

    // -----------------------------------------------------------------
    // Additional sections consolidated under the XLEmbeddedImage test
    // -----------------------------------------------------------------

    SECTION("XLImageUtils - brief test") {
        // Detect and validate
        const XLMimeType mt = XLImageUtils::detectMimeType(XLImageUtils::red1x1PNGData);
        REQUIRE(mt == XLMimeType::PNG);
        REQUIRE(XLImageUtils::validateImageData(XLImageUtils::red1x1PNGData, mt));

        // MIME/content-type round-trip
        const XLContentType ct = XLImageUtils::mimeTypeToContentType(mt);
        const XLMimeType mt2 = XLImageUtils::contentTypeToMimeType(ct);
        REQUIRE(mt2 == XLMimeType::PNG);

        // Extension mapping
        REQUIRE(XLImageUtils::extensionToMimeType(".png") == XLMimeType::PNG);
        REQUIRE(XLImageUtils::mimeTypeToExtension(XLMimeType::PNG) == ".png");

        // EMU conversion sanity
        const uint32_t px = 31;
        const uint32_t emu = XLImageUtils::pixelsToEmus(px);
        REQUIRE(emu > 0u);
        REQUIRE(XLImageUtils::emusToPixels(emu) >= 1u);
    }

    SECTION("XLImageUtils - full test") {
        // Test data - using pointers to avoid vector of references
        struct TestImage {
            const std::vector<uint8_t>* data;
            XLMimeType mimeType;
        };
        const std::vector<TestImage> testImages = {
            {&XLImageUtils::red1x1PNGData, XLMimeType::PNG},
            {&XLImageUtils::png31x15Data, XLMimeType::PNG},
            {&XLImageUtils::jpeg32x18Data, XLMimeType::JPEG},
            {&XLImageUtils::bmp33x16Data, XLMimeType::BMP},
            {&XLImageUtils::gif26x16Data, XLMimeType::GIF}
        };

        // Test detectMimeType() - Convert each test image data vector to mimetype
        REQUIRE(XLImageUtils::detectMimeType(XLImageUtils::red1x1PNGData) == XLMimeType::PNG);
        REQUIRE(XLImageUtils::detectMimeType(XLImageUtils::png31x15Data) == XLMimeType::PNG);
        REQUIRE(XLImageUtils::detectMimeType(XLImageUtils::jpeg32x18Data) == XLMimeType::JPEG);
        REQUIRE(XLImageUtils::detectMimeType(XLImageUtils::bmp33x16Data) == XLMimeType::BMP);
        REQUIRE(XLImageUtils::detectMimeType(XLImageUtils::gif26x16Data) == XLMimeType::GIF);
        
        // Test with invalid data
        const std::vector<uint8_t> invalidData = {0x00, 0x01, 0x02, 0x03};
        REQUIRE(XLImageUtils::detectMimeType(invalidData) == XLMimeType::Unknown);

        // Test mimeTypeToString() and stringToMimeType() - Convert each to and from MIME type string
        REQUIRE(XLImageUtils::mimeTypeToString(XLMimeType::PNG) == "image/png");
        REQUIRE(XLImageUtils::mimeTypeToString(XLMimeType::JPEG) == "image/jpeg");
        REQUIRE(XLImageUtils::mimeTypeToString(XLMimeType::BMP) == "image/bmp");
        REQUIRE(XLImageUtils::mimeTypeToString(XLMimeType::GIF) == "image/gif");
        REQUIRE(XLImageUtils::mimeTypeToString(XLMimeType::Unknown) == "");

        REQUIRE(XLImageUtils::stringToMimeType("image/png") == XLMimeType::PNG);
        REQUIRE(XLImageUtils::stringToMimeType("image/jpeg") == XLMimeType::JPEG);
        REQUIRE(XLImageUtils::stringToMimeType("image/bmp") == XLMimeType::BMP);
        REQUIRE(XLImageUtils::stringToMimeType("image/gif") == XLMimeType::GIF);
        
        // Round-trip tests
        for (const auto& img : testImages) {
            const std::string mimeStr = XLImageUtils::mimeTypeToString(img.mimeType);
            REQUIRE(!mimeStr.empty());
            REQUIRE(XLImageUtils::stringToMimeType(mimeStr) == img.mimeType);
        }

        // Convert random garbage MIME type string to mime type
        REQUIRE(XLImageUtils::stringToMimeType("garbage/string") == XLMimeType::Unknown);
        REQUIRE(XLImageUtils::stringToMimeType("") == XLMimeType::Unknown);
        REQUIRE(XLImageUtils::stringToMimeType("image/unknown") == XLMimeType::Unknown);
        REQUIRE(XLImageUtils::stringToMimeType("text/plain") == XLMimeType::Unknown);

        // Test mimeTypeToExtension() and extensionToMimeType() - Convert each to and from file extension
        REQUIRE(XLImageUtils::mimeTypeToExtension(XLMimeType::PNG) == ".png");
        REQUIRE(XLImageUtils::mimeTypeToExtension(XLMimeType::JPEG) == ".jpg");
        REQUIRE(XLImageUtils::mimeTypeToExtension(XLMimeType::BMP) == ".bmp");
        REQUIRE(XLImageUtils::mimeTypeToExtension(XLMimeType::GIF) == ".gif");
        REQUIRE(XLImageUtils::mimeTypeToExtension(XLMimeType::Unknown) == "");

        REQUIRE(XLImageUtils::extensionToMimeType(".png") == XLMimeType::PNG);
        REQUIRE(XLImageUtils::extensionToMimeType(".jpg") == XLMimeType::JPEG);
        REQUIRE(XLImageUtils::extensionToMimeType(".jpeg") == XLMimeType::JPEG);
        REQUIRE(XLImageUtils::extensionToMimeType(".bmp") == XLMimeType::BMP);
        REQUIRE(XLImageUtils::extensionToMimeType(".gif") == XLMimeType::GIF);
        
        // Round-trip tests
        for (const auto& img : testImages) {
            const std::string ext = XLImageUtils::mimeTypeToExtension(img.mimeType);
            REQUIRE(!ext.empty());
            REQUIRE(XLImageUtils::extensionToMimeType(ext) == img.mimeType);
        }

        // Convert random garbage extension to mime type
        REQUIRE(XLImageUtils::extensionToMimeType(".xyz") == XLMimeType::Unknown);
        REQUIRE(XLImageUtils::extensionToMimeType(".txt") == XLMimeType::Unknown);
        REQUIRE(XLImageUtils::extensionToMimeType("") == XLMimeType::Unknown);
        REQUIRE(XLImageUtils::extensionToMimeType("noextension") == XLMimeType::Unknown);
        REQUIRE(XLImageUtils::extensionToMimeType("png") == XLMimeType::PNG); // without dot should still work

        // Test mimeTypeToContentType() and contentTypeToMimeType() - Convert each to and from XLContentType
        REQUIRE(XLImageUtils::mimeTypeToContentType(XLMimeType::PNG) == XLContentType::ImagePNG);
        REQUIRE(XLImageUtils::mimeTypeToContentType(XLMimeType::JPEG) == XLContentType::ImageJPEG);
        REQUIRE(XLImageUtils::mimeTypeToContentType(XLMimeType::BMP) == XLContentType::ImageBMP);
        REQUIRE(XLImageUtils::mimeTypeToContentType(XLMimeType::GIF) == XLContentType::ImageGIF);
        REQUIRE(XLImageUtils::mimeTypeToContentType(XLMimeType::Unknown) == XLContentType::Unknown);

        REQUIRE(XLImageUtils::contentTypeToMimeType(XLContentType::ImagePNG) == XLMimeType::PNG);
        REQUIRE(XLImageUtils::contentTypeToMimeType(XLContentType::ImageJPEG) == XLMimeType::JPEG);
        REQUIRE(XLImageUtils::contentTypeToMimeType(XLContentType::ImageBMP) == XLMimeType::BMP);
        REQUIRE(XLImageUtils::contentTypeToMimeType(XLContentType::ImageGIF) == XLMimeType::GIF);
        REQUIRE(XLImageUtils::contentTypeToMimeType(XLContentType::Unknown) == XLMimeType::Unknown);

        // Round-trip tests
        for (const auto& img : testImages) {
            const XLContentType ct = XLImageUtils::mimeTypeToContentType(img.mimeType);
            REQUIRE(ct != XLContentType::Unknown);
            REQUIRE(XLImageUtils::contentTypeToMimeType(ct) == img.mimeType);
        }

        // Test pixelsToEmus() and emusToPixels() - Use Catch RNG to pick random pixel value, convert back and forth
        auto& rng = Catch::rng();
        std::uniform_int_distribution<uint32_t> pixelDist(1, 1000);
        
        for (int i = 0; i < 10; ++i) {
            // Generate random pixel value between 1 and 1000
            const uint32_t pixels = pixelDist(rng);

            const uint32_t emus = XLImageUtils::pixelsToEmus(pixels);
            REQUIRE(emus > 0u);
            
            const uint32_t pixelsBack = XLImageUtils::emusToPixels(emus);
            // Allow for some rounding error (within 1 pixel)
            REQUIRE((pixelsBack == pixels || pixelsBack == pixels - 1 || pixelsBack == pixels + 1));
        }

        // Use Catch RNG to pick random EMU value, convert back and forth
        std::uniform_int_distribution<uint32_t> emuDist(9525, 9525000); // 1-1000 pixels
        
        for (int i = 0; i < 10; ++i) {
            // Generate random EMU value between 9525 and 9525000 (1-1000 pixels)
            const uint32_t emus = emuDist(rng);
            
            const uint32_t pixels = XLImageUtils::emusToPixels(emus);
            REQUIRE(pixels >= 1u);
            
            const uint32_t emusBack = XLImageUtils::pixelsToEmus(pixels);
            // Allow for some rounding error
            const uint32_t diff = (emusBack > emus) ? (emusBack - emus) : (emus - emusBack);
            REQUIRE(diff <= 9525u); // Within 1 pixel
        }

        // Test pointsToEmus()
        REQUIRE(XLImageUtils::pointsToEmus(1.0) == 12700u);
        REQUIRE(XLImageUtils::pointsToEmus(0.0) == 0u);
        REQUIRE(XLImageUtils::pointsToEmus(12.0) == 152400u); // 12 * 12700
        
        // Use Catch RNG to pick random point values
        std::uniform_real_distribution<double> pointDist(0.0, 100.0);
        
        for (int i = 0; i < 10; ++i) {
            const double points = pointDist(rng); // 0 to 100 points
            const uint32_t emus = XLImageUtils::pointsToEmus(points);
            REQUIRE(emus == static_cast<uint32_t>(points * 12700));
        }

        // Test calculateAspectRatio() - Use Catch RNG to pick random width and height values
        std::uniform_int_distribution<uint32_t> dimensionDist(1, 999);
        
        for (int i = 0; i < 10; ++i) {
            const uint32_t width = dimensionDist(rng);
            const uint32_t height = dimensionDist(rng);
            
            const double aspectRatio = XLImageUtils::calculateAspectRatio(width, height);
            REQUIRE(aspectRatio > 0.0);
            REQUIRE(aspectRatio == static_cast<double>(width) / static_cast<double>(height));
        }

        // Test edge cases for calculateAspectRatio()
        REQUIRE(XLImageUtils::calculateAspectRatio(100, 100) == 1.0);
        REQUIRE(XLImageUtils::calculateAspectRatio(200, 100) == 2.0);
        REQUIRE(XLImageUtils::calculateAspectRatio(100, 200) == 0.5);
        REQUIRE(XLImageUtils::calculateAspectRatio(0, 100) == 0.0); // height == 0 case

        // Test getImageDimensions() - Get image dimensions from each of the test image data vectors
        const auto dims1x1 = XLImageUtils::getImageDimensions(XLImageUtils::red1x1PNGData, XLMimeType::PNG);
        REQUIRE(dims1x1.first == 1u);
        REQUIRE(dims1x1.second == 1u);

        const auto dims31x15 = XLImageUtils::getImageDimensions(XLImageUtils::png31x15Data, XLMimeType::PNG);
        REQUIRE(dims31x15.first == 31u);
        REQUIRE(dims31x15.second == 15u);

        const auto dims32x18 = XLImageUtils::getImageDimensions(XLImageUtils::jpeg32x18Data, XLMimeType::JPEG);
        REQUIRE(dims32x18.first == 32u);
        REQUIRE(dims32x18.second == 18u);

        const auto dims33x16 = XLImageUtils::getImageDimensions(XLImageUtils::bmp33x16Data, XLMimeType::BMP);
        REQUIRE(dims33x16.first == 33u);
        REQUIRE(dims33x16.second == 16u);

        const auto dims26x16 = XLImageUtils::getImageDimensions(XLImageUtils::gif26x16Data, XLMimeType::GIF);
        REQUIRE(dims26x16.first == 26u);
        REQUIRE(dims26x16.second == 16u);

        // Test getImageDimensions() with auto-detection
        const auto dims1x1Auto = XLImageUtils::getImageDimensions(XLImageUtils::red1x1PNGData);
        REQUIRE(dims1x1Auto.first == 1u);
        REQUIRE(dims1x1Auto.second == 1u);

        const auto dims31x15Auto = XLImageUtils::getImageDimensions(XLImageUtils::png31x15Data);
        REQUIRE(dims31x15Auto.first == 31u);
        REQUIRE(dims31x15Auto.second == 15u);

        const auto dims32x18Auto = XLImageUtils::getImageDimensions(XLImageUtils::jpeg32x18Data);
        REQUIRE(dims32x18Auto.first == 32u);
        REQUIRE(dims32x18Auto.second == 18u);

        // Test with invalid data
        const auto dimsInvalid = XLImageUtils::getImageDimensions(invalidData);
        REQUIRE(dimsInvalid.first == 0u);
        REQUIRE(dimsInvalid.second == 0u);

        // Test isValidImageSize()
        REQUIRE(XLImageUtils::isValidImageSize(1, 1) == true);
        REQUIRE(XLImageUtils::isValidImageSize(100, 100) == true);
        REQUIRE(XLImageUtils::isValidImageSize(10000, 10000) == true);
        REQUIRE(XLImageUtils::isValidImageSize(10001, 10000) == false);
        REQUIRE(XLImageUtils::isValidImageSize(10000, 10001) == false);
        REQUIRE(XLImageUtils::isValidImageSize(0, 100) == false);
        REQUIRE(XLImageUtils::isValidImageSize(100, 0) == false);
        REQUIRE(XLImageUtils::isValidImageSize(0, 0) == false);

        // Test validateImageData()
        for (const auto& img : testImages) {
            REQUIRE(XLImageUtils::validateImageData(*img.data, img.mimeType) == true);
            // Test with wrong MIME type
            if (img.mimeType != XLMimeType::PNG) {
                REQUIRE(XLImageUtils::validateImageData(*img.data, XLMimeType::PNG) == false);
            }
        }

        // Test with invalid data
        REQUIRE(XLImageUtils::validateImageData(invalidData, XLMimeType::PNG) == false);
        REQUIRE(XLImageUtils::validateImageData(invalidData, XLMimeType::JPEG) == false);
        REQUIRE(XLImageUtils::validateImageData(std::vector<uint8_t>(), XLMimeType::PNG) == false);

        // Test imageDataHash()
        const size_t hash1 = XLImageUtils::imageDataHash(XLImageUtils::red1x1PNGData);
        const size_t hash2 = XLImageUtils::imageDataHash(XLImageUtils::red1x1PNGData);
        REQUIRE(hash1 == hash2); // Same data should produce same hash

        const size_t hash3 = XLImageUtils::imageDataHash(XLImageUtils::png31x15Data);
        REQUIRE(hash1 != hash3); // Different data should produce different hash

        // Test imageDataStrHash()
        std::string imageDataStr;
        imageDataStr.insert(imageDataStr.end(), XLImageUtils::red1x1PNGData.begin(), XLImageUtils::red1x1PNGData.end());
        const size_t hashStr = XLImageUtils::imageDataStrHash(imageDataStr);
        REQUIRE(hash1 == hashStr); // Same data in different formats should produce same hash

        // Test with empty data
        REQUIRE(XLImageUtils::imageDataHash(std::vector<uint8_t>()) == XLImageUtils::imageDataStrHash(""));

        // Test matchesElementName()
        REQUIRE(XLImageUtils::matchesElementName("xdr:oneCellAnchor", "oneCellAnchor") == true);
        REQUIRE(XLImageUtils::matchesElementName("oneCellAnchor", "oneCellAnchor") == true);
        REQUIRE(XLImageUtils::matchesElementName("xdr:twoCellAnchor", "twoCellAnchor") == true);
        REQUIRE(XLImageUtils::matchesElementName("twoCellAnchor", "twoCellAnchor") == true);
        REQUIRE(XLImageUtils::matchesElementName("xdr:oneCellAnchor", "twoCellAnchor") == false);
        REQUIRE(XLImageUtils::matchesElementName("oneCellAnchor", "twoCellAnchor") == false);
        REQUIRE(XLImageUtils::matchesElementName("", "") == true);
        REQUIRE(XLImageUtils::matchesElementName("prefix:name", "name") == true);
        REQUIRE(XLImageUtils::matchesElementName("prefix:name", "prefix") == false);
    }

    SECTION("XLImageManager - brief test") {
        XLDocument doc; doc.create("./ut_mgr.xlsx");
        auto &mgr = doc.getImageManager();

        const size_t before = mgr.getImageCount();
        auto img1 = mgr.findOrAddImage("",
                                       XLImageUtils::red1x1PNGData,
                                       XLImageUtils::mimeTypeToContentType(XLMimeType::PNG));
        auto img2 = mgr.findOrAddImage("",
                                       XLImageUtils::red1x1PNGData,
                                       XLImageUtils::mimeTypeToContentType(XLMimeType::PNG));
        REQUIRE(img1.get() != nullptr);
        REQUIRE(img1.get() == img2.get());
        REQUIRE(mgr.getImageCount() == before + 1);

        doc.save(); doc.close(); std::remove("./ut_mgr.xlsx");
    }

    SECTION("XLImageManager/XLImage - full test") {
        // Create empty document
        XLDocument doc;
        doc.create("./ut_mgr_full.xlsx");
        auto &mgr = doc.getImageManager();

        // Test initial state
        REQUIRE(mgr.getImageCount() == 0);
        REQUIRE(mgr.begin() == mgr.end());
        std::string dbgMsg;
        int verifyDataResult = mgr.verifyData(&dbgMsg);
        REQUIRE(verifyDataResult== EXIT_SUCCESS);

        if(verifyDataResult== EXIT_SUCCESS){
            // Add two worksheets "SheetA", "SheetB"
            doc.workbook().worksheet(1).setName("SheetA");
            doc.workbook().addWorksheet("SheetB");
            XLWorksheet sheetA = doc.workbook().worksheet("SheetA");
            XLWorksheet sheetB = doc.workbook().worksheet("SheetB");

            // Map to track image hash -> image data from XLImageUtils
            std::map<size_t, const std::vector<uint8_t>*> hashToImageData;
        
            // Check vectors for each worksheet
            XLEmbImgVec sheetAImagesCheck;
            XLEmbImgVec sheetBImagesCheck;

            // Add two images to Worksheet "SheetA"
            // Image 1: red1x1PNGData (will be shared with SheetB)
            const std::string imageIdA1 = sheetA.embedImageFromImageData(
                XLCellReference(1, 1), XLImageUtils::red1x1PNGData, XLMimeType::PNG);
            REQUIRE(!imageIdA1.empty());
            const auto& embA1 = sheetA.getEmbImageByImageID(imageIdA1);
            sheetAImagesCheck.push_back(embA1);
            size_t hashA1 = XLImageUtils::imageDataHash(XLImageUtils::red1x1PNGData);
            hashToImageData[hashA1] = &XLImageUtils::red1x1PNGData;

            // Image 2: png31x15Data (unique to SheetA)
            const std::string imageIdA2 = sheetA.embedImageFromImageData(
                XLCellReference(2, 2), XLImageUtils::png31x15Data, XLMimeType::PNG);
            REQUIRE(!imageIdA2.empty());
            const auto& embA2 = sheetA.getEmbImageByImageID(imageIdA2);
            sheetAImagesCheck.push_back(embA2);
            size_t hashA2 = XLImageUtils::imageDataHash(XLImageUtils::png31x15Data);
            hashToImageData[hashA2] = &XLImageUtils::png31x15Data;

            // Add three images to Worksheet "SheetB"
            // Image 1: red1x1PNGData (shared with SheetA)
            const std::string imageIdB1 = sheetB.embedImageFromImageData(
                XLCellReference(1, 1), XLImageUtils::red1x1PNGData, XLMimeType::PNG);
            REQUIRE(!imageIdB1.empty());
            const auto& embB1 = sheetB.getEmbImageByImageID(imageIdB1);
            sheetBImagesCheck.push_back(embB1);

            // Image 2: jpeg32x18Data (unique to SheetB)
            const std::string imageIdB2 = sheetB.embedImageFromImageData(
                XLCellReference(3, 3), XLImageUtils::jpeg32x18Data, XLMimeType::JPEG);
            REQUIRE(!imageIdB2.empty());
            const auto& embB2 = sheetB.getEmbImageByImageID(imageIdB2);
            sheetBImagesCheck.push_back(embB2);
            size_t hashB2 = XLImageUtils::imageDataHash(XLImageUtils::jpeg32x18Data);
            hashToImageData[hashB2] = &XLImageUtils::jpeg32x18Data;

            // Image 3: bmp33x16Data (unique to SheetB)
            const std::string imageIdB3 = sheetB.embedImageFromImageData(
                XLCellReference(4, 4), XLImageUtils::bmp33x16Data, XLMimeType::BMP);
            REQUIRE(!imageIdB3.empty());
            const auto& embB3 = sheetB.getEmbImageByImageID(imageIdB3);
            sheetBImagesCheck.push_back(embB3);
            size_t hashB3 = XLImageUtils::imageDataHash(XLImageUtils::bmp33x16Data);
            hashToImageData[hashB3] = &XLImageUtils::bmp33x16Data;

            // Test: mgr.getImageCount() == 4 (four unique images)
            REQUIRE(mgr.getImageCount() == 4);
            REQUIRE(mgr.begin() != mgr.end());

            // Get embedded images for each sheet
            auto sheetAImagesPtr = mgr.getEmbImgsForSheetName("SheetA");
            auto sheetBImagesPtr = mgr.getEmbImgsForSheetName("SheetB");
            REQUIRE(sheetAImagesPtr != nullptr);
            REQUIRE(sheetBImagesPtr != nullptr);
        
            XLEmbImgVec sheetAImages = *sheetAImagesPtr;
            XLEmbImgVec sheetBImages = *sheetBImagesPtr;

            // Sort vectors by image ID for comparison
            auto sortByImageId = [](const XLEmbeddedImage& a, const XLEmbeddedImage& b) {
                return a.id() < b.id();
            };
            std::sort(sheetAImages.begin(), sheetAImages.end(), sortByImageId);
            std::sort(sheetBImages.begin(), sheetBImages.end(), sortByImageId);
            std::sort(sheetAImagesCheck.begin(), sheetAImagesCheck.end(), sortByImageId);
            std::sort(sheetBImagesCheck.begin(), sheetBImagesCheck.end(), sortByImageId);

            // Test: sheetAImages == sheetAImagesCheck
            REQUIRE(sheetAImages.size() == sheetAImagesCheck.size());
            for (size_t i = 0; i < sheetAImages.size(); ++i) {
                REQUIRE(sheetAImages[i].id() == sheetAImagesCheck[i].id());
            }

            // Test: sheetBImages == sheetBImagesCheck
            REQUIRE(sheetBImages.size() == sheetBImagesCheck.size());
            for (size_t i = 0; i < sheetBImages.size(); ++i) {
                REQUIRE(sheetBImages[i].id() == sheetBImagesCheck[i].id());
            }

            // Test: mgr.verifyData() == EXIT_SUCCESS
            dbgMsg.clear();
            REQUIRE(mgr.verifyData(&dbgMsg) == EXIT_SUCCESS);

            // Iterate through all images and test properties
            for (auto imageItr = mgr.begin(); imageItr != mgr.end(); ++imageItr) {
                XLImageShPtr image = *imageItr;
                REQUIRE(image != nullptr);

                // Test: mgr.findImageByImageData(image->getImageData())
                std::vector<uint8_t> imageData = image->getImageData();
                XLImageShPtr foundByData = mgr.findImageByImageData(imageData);
                REQUIRE(foundByData != nullptr);
                REQUIRE(foundByData.get() == image.get());

                // Test: mgr.findImageByFilePackagePath(image->filePackagePath())
                XLImageShPtr foundByPath = mgr.findImageByFilePackagePath(image->filePackagePath());
                REQUIRE(foundByPath != nullptr);
                REQUIRE(foundByPath.get() == image.get());

                // Test: image->parentDoc() == doc
                REQUIRE(image->parentDoc() == &doc);

                // Get hash and look up image data from map
                size_t hash = image->imageDataHash();
                auto it = hashToImageData.find(hash);
                REQUIRE(it != hashToImageData.end());
                const std::vector<uint8_t>* expectedData = it->second;

                // Test: image->imageDataHash()
                REQUIRE(image->imageDataHash() == hash);
                REQUIRE(image->imageDataHash() == XLImageUtils::imageDataHash(*expectedData));

                // Test: image->mimeType()
                XLMimeType expectedMimeType = XLImageUtils::detectMimeType(*expectedData);
                REQUIRE(image->mimeType() == expectedMimeType);

                // Test: image->extension()
                std::string expectedExt = XLImageUtils::mimeTypeToExtension(expectedMimeType);
                REQUIRE(image->extension() == expectedExt);

                // Test: image->filePackagePath()
                REQUIRE(!image->filePackagePath().empty());
                REQUIRE(image->filePackagePath().find("xl/media/") == 0);

                // Test: image->widthPixels() and image->heightPixels()
                auto expectedDims = XLImageUtils::getImageDimensions(*expectedData, expectedMimeType);
                REQUIRE(image->widthPixels() == expectedDims.first);
                REQUIRE(image->heightPixels() == expectedDims.second);

                // Test: image->verifyData() == EXIT_SUCCESS
                dbgMsg.clear();
                REQUIRE(image->verifyData(&dbgMsg) == EXIT_SUCCESS);
            }

            // Remove sheetB
            doc.workbook().deleteSheet("SheetB");

            // Test: mgr.getImageCount() == 4 (images still exist until prune)
            REQUIRE(mgr.getImageCount() == 4);
        }

        // Save document, triggering prune()
        doc.save();

        // Test: mgr.getImageCount() == 2 after prune
        REQUIRE(mgr.getImageCount() == 2); // red1x1PNGData and png31x15Data

        doc.close();
        std::remove("./ut_mgr_full.xlsx");
    }


    SECTION("XLWorksheet - embed image -- brief test") {
        XLDocument doc; doc.create("./ut_ws.xlsx");
        auto ws = doc.workbook().worksheet(1);

        const size_t before = ws.imageCount();
        const std::string id = ws.embedImageFromImageData(XLCellReference(2, 2),
                                                          XLImageUtils::red1x1PNGData,
                                                          XLMimeType::PNG);
        REQUIRE(!id.empty());
        REQUIRE(ws.imageCount() == before + 1);
        const auto &emb = ws.getEmbImageByImageID(id);
        REQUIRE(!emb.isEmpty());
        // XLCellReference(2,2) -> row 2, col 2 => "B2"
        REQUIRE(emb.anchorCell() == std::string("B2"));

        // Remove and verify count restored
        REQUIRE(ws.removeImageByImageID(id));
        REQUIRE(ws.imageCount() == before);

        doc.save(); doc.close(); std::remove("./ut_ws.xlsx");
    }


    SECTION("XLImageAnchor -- embed image -- full test") {
/*
        XLImageAnchor::XLImageAnchor()
        XLImageAnchor::reset();
        XLImageAnchor::initOneCell();
        XLImageAnchor::initTwoCell();
        XLImageAnchor::getFromCellReference() const;
        XLImageAnchor::setFromCellReference( const XLCellReference& fromCellRef );
        XLImageAnchor::getToCellReference() const;
        XLImageAnchor::setToCellReference();
        XLImageAnchor::setDisplaySizeWithAspectRatio();
        XLImageAnchor::setDisplaySizeWithAspectRatio();
        XLImageAnchor::setDisplaySizeWithAspectRatio();
        XLImageAnchor::isTwoCell()
        XLImageAnchor::getDisplayDimensions()
        XLImageAnchor::getActualDimensions()
        XLImageAnchor::setActualDimensions()
        XLImageAnchor::getDisplayDimensionsPixels()
        XLImageAnchor::getActualDimensionsPixels()
*/
        auto& rng = Catch::rng();
        std::uniform_int_distribution<uint32_t> rowDist(1, 200);
        std::uniform_int_distribution<uint16_t> colDist(1, 20);
        std::uniform_int_distribution<int32_t> rowOffsetDist(1, 10000);
        std::uniform_int_distribution<int32_t> colOffsetDist(1, 10000);

        const uint32_t fromRowA = rowDist(rng);
        const uint16_t fromColA = colDist(rng);
        const int32_t rowOffsetA = rowOffsetDist(rng);
        const int32_t colOffsetA = colOffsetDist(rng);

        XLCellReference fromCellRefA( fromRowA, fromColA );

        XLImageAnchor anchorA;
        anchorA.initOneCell( fromCellRefA, rowOffsetA, colOffsetA );

        REQUIRE(anchorA.fromRow == fromRowA - 1);
        REQUIRE(anchorA.fromCol == fromColA - 1);
        REQUIRE(anchorA.fromRowOffset == rowOffsetA);
        REQUIRE(anchorA.fromColOffset == colOffsetA);
        REQUIRE(anchorA.getFromCellReference() == fromCellRefA);
        REQUIRE(!anchorA.isTwoCell());

        const uint32_t fromRowB = rowDist(rng);
        const uint16_t fromColB = colDist(rng);
        const uint32_t toRowB = fromRowB + rowDist(rng);
        const uint16_t toColB = fromColB + colDist(rng);

        XLCellReference fromCellRefB( fromRowB, fromColB );
        XLCellReference toCellRefB( toRowB, toColB );

        XLImageAnchor anchorB;
        anchorB.initTwoCell( fromCellRefB, toCellRefB );

        REQUIRE(anchorB.fromRow == fromRowB - 1);
        REQUIRE(anchorB.fromCol == fromColB - 1);
        REQUIRE(anchorB.fromRowOffset == 0);
        REQUIRE(anchorB.fromColOffset == 0);
        REQUIRE(anchorB.getFromCellReference() == fromCellRefB);
        REQUIRE(anchorB.toRow == toRowB - 1);
        REQUIRE(anchorB.toCol == toColB - 1);
        REQUIRE(anchorB.toRowOffset == 0);
        REQUIRE(anchorB.toColOffset == 0);
        REQUIRE(anchorB.getToCellReference() == toCellRefB);
        REQUIRE(anchorB.isTwoCell());

        const uint32_t fromRowC = rowDist(rng);
        const uint16_t fromColC = colDist(rng);
        const int32_t fromRowOffsetC = rowOffsetDist(rng);
        const int32_t fromColOffsetC = colOffsetDist(rng);
        const uint32_t toRowC = fromRowC + rowDist(rng);
        const uint16_t toColC = fromColC + colDist(rng);
        const int32_t toRowOffsetC = rowOffsetDist(rng);
        const int32_t toColOffsetC = colOffsetDist(rng);

        XLCellReference fromCellRefC( fromRowC, fromColC );
        XLCellReference toCellRefC( toRowC, toColC );

        XLImageAnchor anchorC;
        anchorC.initTwoCell( fromCellRefC, toCellRefC, 
            fromRowOffsetC, fromColOffsetC, toRowOffsetC, toColOffsetC );

        REQUIRE(anchorC.fromRow == fromRowC - 1);
        REQUIRE(anchorC.fromCol == fromColC - 1);
        REQUIRE(anchorC.fromRowOffset == fromRowOffsetC);
        REQUIRE(anchorC.fromColOffset == fromColOffsetC);
        REQUIRE(anchorC.getFromCellReference() == fromCellRefC);
        REQUIRE(anchorC.toRow == toRowC - 1);
        REQUIRE(anchorC.toCol == toColC - 1);
        REQUIRE(anchorC.toRowOffset == toRowOffsetC);
        REQUIRE(anchorC.toColOffset == toColOffsetC);
        REQUIRE(anchorC.getToCellReference() == toCellRefC);
        REQUIRE(anchorC.isTwoCell());

        // Test setFromCellReference() and getFromCellReference() standalone
        anchorC.setFromCellReference(fromCellRefB);
        REQUIRE(anchorC.fromRow == fromRowB-1);
        REQUIRE(anchorC.fromCol == fromColB-1);
        REQUIRE(anchorC.getFromCellReference() == fromCellRefB);
        anchorC.setFromCellReference(fromCellRefC);
        REQUIRE(anchorC.fromRow == fromRowC-1);
        REQUIRE(anchorC.fromCol == fromColC-1);
        REQUIRE(anchorC.getFromCellReference() == fromCellRefC);

        // Test setToCellReference() and getToCellReference() standalone
        anchorC.setToCellReference(toCellRefB);
        REQUIRE(anchorC.toRow == toRowB-1);
        REQUIRE(anchorC.toCol == toColB-1);
        REQUIRE(anchorC.getToCellReference() == toCellRefB);
        anchorC.setToCellReference(toCellRefC);
        REQUIRE(anchorC.toRow == toRowC-1);
        REQUIRE(anchorC.toCol == toColC-1);
        REQUIRE(anchorC.getToCellReference() == toCellRefC);

        // Test setDisplaySizeWithAspectRatio() with pixel dimensions
        const uint32_t testWidthPx = 100 + rowDist(rng) % 400; // 100-500 pixels
        const uint32_t testHeightPx = 100 + rowDist(rng) % 400; // 100-500 pixels
        const uint32_t maxWidthEmus = XLImageUtils::pixelsToEmus(300);
        const uint32_t maxHeightEmus = XLImageUtils::pixelsToEmus(200);
        anchorA.setDisplaySizeWithAspectRatio(testWidthPx, testHeightPx, maxWidthEmus, maxHeightEmus);
        auto displayDims = anchorA.getDisplayDimensions();
        REQUIRE(displayDims.first > 0);
        REQUIRE(displayDims.second > 0);
        REQUIRE(displayDims.first <= maxWidthEmus);
        REQUIRE(displayDims.second <= maxHeightEmus);
        // Verify aspect ratio is maintained (within tolerance)
        double expectedAspectRatio = XLImageUtils::calculateAspectRatio(testWidthPx, testHeightPx);
        double actualAspectRatio = static_cast<double>(displayDims.first) / static_cast<double>(displayDims.second);
        REQUIRE(std::abs(expectedAspectRatio - actualAspectRatio) < 0.01);

        // Test setDisplaySizeWithAspectRatio() with image data
        const uint32_t maxWidthEmusH = XLImageUtils::pixelsToEmus(500);
        const uint32_t maxHeightEmusH = XLImageUtils::pixelsToEmus(300);
        anchorA.setDisplaySizeWithAspectRatio(XLImageUtils::png31x15Data, XLMimeType::PNG, 
                                               maxWidthEmusH, maxHeightEmusH);
        auto displayDimsH = anchorA.getDisplayDimensions();
        REQUIRE(displayDimsH.first > 0);
        REQUIRE(displayDimsH.second > 0);
        REQUIRE(displayDimsH.first <= maxWidthEmusH);
        REQUIRE(displayDimsH.second <= maxHeightEmusH);
        // Verify aspect ratio matches image dimensions (31x15)
        double expectedAspectRatioH = XLImageUtils::calculateAspectRatio(31, 15);
        double actualAspectRatioH = static_cast<double>(displayDimsH.first) / static_cast<double>(displayDimsH.second);
        REQUIRE(std::abs(expectedAspectRatioH - actualAspectRatioH) < 0.01);

        // Test getDisplayDimensions()
        const uint32_t testDisplayWidth = 100000; // EMUs
        const uint32_t testDisplayHeight = 50000; // EMUs
        anchorA.displayWidthEMUs = testDisplayWidth;
        anchorA.displayHeightEMUs = testDisplayHeight;
        auto displayDimsI = anchorA.getDisplayDimensions();
        REQUIRE(displayDimsI.first == testDisplayWidth);
        REQUIRE(displayDimsI.second == testDisplayHeight);

        // Test getActualDimensions() and setActualDimensions()
        const uint32_t testActualWidth = 200000; // EMUs
        const uint32_t testActualHeight = 100000; // EMUs
        anchorA.setActualDimensions(testActualWidth, testActualHeight);
        auto actualDimsJ = anchorA.getActualDimensions();
        REQUIRE(actualDimsJ.first == testActualWidth);
        REQUIRE(actualDimsJ.second == testActualHeight);

        // Test getDisplayDimensionsPixels()
        const uint32_t testDisplayWidthK = XLImageUtils::pixelsToEmus(150);
        const uint32_t testDisplayHeightK = XLImageUtils::pixelsToEmus(75);
        anchorA.displayWidthEMUs = testDisplayWidthK;
        anchorA.displayHeightEMUs = testDisplayHeightK;
        auto displayDimsPxK = anchorA.getDisplayDimensionsPixels();
        REQUIRE(displayDimsPxK.first >= 149); // Allow for rounding
        REQUIRE(displayDimsPxK.first <= 151);
        REQUIRE(displayDimsPxK.second >= 74);
        REQUIRE(displayDimsPxK.second <= 76);

        // Test getActualDimensionsPixels()
        const uint32_t testActualWidthL = XLImageUtils::pixelsToEmus(200);
        const uint32_t testActualHeightL = XLImageUtils::pixelsToEmus(100);
        anchorA.setActualDimensions(testActualWidthL, testActualHeightL);
        auto actualDimsPxL = anchorA.getActualDimensionsPixels();
        REQUIRE(actualDimsPxL.first >= 199); // Allow for rounding
        REQUIRE(actualDimsPxL.first <= 201);
        REQUIRE(actualDimsPxL.second >= 99);
        REQUIRE(actualDimsPxL.second <= 101);

        // Test setDisplaySizeWithAspectRatio() with various aspect ratios using Catch RNG
        for (int i = 0; i < 5; ++i) {
            const uint32_t widthPx = 50 + rowDist(rng) % 450; // 50-500 pixels
            const uint32_t heightPx = 50 + rowDist(rng) % 450; // 50-500 pixels
            const uint32_t maxWidthEmusM = XLImageUtils::pixelsToEmus(300);
            const uint32_t maxHeightEmusM = XLImageUtils::pixelsToEmus(200);
            anchorA.setDisplaySizeWithAspectRatio(widthPx, heightPx, maxWidthEmusM, maxHeightEmusM);
            auto displayDimsM = anchorA.getDisplayDimensions();
            REQUIRE(displayDimsM.first > 0);
            REQUIRE(displayDimsM.second > 0);
            REQUIRE(displayDimsM.first <= maxWidthEmusM);
            REQUIRE(displayDimsM.second <= maxHeightEmusM);
        }

        // Test reset()
        anchorC.reset();
        REQUIRE(anchorC.anchorType == XLImageAnchorType::Unknown);
        REQUIRE(anchorC.fromRow == 0);
        REQUIRE(anchorC.fromCol == 0);
        REQUIRE(anchorC.displayWidthEMUs == 0);
        REQUIRE(anchorC.displayHeightEMUs == 0);
    }


    SECTION("XLWorksheet - full test") {

        XLDocument doc; doc.create("./ut_ws.xlsx");
        auto ws = doc.workbook().worksheet(1);

        REQUIRE(ws.imageCount() == 0);
        REQUIRE(!ws.hasImages());
        REQUIRE(!ws.hasImagesInXML());

        std::string errMsg;
        auto& rng = Catch::rng();
        std::uniform_int_distribution<uint32_t> rowDist(1, 20);
        std::uniform_int_distribution<uint16_t> colDist(1, 5);

        const uint32_t fromRowA = rowDist(rng);
        const uint16_t fromColA = colDist(rng);
        XLCellReference fromCellRefA( fromRowA, fromColA );

        const uint32_t fromRowB = rowDist(rng) + fromRowA;
        const uint16_t fromColB = colDist(rng);
        XLCellReference fromCellRefB( fromRowB, fromColB );
        XLImageAnchor anchorB;
        anchorB.initOneCell( fromCellRefB );

        const uint32_t fromRowC = fromRowB + rowDist(rng);
        const uint16_t fromColC = colDist(rng);
        XLCellReference fromCellRefC( fromRowC, fromColC );

        const uint32_t fromRowD = fromRowC + rowDist(rng);
        const uint16_t fromColD = colDist(rng);
        const uint32_t toRowD = fromRowD + rowDist(rng);
        const uint16_t toColD = fromColD + colDist(rng);
        XLCellReference fromCellRefD( fromRowD, fromColD );
        XLCellReference toCellRefD( toRowD, toColD );
        XLImageAnchor anchorD;
        const uint32_t maxWidthEmusD = XLImageUtils::pixelsToEmus(500);
        const uint32_t maxHeightEmusD = XLImageUtils::pixelsToEmus(300);
        anchorD.setDisplaySizeWithAspectRatio(XLImageUtils::gif26x16Data,
            XLMimeType::GIF, maxWidthEmusD, maxHeightEmusD);
        anchorD.initTwoCell( fromCellRefD, toCellRefD );

        const std::string embImgIdA = ws.embedImageFromImageData( 
            fromCellRefA, XLImageUtils::png31x15Data, XLMimeType::PNG );
        const XLEmbeddedImage embImgA = ws.getEmbImageByImageID(embImgIdA);
        const XLEmbeddedImage embImgAByRelId =
            ws.getEmbImageByRelationshipId(embImgA.relationshipId());
        REQUIRE( !embImgIdA.empty() );
        REQUIRE( embImgA.id() == embImgIdA );
        REQUIRE( embImgA == embImgAByRelId );
        REQUIRE(embImgA.anchorCell() == fromCellRefA.address());
        REQUIRE(embImgA.verifyData(&errMsg) == EXIT_SUCCESS);

        const std::string embImgIdB = ws.embedImageFromImageData( 
            anchorB, XLImageUtils::jpeg32x18Data );
        const XLEmbeddedImage embImgB = ws.getEmbImageByImageID(embImgIdB);
        const XLEmbeddedImage embImgBByRelId =
            ws.getEmbImageByRelationshipId(embImgB.relationshipId());
        REQUIRE( !embImgIdB.empty() );
        REQUIRE( embImgB.id() == embImgIdB );
        REQUIRE( embImgB == embImgBByRelId );
        REQUIRE(embImgB.anchorCell() == fromCellRefB.address());
        REQUIRE(embImgB.verifyData(&errMsg) == EXIT_SUCCESS);

        const std::string bmp33x16fileName = "bmp33x16.bmp";
        {
            std::ofstream file(bmp33x16fileName, std::ios::binary);
            file.write(reinterpret_cast<const char*>(XLImageUtils::bmp33x16Data.data()), 
                       XLImageUtils::bmp33x16Data.size());
        }
        const std::string embImgIdC = ws.embedImageFromFile( 
            fromCellRefC, bmp33x16fileName, XLMimeType::BMP );
        std::remove(bmp33x16fileName.c_str());
        const XLEmbeddedImage embImgC = ws.getEmbImageByImageID(embImgIdC);
        const XLEmbeddedImage embImgCByRelId =
            ws.getEmbImageByRelationshipId(embImgC.relationshipId());
        REQUIRE( !embImgIdC.empty() );
        REQUIRE( embImgC.id() == embImgIdC );
        REQUIRE( embImgC == embImgCByRelId );
        REQUIRE(embImgC.anchorCell() == fromCellRefC.address());
        REQUIRE(embImgC.verifyData(&errMsg) == EXIT_SUCCESS);

        const std::string gif26x16fileName = "gif26x16.gif";
        {
            std::ofstream file(gif26x16fileName, std::ios::binary);
            file.write(reinterpret_cast<const char*>(XLImageUtils::gif26x16Data.data()), 
                       XLImageUtils::gif26x16Data.size());
        }
        const std::string embImgIdD = ws.embedImageFromFile( 
            anchorD, gif26x16fileName );
        std::remove(gif26x16fileName.c_str());
        const XLEmbeddedImage embImgD = ws.getEmbImageByImageID(embImgIdD);
        const XLEmbeddedImage embImgDByRelId =
            ws.getEmbImageByRelationshipId(embImgD.relationshipId());
        REQUIRE( !embImgIdD.empty() );
        REQUIRE( embImgD.id() == embImgIdD );
        REQUIRE( embImgD == embImgDByRelId );
        REQUIRE(embImgD.anchorCell() == fromCellRefD.address());
        REQUIRE(embImgD.verifyData(&errMsg) == EXIT_SUCCESS);

        REQUIRE(ws.imageCount() == 4);
        REQUIRE(ws.hasImages());
        REQUIRE(ws.hasImagesInXML());
        REQUIRE(ws.verifyData(&errMsg) == EXIT_SUCCESS);
        REQUIRE(doc.verifyData(&errMsg) == EXIT_SUCCESS);

        std::vector<XLEmbeddedImage> embImagesA =
            ws.getEmbImagesAtCell(fromCellRefA.address());
        REQUIRE( !embImagesA.empty() );
        const std::string cellRangeAD = fromCellRefA.address() + ":" +
            toCellRefD.address();
        std::vector<XLEmbeddedImage> embImagesAD =
            ws.getEmbImagesInRange( cellRangeAD );
        REQUIRE( !embImagesAD.empty() );
        std::vector<XLEmbeddedImage> embImages = ws.getEmbImages();
        REQUIRE( !embImages.empty() );

        doc.save();
        
        ws.removeImageByImageID(embImgIdB);
        REQUIRE(ws.imageCount() == 3);
        ws.removeImageByRelationshipId(embImgD.relationshipId());
        REQUIRE(ws.imageCount() == 2);
        REQUIRE(ws.verifyData(&errMsg) == EXIT_SUCCESS);
        REQUIRE(doc.verifyData(&errMsg) == EXIT_SUCCESS);
        ws.clearImages();
        REQUIRE(ws.imageCount() == 0);
        REQUIRE(!ws.hasImages());
        REQUIRE(!ws.hasImagesInXML());
        REQUIRE(ws.verifyData(&errMsg) == EXIT_SUCCESS);
        REQUIRE(doc.verifyData(&errMsg) == EXIT_SUCCESS);

        doc.close(); 
        std::remove("./ut_ws.xlsx");
    }
}


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

// ===== External Includes ===== //
#include <algorithm>  // std::find
#include <cmath>      // std::round, std::fabs
#include <iostream>    // std::ifstream
#include <fstream>    // std::ifstream
#include <set>    // std::stringstream
#include <sstream>    // std::stringstream

// ===== OpenXLSX Includes ===== //
#include "XLImage.hpp"
#include "XLDocument.hpp"
#include "XLException.hpp"
#include "XLCellReference.hpp"
#include <memory>
#include <list>

using namespace OpenXLSX;

namespace OpenXLSX
{
    //=============================================================================
    // XLImage Implementation
    //=============================================================================

    XLImage::XLImage(XLDocument* parentDoc, size_t imageDataHash, XLMimeType mimeType, 
                     const std::string& extension, const std::string& filePackagePath, 
                     uint32_t widthPixels, uint32_t heightPixels)
        : m_parentDoc(parentDoc)
        , m_imageDataHash(imageDataHash)
        , m_mimeType(mimeType)
        , m_extension(extension)
        , m_filePackagePath(filePackagePath)
        , m_widthPixels(widthPixels)
        , m_heightPixels(heightPixels)
    {
    }

    std::vector<uint8_t> XLImage::getImageData() const
    {
        if (!m_parentDoc || m_filePackagePath.empty()) {
            return {};
        }
        
        try {
            std::string dataStr = m_parentDoc->readFile(m_filePackagePath);
            return std::vector<uint8_t>(dataStr.begin(), dataStr.end());
        } catch (...) {
            return {};
        }
    }

    int XLImage::verifyData(std::string* dbgMsg) const
    {
        int issueCount = 0;
        
        if (!m_parentDoc) {
            appendDbgMsg(dbgMsg, "XLImage: parentDoc is null");
            issueCount++;
            return issueCount;
        }
        
        if (m_filePackagePath.empty()) {
            appendDbgMsg(dbgMsg, "XLImage: filePackagePath is empty");
            issueCount++;
            return issueCount;
        }
        
        // Get image data from archive
        std::vector<uint8_t> imageData = getImageData();
        if (imageData.empty()) {
            appendDbgMsg(dbgMsg, "XLImage: image data is empty");
            issueCount++;
            return issueCount;
        }
        
        // Verify hash matches
        size_t computedHash = XLImageUtils::imageDataHash(imageData);
        if (computedHash != m_imageDataHash) {
            appendDbgMsg(dbgMsg, "XLImage: hash mismatch - stored: " + std::to_string(m_imageDataHash) + 
                         ", computed: " + std::to_string(computedHash));
            issueCount++;
        }
        
        // Compute metadata from image data
        XLMimeType computedMimeType = XLImageUtils::mimeTypeFromExtension(m_extension);
        if (computedMimeType != m_mimeType) {
            std::string storedMimeTypeStr = XLImageUtils::mimeTypeToString(m_mimeType);
            std::string computedMimeTypeStr = XLImageUtils::mimeTypeToString(computedMimeType);
            appendDbgMsg(dbgMsg, "XLImage: MIME type mismatch - stored: " + storedMimeTypeStr + 
                         ", computed: " + computedMimeTypeStr);
            issueCount++;
        }
        
        // Get dimensions from image data
        auto dimensions = XLImageUtils::getImageDimensions(imageData);
        uint32_t computedWidth = dimensions.first;
        uint32_t computedHeight = dimensions.second;
        
        if (computedWidth != m_widthPixels) {
            appendDbgMsg(dbgMsg, "XLImage: width mismatch - stored: " + std::to_string(m_widthPixels) + 
                         ", computed: " + std::to_string(computedWidth));
            issueCount++;
        }
        if (computedHeight != m_heightPixels) {
            appendDbgMsg(dbgMsg, "XLImage: height mismatch - stored: " + std::to_string(m_heightPixels) + 
                         ", computed: " + std::to_string(computedHeight));
            issueCount++;
        }
        
        return issueCount;
    }

    //=============================================================================
    // XLImageManager Implementation
    //=============================================================================

    XLImageManager::XLImageManager(XLDocument* parentDoc)
        : m_parentDoc(parentDoc)
    {
    }

    void XLImageManager::prune()
    {
        auto it = m_images.begin();
        while (it != m_images.end()) {
            if (!it->get() || it->use_count() <= 1) {
                // Remove from archive if it has a file package path
                if (it->get() && !it->get()->filePackagePath().empty()) {
                    try {
                        m_parentDoc->deleteEntry(it->get()->filePackagePath());
                    } catch (...) {
                        // Ignore deletion errors
                    }
                }
                it = m_images.erase(it);
            } else {
                ++it;
            }
        }
    }

    XLImageShPtr XLImageManager::findOrAddImage(const std::string& packagePath, 
                                                const std::vector<uint8_t>& imageData, 
                                                XLContentType contentType)
    {
        // First, try to find existing image by data hash
        XLImageShPtr existingImage = findImageByImageData(imageData);
        if (existingImage) {
            return existingImage;
        }
        
        // Image doesn't exist, create new one
        size_t imageDataHash = XLImageUtils::imageDataHash(imageData);
        XLMimeType mimeType = XLImageUtils::contentTypeToMimeType(contentType);
        std::string extension = XLImageUtils::extensionFromMimeType(mimeType);
        
        // Generate package filename
        std::string packageImageFilename;
        
        if (!packagePath.empty()) {
            // Use the provided package path (from XML when loading existing files)
            packageImageFilename = packagePath;
        } else {
            // Generate unique package filename using efficient algorithm
            packageImageFilename = generateUniquePackageFilename(extension);
        }
        
        // Get image dimensions
        auto dimensions = XLImageUtils::getImageDimensions(imageData);
        uint32_t widthPixels = dimensions.first;
        uint32_t heightPixels = dimensions.second;
        
        // Add image entry to document
        if (!m_parentDoc->addImageEntry(packageImageFilename, imageData, contentType)) {
            return nullptr;
        }
        
        // Create new XLImage object
        auto newImage = std::make_shared<XLImage>(m_parentDoc, imageDataHash, mimeType, extension,
                                                  packageImageFilename, widthPixels, heightPixels);
        
        // Add to collection
        m_images.push_back(newImage);
        
        return newImage;
    }

    XLImageShPtr XLImageManager::findImageByImageData(const std::vector<uint8_t>& imageData) const
    {
        size_t targetHash = XLImageUtils::imageDataHash(imageData);
        
        for (const auto& image : m_images) {
            if (image && image->imageDataHash() == targetHash) {
                return image;
            }
        }
        
        return nullptr;
    }


    std::string XLImageManager::generateUniquePackageFilename(const std::string& extension) const
    {
        // Collect all existing package filenames and extract base names (without extension)
        std::vector<std::string> existingBaseNames;
        for (const auto& image : m_images) {
            if (image && !image->filePackagePath().empty()) {
                std::string packagePath = image->filePackagePath();
                // Extract filename from "xl/media/image1.png" -> "image1.png"
                size_t lastSlash = packagePath.find_last_of('/');
                if (lastSlash != std::string::npos) {
                    std::string filename = packagePath.substr(lastSlash + 1);
                    // Extract base name without extension: "image1.png" -> "image1"
                    size_t lastDot = filename.find_last_of('.');
                    if (lastDot != std::string::npos) {
                        std::string baseName = filename.substr(0, lastDot);
                        existingBaseNames.push_back(baseName);
                    }
                }
            }
        }
        
        // Sort the base names for efficient binary search
        std::sort(existingBaseNames.begin(), existingBaseNames.end());
        
        // Find the next available image number
        int imageNumber = 1;
        std::string candidateBaseName;
        
        do {
            candidateBaseName = "image" + std::to_string(imageNumber);
            imageNumber++;
        } while (std::binary_search(existingBaseNames.begin(), existingBaseNames.end(), candidateBaseName));
        
        // Return the unique package filename
        return "xl/media/" + candidateBaseName + extension;
    }


    XLImageShPtr XLImageManager::findImageByFilePackagePath(const std::string& filePackagePath) const
    {
        for (const auto& image : m_images) {
            if (image && image->filePackagePath() == filePackagePath) {
                return image;
            }
        }
        
        return nullptr;
    }

    size_t XLImageManager::getImageCount() const
    {
        return m_images.size();
    }

    std::vector<XLImageShPtr>::const_iterator XLImageManager::begin() const
    {
        return m_images.begin();
    }

    std::vector<XLImageShPtr>::const_iterator XLImageManager::end() const
    {
        return m_images.end();
    }


    /**
     * @brief Remove embedded image vector for worksheet
     * @param sheetName The name of the worksheet
     */
    void XLImageManager::eraseEmbImgsForSheetName(const std::string& sheetName)
    {
        m_sheetNmEmbImgVMap.erase(sheetName);
    }

    /**
     * @brief Find or create embedded image vector for worksheet
     * @param sheetName The name of the worksheet
     * @return Shared pointer to the embedded image vector
     */
    XLEmbImgVecShPtr XLImageManager::getEmbImgsForSheetName(const std::string& sheetName)
    {
        auto itr = m_sheetNmEmbImgVMap.lower_bound(sheetName);
        if ((itr != m_sheetNmEmbImgVMap.end()) && (sheetName == itr->first)) {
            return itr->second;
        } else {
            auto result = std::make_shared<XLEmbImgVec>();
            m_sheetNmEmbImgVMap.insert(itr, std::make_pair(sheetName, result));
            return result;
        }
    }

    /**
     * @brief Find embedded image vector for worksheet, return null pointer if not found
     * @param sheetName The name of the worksheet
     * @return Shared pointer to the embedded image vector, or nullptr if not found
     */
    XLEmbImgVecShPtr XLImageManager::getEmbImgsForSheetName(const std::string& sheetName) const
    {
        auto itr = m_sheetNmEmbImgVMap.find(sheetName);
        if (itr != m_sheetNmEmbImgVMap.end()) {
            return itr->second;
        }
        return nullptr;
    }

    int XLImageManager::verifyData(std::string* dbgMsg) const
    {
        int totalIssues = 0;
        
        // Check each image has same parentDoc
        for (const auto& image : m_images) {
            if (image && image->parentDoc() != m_parentDoc) {
                appendDbgMsg(dbgMsg, "XLImageManager: image has different parentDoc");
                totalIssues++;
            }
        }
        
        
        // Check each image has unique imageData hash
        // TODO: allow for hash collision -- compare binary image data when hash values are same.
        std::set<size_t> imageDataHashes;
        for (const auto& image : m_images) {
            if (image) {
                if (imageDataHashes.find(image->imageDataHash()) != imageDataHashes.end()) {
                    appendDbgMsg(dbgMsg, "XLImageManager: duplicate imageDataHash: " + std::to_string(image->imageDataHash()));
                    totalIssues++;
                } else {
                    imageDataHashes.insert(image->imageDataHash());
                }
            }
        }
        
        // Call verifyData() for each non-null image
        for (const auto& image : m_images) {
            if (image) {
                totalIssues += image->verifyData(dbgMsg);
            }
        }

        // Check package path uniqueness and format
        std::set<std::string> packagePaths;
        for (const auto& image : m_images) {
            if (image && !image->filePackagePath().empty()) {
                std::string packagePath = image->filePackagePath();
                
                // Check format
                if (packagePath.find("xl/media/") != 0) {
                    appendDbgMsg(dbgMsg, "XLImageManager: package path '" + packagePath + "' does not start with 'xl/media/'");
                    totalIssues++;
                }
                
                // Check uniqueness
                if (packagePaths.find(packagePath) != packagePaths.end()) {
                    appendDbgMsg(dbgMsg, "XLImageManager: duplicate package path: " + packagePath);
                    totalIssues++;
                } else {
                    packagePaths.insert(packagePath);
                }
            }
        }

        // Check that all images have valid parent document
        for (const auto& image : m_images) {
            if (image && !image->parentDoc()) {
                appendDbgMsg(dbgMsg, "XLImageManager: image has null parentDoc");
                totalIssues++;
            }
        }

        // Check that all images have valid image data
        for (const auto& image : m_images) {
            if (image && image->getImageData().empty()) {
                appendDbgMsg(dbgMsg, "XLImageManager: image has empty image data");
                totalIssues++;
            }
        }
        
        // Check consistency between m_images vector and archive
        if (m_parentDoc) {
            // Check that all images in m_images actually exist in the archive
            for (const auto& image : m_images) {
                if (image && !image->filePackagePath().empty()) {
                    std::string archiveContent = m_parentDoc->readFile(image->filePackagePath());
                    if (archiveContent.empty()) {
                        appendDbgMsg(dbgMsg, "XLImageManager: image '" + image->filePackagePath() + "' not found in archive");
                        totalIssues++;
                    }
                }
            }
            
            // Check for orphaned files in archive (files that exist in archive but not in m_images)
            // This is more complex and would require iterating through all archive entries
            // For now, we'll skip this check as it's expensive and the above check is more critical
        }

        std::set<XLEmbImgVecShPtr> uniqueEmbImgPtrs;
        // Check that each ptr in m_sheetNmEmbImgVMap is unique
        std::set<XLEmbImgVecShPtr> uniquePtrs;
        for (const auto& pair : m_sheetNmEmbImgVMap) {
            if (pair.second) {
                if (uniquePtrs.find(pair.second) != uniquePtrs.end()) {
                    appendDbgMsg(dbgMsg, "XLImageManager: duplicate XLEmbImgVecShPtr for worksheet: " + pair.first);
                    totalIssues++;
                } else {
                    uniquePtrs.insert(pair.second);
                }
            }
        }
        
        // Check that each ptr in m_sheetNmEmbImgVMap is non-NULL
        for (const auto& pair : m_sheetNmEmbImgVMap) {
            if (!pair.second) {
                appendDbgMsg(dbgMsg, "XLImageManager: NULL XLEmbImgVecShPtr for worksheet: " + pair.first);
                totalIssues++;
            }
        }
        
        // For each ptr in m_sheetNmEmbImgVMap, call ptr->verifyData()
        for (const auto& pair : m_sheetNmEmbImgVMap) {
            if (pair.second) {
                for (const auto& embImage : *pair.second) {
                    std::string embImageDbgMsg;
                    int embImageIssues = embImage.verifyData(&embImageDbgMsg);
                    if (embImageIssues > 0) {
                        appendDbgMsg(dbgMsg, "XLImageManager: worksheet '" + pair.first + "' has embedded image issues: " + embImageDbgMsg);
                        totalIssues += embImageIssues;
                    }
                }
            }
        }
        
        return totalIssues;
    }

    void XLImageManager::printReferenceCounts(const std::string& title) const
    {
        std::cout << "\n=== " << title << " ===" << std::endl;
        std::cout << "Total XLImageShPtr objects: " << m_images.size() << std::endl;
        
        for (size_t i = 0; i < m_images.size(); ++i) {
            const auto& image = m_images[i];
            if (image) {
                std::cout << "  [" << i << "] refs=" << image.use_count() 
                          << " package=" << image->filePackagePath()
                          << " mime=" << XLImageUtils::mimeTypeToString(image->mimeType())
                          << " size=" << image->widthPixels() << "x" << image->heightPixels()
                          << " hash=" << image->imageDataHash() << std::endl;
            } else {
                std::cout << "  [" << i << "] NULL pointer" << std::endl;
            }
        }
        std::cout << "=== End " << title << " ===" << std::endl;
    }


    // ========== XLEmbeddedImage Member Functions ========== //

    /**
     * @details Set the image ID
     */
    void XLEmbeddedImage::setId(const std::string& imageId)
    {
        m_imageAnchor.imageId = imageId;
    }

    /**
     * @details Get the image data
     */
    std::vector<uint8_t> XLEmbeddedImage::data() const
    {
        static const std::vector<uint8_t> empty;
        if (!m_image) return empty;
        return m_image->getImageData();
    }

    /**
     * @details Get the MIME type
     */
    /**
     * @details Get the file extension
     */
    const std::string& XLEmbeddedImage::extension() const
    {
        static std::string empty;
        if (!m_image) return empty;
        return m_image->extension();
    }

    /**
     * @details Get the unique image ID
     */
    const std::string& XLEmbeddedImage::id() const
    {
        return m_imageAnchor.imageId;
    }

    /**
     * @details Get the image width in pixels
     */
    uint32_t XLEmbeddedImage::widthPixels() const
    {
        if (!m_image) return 0;
        return m_image->widthPixels();
    }

    /**
     * @details Get the image height in pixels
     */
    uint32_t XLEmbeddedImage::heightPixels() const
    {
        if (!m_image) return 0;
        return m_image->heightPixels();
    }

    /**
     * @details Get the image MIME type
     */
    XLMimeType XLEmbeddedImage::mimeType() const
    {
        if (!m_image) return XLMimeType::Unknown;
        return m_image->mimeType();
    }

    /**
     * @details Set the display width in Excel units (EMUs)
     */
    void XLEmbeddedImage::setDisplayWidthEMUs(uint32_t width)
    {
        m_imageAnchor.displayWidthEMUs = width;
    }

    /**
     * @details Set the display height in Excel units (EMUs)
     */
    void XLEmbeddedImage::setDisplayHeightEMUs(uint32_t height)
    {
        m_imageAnchor.displayHeightEMUs = height;
    }

    /**
     * @details Get the display width in Excel units (EMUs)
     */
    uint32_t XLEmbeddedImage::displayWidthEMUs() const
    {
        return m_imageAnchor.displayWidthEMUs;
    }

    /**
     * @details Get the display height in Excel units (EMUs)
     */
    uint32_t XLEmbeddedImage::displayHeightEMUs() const
    {
        return m_imageAnchor.displayHeightEMUs;
    }

    /**
     * @details Check if the image is valid
     */
    bool XLEmbeddedImage::isValid() const
    {
        return m_image && !m_imageAnchor.imageId.empty();
    }

    /**
     * @details Get the content type for this image
     */
    XLContentType XLEmbeddedImage::contentType() const
    {
        if (!m_image) return XLContentType::Unknown;
        return XLImageUtils::mimeTypeToContentType(m_image->mimeType());
    }

    // ========== XLEmbeddedImage Private Member Functions ========== //

    /**
     * @details Generate a unique image ID
     */
    std::string XLEmbeddedImage::generateId(uint32_t imageNumber)
    {
        return "img" + std::to_string(imageNumber);
    }

    /**
     * @details Set display size in pixels (converts to EMUs automatically)
     */
    void XLEmbeddedImage::setDisplaySizePixels(uint32_t widthPixels, uint32_t heightPixels)
    {
        m_imageAnchor.displayWidthEMUs = XLImageUtils::pixelsToEmus(widthPixels);
        m_imageAnchor.displayHeightEMUs = XLImageUtils::pixelsToEmus(heightPixels);
    }

    /**
     * @details Compare two images for debugging
     */
    int XLEmbeddedImage::compare(const XLEmbeddedImage& other, std::string* diffMsg) const
    {

        // Compare image data (most important for embedded images)
        if (m_image != other.m_image) {
            auto thisData = m_image ? m_image->getImageData() : std::vector<uint8_t>();
            auto otherData = other.m_image ? other.m_image->getImageData() : std::vector<uint8_t>();
            if (thisData != otherData) {
                appendDbgMsg(diffMsg, "image data differs (size: " + std::to_string(thisData.size()) + 
                          " vs " + std::to_string(otherData.size()) + " bytes)");
                return thisData.size() < otherData.size() ? -1 : 1;
            }
        }

        // Compare MIME type
        const XLMimeType thisMimeType = mimeType();
        const XLMimeType otherMimeType = other.mimeType();
        int mimeCompare = static_cast<int>(thisMimeType) - static_cast<int>(otherMimeType);
        if (mimeCompare != 0) {
            std::string thisMimeTypeStr = m_image ? XLImageUtils::mimeTypeToString(m_image->mimeType()) : "";
            std::string otherMimeTypeStr = other.m_image ? XLImageUtils::mimeTypeToString(other.m_image->mimeType()) : "";
            appendDbgMsg(diffMsg, "MIME type differs: '" + thisMimeTypeStr + "' vs '" + otherMimeTypeStr + "'");
            return mimeCompare;
        }

        // Compare extension
        std::string thisExtension = m_image ? m_image->extension() : "";
        std::string otherExtension = other.m_image ? other.m_image->extension() : "";
        int extCompare = thisExtension.compare(otherExtension);
        if (extCompare != 0) {
            appendDbgMsg(diffMsg, "extension differs: '" + thisExtension + "' vs '" + otherExtension + "'");
            return extCompare;
        }

        // Compare ID
        int idCompare = m_imageAnchor.imageId.compare(other.m_imageAnchor.imageId);
        if (idCompare != 0) {
            appendDbgMsg(diffMsg, "image ID differs: '" + m_imageAnchor.imageId + "' vs '" + other.m_imageAnchor.imageId + "'");
            return idCompare;
        }

        // Compare dimensions
        uint32_t thisWidth = m_image ? m_image->widthPixels() : 0;
        uint32_t otherWidth = other.m_image ? other.m_image->widthPixels() : 0;
        if (thisWidth != otherWidth) {
            appendDbgMsg(diffMsg, "width differs: " + std::to_string(thisWidth) + 
                      " vs " + std::to_string(otherWidth) + " pixels");
            return thisWidth < otherWidth ? -1 : 1;
        }

        uint32_t thisHeight = m_image ? m_image->heightPixels() : 0;
        uint32_t otherHeight = other.m_image ? other.m_image->heightPixels() : 0;
        if (thisHeight != otherHeight) {
            appendDbgMsg(diffMsg, "height differs: " + std::to_string(thisHeight) + 
                      " vs " + std::to_string(otherHeight) + " pixels");
            return thisHeight < otherHeight ? -1 : 1;
        }

        // Compare display dimensions
        if (m_imageAnchor.displayWidthEMUs != other.m_imageAnchor.displayWidthEMUs) {
            appendDbgMsg(diffMsg, "display width differs: " + std::to_string(m_imageAnchor.displayWidthEMUs) + 
                      " vs " + std::to_string(other.m_imageAnchor.displayWidthEMUs) + " EMUs");
            return m_imageAnchor.displayWidthEMUs < other.m_imageAnchor.displayWidthEMUs ? -1 : 1;
        }

        if (m_imageAnchor.displayHeightEMUs != other.m_imageAnchor.displayHeightEMUs) {
            appendDbgMsg(diffMsg, "display height differs: " + std::to_string(m_imageAnchor.displayHeightEMUs) + 
                      " vs " + std::to_string(other.m_imageAnchor.displayHeightEMUs) + " EMUs");
            return m_imageAnchor.displayHeightEMUs < other.m_imageAnchor.displayHeightEMUs ? -1 : 1;
        }

        // Compare registry/relationship fields
        int relIdCompare = m_imageAnchor.relationshipId.compare(other.m_imageAnchor.relationshipId);
        if (relIdCompare != 0) {
            appendDbgMsg(diffMsg, "relationship ID differs: '" + m_imageAnchor.relationshipId + "' vs '" + other.m_imageAnchor.relationshipId + "'");
            return relIdCompare;
        }

        int cellCompare = anchorCell().compare(other.anchorCell());
        if (cellCompare != 0) {
            appendDbgMsg(diffMsg, "anchor cell differs: '" + anchorCell() + "' vs '" + other.anchorCell() + "'");
            return cellCompare;
        }

        int typeCompare = static_cast<int>(m_imageAnchor.anchorType) - static_cast<int>(other.m_imageAnchor.anchorType);
        if (typeCompare != 0) {
            std::string thisAnchorTypeStr = XLImageAnchorUtils::anchorTypeToString(m_imageAnchor.anchorType);
            std::string otherAnchorTypeStr = XLImageAnchorUtils::anchorTypeToString(other.m_imageAnchor.anchorType);
            appendDbgMsg(diffMsg, "anchor type differs: '" + thisAnchorTypeStr + "' vs '" + otherAnchorTypeStr + "'");
            return typeCompare;
        }

        // Images are identical
        return 0;
    }

    /**
     * @details Verify internal data integrity and class invariants
     */
    int XLEmbeddedImage::verifyData(std::string* dbgMsg) const
    {
        int errorCount = 0;

        // Check basic data integrity
        if (!m_image || m_image->getImageData().empty()) {
            appendDbgMsg(dbgMsg, "image data is empty");
            errorCount++;
        }

        // MIME type
        const XLMimeType thisMimeType = mimeType();
        const std::string mimeTypeStr = XLImageUtils::mimeTypeToString(thisMimeType);

        // Check extension consistency
        std::string extension = m_image ? m_image->extension() : "";
        if (extension.empty()) {
            appendDbgMsg(dbgMsg, "file extension is empty");
            errorCount++;
        }

        // Check ID consistency
        if (m_imageAnchor.imageId.empty()) {
            appendDbgMsg(dbgMsg, "image ID is empty");
            errorCount++;
        }

        // Check relationship ID consistency
        if (m_imageAnchor.relationshipId.empty()) {
            appendDbgMsg(dbgMsg, "relationship ID is empty");
            errorCount++;
        }

        // Check anchor type consistency
        if (m_imageAnchor.anchorType == XLImageAnchorType::Unknown) {
            appendDbgMsg(dbgMsg, "anchor type is unknown");
            errorCount++;
        }

        // Check dimension consistency
        uint32_t widthPixels = m_image ? m_image->widthPixels() : 0;
        uint32_t heightPixels = m_image ? m_image->heightPixels() : 0;
        if (widthPixels == 0) {
            appendDbgMsg(dbgMsg, "width in pixels is zero");
            errorCount++;
        }

        if (heightPixels == 0) {
            appendDbgMsg(dbgMsg, "height in pixels is zero");
            errorCount++;
        }

        // Check display dimension consistency
        if (m_imageAnchor.displayWidthEMUs == 0) {
            appendDbgMsg(dbgMsg, "display width in EMUs is zero");
            errorCount++;
        }

        if (m_imageAnchor.displayHeightEMUs == 0) {
            appendDbgMsg(dbgMsg, "display height in EMUs is zero");
            errorCount++;
        }

        // Check MIME type and extension consistency
        if ( (thisMimeType != XLMimeType::Unknown) && !extension.empty()) {
            if (( XLMimeType::PNG ==  thisMimeType && extension != ".png") ||
                ( XLMimeType::JPEG ==  thisMimeType && extension != ".jpg" && extension != ".jpeg") ||
                ( XLMimeType::GIF == thisMimeType && extension != ".gif") ||
                ( XLMimeType::BMP == thisMimeType && extension != ".bmp")) {
                appendDbgMsg(dbgMsg, "MIME type '" + mimeTypeStr + 
                    "' does not match extension '" + extension + "'");
                errorCount++;
            }
        }

        // Check package path consistency
        if (m_image && !m_image->filePackagePath().empty()) {
            std::string packagePath = m_image->filePackagePath();
            if (packagePath.find("xl/media/") != 0) {
                appendDbgMsg(dbgMsg, "package path '" + packagePath + "' does not start with 'xl/media/'");
                errorCount++;
            }
        }

        return errorCount;
    }

    // Registry/relationship setter functions (for XLImageInfo compatibility)
    /**
     * @details Set the relationship ID
     */
    void XLEmbeddedImage::setRelationshipId(const std::string& relationshipId)
    {
        m_imageAnchor.relationshipId = relationshipId;
    }

    /**
     * @details Set the anchor cell reference
     */
    void XLEmbeddedImage::setAnchorCell(const std::string& anchorCell)
    {
        auto [row, col] = cellRefToRowCol(anchorCell);
        m_imageAnchor.fromRow = row;
        m_imageAnchor.fromCol = col;
    }

    /**
     * @details Set the anchor type
     */
    void XLEmbeddedImage::setAnchorType(XLImageAnchorType anchorType)
    {
        m_imageAnchor.anchorType = anchorType;
    }

    // Registry/relationship getter functions (for XLImageInfo compatibility)
    /**
     * @details Get the relationship ID
     */
    const std::string& XLEmbeddedImage::relationshipId() const
    {
        return m_imageAnchor.relationshipId;
    }

    /**
     * @details Get the anchor cell reference
     */
    std::string XLEmbeddedImage::anchorCell() const
    {
        return rowColToCellRef(m_imageAnchor.fromRow, m_imageAnchor.fromCol);
    }

    /**
     * @details Get the anchor type
     */
    const XLImageAnchorType& XLEmbeddedImage::anchorType() const
    {
        return m_imageAnchor.anchorType;
    }

    /**
     * @details Check if this image is empty (default constructed)
     */
    bool XLEmbeddedImage::isEmpty() const
    {
        return m_imageAnchor.imageId.empty();
    }

    /**
     * @details Get the shared pointer to the image data
     */
    XLImageShPtr XLEmbeddedImage::getImage() const
    {
        return m_image;
    }

    /**
     * @details Set the shared pointer to the image data
     */
    void XLEmbeddedImage::setImage(XLImageShPtr image)
    {
        m_image = image;
    }

    /* TODO: add unit test */
    std::string XLEmbeddedImage::rowColToCellRef(uint32_t row, uint16_t col)
    {
        std::string result;
        
        // Convert column to letters (A, B, C, ..., Z, AA, AB, ...)
        ++col;
        while (col > 0) {
            col--; // Convert to 0-based first
            result = static_cast<char>('A' + (col % 26)) + result;
            col = col / 26;
        }
        
        // Add row number (1-based)
        result += std::to_string(row + 1);
        
        return result;
    }

    /* TODO: add unit test */
    std::pair<uint32_t, uint16_t> XLEmbeddedImage::cellRefToRowCol(const std::string& cellRef)
    {
        uint32_t row = 0;
        uint16_t col = 0;
        
        // Parse the cell reference (e.g., "A1", "B5", "AA10")
        size_t i = 0;
        while (i < cellRef.length() && std::isalpha(cellRef[i])) {
            col = col * 26 + (cellRef[i] - 'A' + 1);
            i++;
        }
        col--; // Convert to 0-based
        
        if (i < cellRef.length()) {
            row = std::stoul(cellRef.substr(i)) - 1; // Convert to 0-based
        }
        
        return {row, col};
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
    bool XLImageUtils::matchesElementName(const std::string& elementName, const std::string& targetName)
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

}    // namespace OpenXLSX

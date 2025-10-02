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
#include <fstream>    // std::ifstream
#include <sstream>    // std::stringstream
#include <algorithm>  // std::find
#include <cmath>      // std::round, std::fabs

// ===== OpenXLSX Includes ===== //
#include "XLImage.hpp"
#include "XLException.hpp"

using namespace OpenXLSX;

namespace OpenXLSX
{
    // ========== XLImage Member Functions ========== //

    /**
     * @details Constructor from file path
     */
    XLImage::XLImage(const std::string& imagePath)
    {
        loadFromFile(imagePath);
    }

    /**
     * @details Constructor from binary data
     */
    XLImage::XLImage(const std::vector<uint8_t>& imageData, const std::string& mimeType)
    {
        loadFromData(imageData, mimeType);
    }

    /**
     * @details Load image from file (backward compatibility)
     */
    bool XLImage::loadFromFile(const std::string& imagePath)
    {
        // Generate a temporary ID for backward compatibility
        std::string tempId = "temp_id";
        return loadFromFile(imagePath, tempId);
    }

    /**
     * @details Load image from file
     */
    bool XLImage::loadFromFile(const std::string& imagePath, const std::string& imageId)
    {
        std::ifstream file(imagePath, std::ios::binary);
        if (!file.is_open()) {
            return false;
        }

        // Read the entire file into a vector
        file.seekg(0, std::ios::end);
        size_t fileSize = file.tellg();
        file.seekg(0, std::ios::beg);

        m_imageData.resize(fileSize);
        file.read(reinterpret_cast<char*>(m_imageData.data()), fileSize);
        file.close();

        // Determine file extension and MIME type
        size_t dotPos = imagePath.find_last_of('.');
        if (dotPos != std::string::npos) {
            m_extension = imagePath.substr(dotPos);
            std::transform(m_extension.begin(), m_extension.end(), m_extension.begin(), ::tolower);
        }

        m_mimeType = mimeTypeFromExtension(m_extension);

        // Get image dimensions
        auto dimensions = getImageDimensions(m_imageData, m_mimeType);
        m_widthPixels = dimensions.first;
        m_heightPixels = dimensions.second;

        // Set default display size (same as original size)
        m_displayWidth = pixelsToEMUs(m_widthPixels);
        m_displayHeight = pixelsToEMUs(m_heightPixels);

        // Set the provided ID
        m_id = imageId;

        return true;
    }

    /**
     * @details Load image from binary data (backward compatibility)
     */
    bool XLImage::loadFromData(const std::vector<uint8_t>& imageData, const std::string& mimeType)
    {
        // Generate a temporary ID for backward compatibility
        std::string tempId = "temp_id";
        return loadFromData(imageData, mimeType, tempId);
    }

    /**
     * @details Load image from binary data
     */
    bool XLImage::loadFromData(const std::vector<uint8_t>& imageData, const std::string& mimeType, const std::string& imageId)
    {
        m_imageData = imageData;
        m_mimeType = mimeType;
        m_extension = extensionFromMimeType(mimeType);

        // Get image dimensions
        auto dimensions = getImageDimensions(m_imageData, m_mimeType);
        m_widthPixels = dimensions.first;
        m_heightPixels = dimensions.second;

        // Set default display size (same as original size)
        m_displayWidth = pixelsToEMUs(m_widthPixels);
        m_displayHeight = pixelsToEMUs(m_heightPixels);

        // Set the provided ID
        m_id = imageId;

        return true;
    }

    /**
     * @details Set the image ID
     */
    void XLImage::setId(const std::string& imageId)
    {
        m_id = imageId;
    }

    /**
     * @details Get the image data
     */
    const std::vector<uint8_t>& XLImage::data() const
    {
        return m_imageData;
    }

    /**
     * @details Get the MIME type
     */
    const std::string& XLImage::mimeType() const
    {
        return m_mimeType;
    }

    /**
     * @details Get the file extension
     */
    const std::string& XLImage::extension() const
    {
        return m_extension;
    }

    /**
     * @details Get the unique image ID
     */
    const std::string& XLImage::id() const
    {
        return m_id;
    }

    /**
     * @details Get the image width in pixels
     */
    uint32_t XLImage::widthPixels() const
    {
        return m_widthPixels;
    }

    /**
     * @details Get the image height in pixels
     */
    uint32_t XLImage::heightPixels() const
    {
        return m_heightPixels;
    }

    /**
     * @details Set the display width in Excel units (EMUs)
     */
    void XLImage::setDisplayWidth(uint32_t width)
    {
        m_displayWidth = width;
    }

    /**
     * @details Set the display height in Excel units (EMUs)
     */
    void XLImage::setDisplayHeight(uint32_t height)
    {
        m_displayHeight = height;
    }

    /**
     * @details Get the display width in Excel units (EMUs)
     */
    uint32_t XLImage::displayWidth() const
    {
        return m_displayWidth;
    }

    /**
     * @details Get the display height in Excel units (EMUs)
     */
    uint32_t XLImage::displayHeight() const
    {
        return m_displayHeight;
    }

    /**
     * @details Check if the image is valid
     */
    bool XLImage::isValid() const
    {
        return !m_imageData.empty() && !m_mimeType.empty() && !m_id.empty();
    }

    /**
     * @details Get the content type for this image
     */
    XLContentType XLImage::contentType() const
    {
        if (m_mimeType == ImageMimeTypes::PNG) return XLContentType::ImagePNG;
        if (m_mimeType == ImageMimeTypes::JPEG) return XLContentType::ImageJPEG;
        if (m_mimeType == ImageMimeTypes::BMP) return XLContentType::ImageBMP;
        if (m_mimeType == ImageMimeTypes::GIF) return XLContentType::ImageGIF;
        return XLContentType::Unknown;
    }

    // ========== XLImage Private Member Functions ========== //

    /**
     * @details Generate a unique image ID
     */
    std::string XLImage::generateId(uint32_t imageNumber)
    {
        return "img" + std::to_string(imageNumber);
    }

    /**
     * @details Determine MIME type from file extension
     */
    std::string XLImage::mimeTypeFromExtension(const std::string& extension)
    {
        if (extension == ".png") return ImageMimeTypes::PNG;
        if (extension == ".jpg" || extension == ".jpeg") return ImageMimeTypes::JPEG;
        if (extension == ".bmp") return ImageMimeTypes::BMP;
        if (extension == ".gif") return ImageMimeTypes::GIF;
        return "";
    }

    /**
     * @details Determine file extension from MIME type
     */
    std::string XLImage::extensionFromMimeType(const std::string& mimeType)
    {
        if (mimeType == ImageMimeTypes::PNG) return ".png";
        if (mimeType == ImageMimeTypes::JPEG) return ".jpg";
        if (mimeType == ImageMimeTypes::BMP) return ".bmp";
        if (mimeType == ImageMimeTypes::GIF) return ".gif";
        return "";
    }

    /**
     * @details Convert pixels to EMUs (Excel Measurement Units)
     * 1 pixel = 9525 EMUs (approximately)
     */
    uint32_t XLImage::pixelsToEMUs(uint32_t pixels)
    {
        return pixels * EMU_TO_PIXEL_RATIO;
    }

    /**
     * @details Get image dimensions from binary data
     * This is a simplified implementation that works for basic cases
     * For production use, consider using a proper image library like libpng, libjpeg, etc.
     */
    std::pair<uint32_t, uint32_t> XLImage::getImageDimensions(const std::vector<uint8_t>& data, const std::string& mimeType)
    {
        if (data.size() < 8) return {0, 0};

        // PNG signature check and dimension extraction
        if (mimeType == ImageMimeTypes::PNG) {
            if (data.size() >= 24 && data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47) {
                // PNG IHDR chunk contains width and height at bytes 16-23
                uint32_t width = (data[16] << 24) | (data[17] << 16) | (data[18] << 8) | data[19];
                uint32_t height = (data[20] << 24) | (data[21] << 16) | (data[22] << 8) | data[23];
                return {width, height};
            }
        }
        // JPEG signature check and dimension extraction
        else if (mimeType == ImageMimeTypes::JPEG) {
            if (data.size() >= 2 && data[0] == 0xFF && data[1] == 0xD8) {
                // For JPEG, we need to find the SOF0 marker (0xFFC0) and extract dimensions
                for (size_t i = 2; i < data.size() - 8; ++i) {
                    if (data[i] == 0xFF && data[i + 1] == 0xC0) {
                        uint32_t height = (data[i + 5] << 8) | data[i + 6];
                        uint32_t width = (data[i + 7] << 8) | data[i + 8];
                        return {width, height};
                    }
                }
            }
        }
        // BMP signature check and dimension extraction
        else if (mimeType == ImageMimeTypes::BMP) {
            if (data.size() >= 26 && data[0] == 0x42 && data[1] == 0x4D) {
                // BMP header contains width and height at bytes 18-25
                uint32_t width = data[18] | (data[19] << 8) | (data[20] << 16) | (data[21] << 24);
                uint32_t height = data[22] | (data[23] << 8) | (data[24] << 16) | (data[25] << 24);
                return {width, height};
            }
        }
        // GIF signature check and dimension extraction
        else if (mimeType == ImageMimeTypes::GIF) {
            if (data.size() >= 10 && data[0] == 0x47 && data[1] == 0x49 && data[2] == 0x46) {
                // GIF header contains width and height at bytes 6-9
                uint32_t width = data[6] | (data[7] << 8);
                uint32_t height = data[8] | (data[9] << 8);
                return {width, height};
            }
        }

        // Default fallback - assume square image
        return {100, 100};
    }

    /**
     * @details Set display size in pixels (converts to EMUs automatically)
     */
    void XLImage::setDisplaySizePixels(uint32_t widthPixels, uint32_t heightPixels)
    {
        m_displayWidth = pixelsToEmus(widthPixels);
        m_displayHeight = pixelsToEmus(heightPixels);
    }

    /**
     * @details Set display size maintaining aspect ratio
     */
    void XLImage::setDisplaySizeWithAspectRatio(uint32_t maxWidthEmus, uint32_t maxHeightEmus)
    {
        uint32_t targetWidth = pixelsToEmus(m_widthPixels);
        uint32_t targetHeight = pixelsToEmus(m_heightPixels);

        if (m_widthPixels > 0 || m_heightPixels > 0) {
            double aspectRatio = static_cast<double>(m_widthPixels) / static_cast<double>(m_heightPixels);
            assert(aspectRatio > 0.0);

            const double maxWidth = static_cast<double>(maxWidthEmus);
            const double maxHeight = static_cast<double>(maxHeightEmus);
            const double widthA = maxHeight * aspectRatio;

            if (widthA <= maxWidth) {
                /*  
                +------------+--------+
                |            |        |
                |            |        |
                |            |        |
                |            |        |maxHeight
                |            |        |
                |            |        |
                |            |        |
                |   widthA   |        |
                +------------+--------+
                       maxWidth 
                */
                targetWidth = static_cast<uint32_t>(round(widthA));
                targetHeight = maxHeightEmus;
            }
            else {
                /* 
                +---------------------+
                |                     |
                |                     |
                +---------------------+
                |                     |maxHeight
                |                     |
                |              heightB|
                |                     |
                |                     |
                +---------------------+
                       maxWidth  
                */
                const double heightB = maxWidth / aspectRatio;
                assert(heightB <= (1.00000001 * maxHeight));
                targetWidth = maxWidthEmus;
                targetHeight = static_cast<uint32_t>(round(heightB));
            }

            assert(fabs(aspectRatio - (static_cast<double>(targetWidth) / 
                    static_cast<double>(targetHeight))) < 1e-4);
        }

        m_displayWidth = targetWidth;
        m_displayHeight = targetHeight;
    }


    /**
     * @details Set display size maintaining aspect ratio
     */
    void XLImage::setDisplaySizePixelsWithAspectRatio(uint32_t maxWidthPixels, uint32_t maxHeightPixels)
    {
        const uint32_t maxWidthEmus = pixelsToEmus(maxWidthPixels);
        const uint32_t maxHeightEmus = pixelsToEmus(maxHeightPixels);
        setDisplaySizeWithAspectRatio(maxWidthEmus, maxHeightEmus);
    }

    /**
     * @details Convert pixels to EMUs (Excel units)
     * 1 pixel = 9525 EMUs (approximately)
     */
    uint32_t XLImage::pixelsToEmus(uint32_t pixels)
    {
        return pixels * EMU_TO_PIXEL_RATIO;
    }

    /**
     * @details Convert EMUs to pixels
     */
    uint32_t XLImage::emusToPixels(uint32_t emus)
    {
        return emus / EMU_TO_PIXEL_RATIO;
    }

}    // namespace OpenXLSX

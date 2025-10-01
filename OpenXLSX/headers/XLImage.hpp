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

namespace OpenXLSX
{
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
        uint32_t width() const;

        /**
         * @brief Get the image height in pixels
         * @return The height in pixels
         */
        uint32_t height() const;

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
         * @param maxWidthPixels Maximum width in pixels (0 = no limit)
         * @param maxHeightPixels Maximum height in pixels (0 = no limit)
         */
        void setDisplaySizeWithAspectRatio(uint32_t maxWidthPixels = 0, uint32_t maxHeightPixels = 0);

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

    private:
        std::vector<uint8_t> m_imageData;     /**< Binary image data */
        std::string m_mimeType;               /**< MIME type of the image */
        std::string m_extension;              /**< File extension */
        std::string m_id;                     /**< Unique image ID */
        uint32_t m_width{0};                  /**< Image width in pixels */
        uint32_t m_height{0};                 /**< Image height in pixels */
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

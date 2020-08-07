/*

   ____                               ____      ___ ____       ____  ____      ___
  6MMMMb                              `MM(      )M' `MM'      6MMMMb\`MM(      )M'
 8P    Y8                              `MM.     d'   MM      6M'    ` `MM.     d'
6M      Mb __ ____     ____  ___  __    `MM.   d'    MM      MM        `MM.   d'
MM      MM `M6MMMMb   6MMMMb `MM 6MMb    `MM. d'     MM      YM.        `MM. d'
MM      MM  MM'  `Mb 6M'  `Mb MMM9 `Mb    `MMd       MM       YMMMMb     `MMd
MM      MM  MM    MM MM    MM MM'   MM     dMM.      MM           `Mb     dMM.
MM      MM  MM    MM MMMMMMMM MM    MM    d'`MM.     MM            MM    d'`MM.
YM      M9  MM    MM MM       MM    MM   d'  `MM.    MM            MM   d'  `MM.
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
  DISCLAIMED. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY
  DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
  (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
  LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
  ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
  (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
  SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

 */

#ifndef OPENXLSX_XLXMLFILE_HPP
#define OPENXLSX_XLXMLFILE_HPP

#pragma warning(push)
#pragma warning(disable : 4251)
#pragma warning(disable : 4275)

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    class XLXmlData;
    class XLDocument;

    /**
     * @brief The XLXmlFile class provides an interface for derived classes to use.
     * It functions as an ancestor to all classes which are represented by an .xml file in an .xlsx package.
     * @warning The XLXmlFile class is not intended to be instantiated on it's own, but to provide an interface for
     * derived classes. Also, it should not be used polymorphically. For that reason, the destructor is not declared virtual.
     */
    class OPENXLSX_EXPORT XLXmlFile
    {
    public:    // ===== PUBLIC MEMBER FUNCTIONS
        /**
         * @brief Default constructor.
         */
        XLXmlFile() = default;

        /**
         * @brief Constructor. Creates an object based on the xmlData input.
         * @param xmlData An XLXmlData object with the XML data to be represented by the object.
         */
        explicit XLXmlFile(XLXmlData* xmlData);

        /**
         * @brief Copy constructor. Default implementation used.
         * @param other The object to copy.
         */
        XLXmlFile(const XLXmlFile& other) = default;

        /**
         * @brief Move constructor. Default implementation used.
         * @param other The object to move.
         */
        XLXmlFile(XLXmlFile&& other) noexcept = default;

        /**
         * @brief Destructor. Default implementation used.
         */
        ~XLXmlFile();

        /**
         * @brief The copy assignment operator. The default implementation has been used.
         * @param other The object to copy.
         * @return A reference to the new object.
         */
        XLXmlFile& operator=(const XLXmlFile& other) = default;

        /**
         * @brief The move assignment operator. The default implementation has been used.
         * @param other The object to move.
         * @return A reference to the new object.
         */
        XLXmlFile& operator=(XLXmlFile&& other) noexcept = default;

    protected:    // ===== PROTECTED MEMBER FUNCTIONS
        /**
         * @brief Method for getting the XML data represented by the object.
         * @return A std::string with the XML data.
         */
        std::string xmlData() const;

        /**
         * @brief Provide the XML data represented by the object.
         * @param xmlData A std::string with the XML data.
         */
        void setXmlData(const std::string& xmlData);

        /**
         * @brief This function returns the relationship ID (the ID used in the XLRelationships objects) for the object.
         * @return A std::string with the ID. Not all spreadsheet objects may have a relationship ID. In those cases an empty string is
         * returned.
         */
        std::string relationshipID() const;

        /**
         * @brief This function provides access to the parent XLDocument object.
         * @return A reference to the parent XLDocument object.
         */
        XLDocument& parentDoc();

        /**
         * @brief This function provides access to the parent XLDocument object.
         * @return A const reference to the parent XLDocument object.
         */
        const XLDocument& parentDoc() const;

        /**
         * @brief This function provides access to the underlying XMLDocument object.
         * @return A reference to the XMLDocument object.
         */
        XMLDocument& xmlDocument();

        /**
         * @brief This function provides access to the underlying XMLDocument object.
         * @return A const reference to the XMLDocument object.
         */
        const XMLDocument& xmlDocument() const;

    private:                              // ===== PRIVATE MEMBER VARIABLES
        XLXmlData* m_xmlData { nullptr }; /**< The underlying XML data object. */
    };
}    // namespace OpenXLSX

#pragma warning(pop)
#endif    // OPENXLSX_XLXMLFILE_HPP

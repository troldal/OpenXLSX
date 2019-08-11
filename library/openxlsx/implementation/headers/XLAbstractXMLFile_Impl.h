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

#ifndef OPENXLSX_IMPL_XLABSTRACTXMLFILE_H
#define OPENXLSX_IMPL_XLABSTRACTXMLFILE_H

// ===== Standard Library Includes ===== //
#include <iostream>
#include <memory>
#include <string>
#include <map>

// ===== OpenXLSX Includes ===== //
#include "XLXml_Impl.h"

namespace OpenXLSX::Impl
{
    class XLDocument;

    //======================================================================================================================
    //========== XLAbstractXMLFile Class ===================================================================================
    //======================================================================================================================

    /**
     * @brief The XLAbstractXMLFile is an pure abstract class, which provides an interface
     * for derived classes to use. It functions as an ancestor to all classes which are represented by an .xml
     * file in an .xlsx package
     */
    class XLAbstractXMLFile
    {

        //----------------------------------------------------------------------------------------------------------------------
        //           Public Member Functions
        //----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief Constructor. Creates an object using the parent XLDocument object, the relative file path
         * and a data object as input.
         * @param parent
         * @param filePath The path of the XML file, relative to the root.
         * @param xmlData An std::string object with the XML data to be represented by the object.
         */
        explicit XLAbstractXMLFile(XLDocument& parent, std::string filePath, const std::string& xmlData = "");

        /**
         * @brief Copy constructor. Default (shallow) implementation used.
         */
        XLAbstractXMLFile(const XLAbstractXMLFile&) = delete;

        /**
         * @brief Destructor. Default implementation used.
         */
        virtual ~XLAbstractXMLFile() = default;

        /**
         * @brief The assignment operator. The default implementation has been used.
         * @return A reference to the new object.
         */
        XLAbstractXMLFile& operator=(const XLAbstractXMLFile&) = delete;

        /**
         * @brief
         * @return
         * @todo implement a "safe operator bool" instead
         */
        virtual operator bool() const;

        /**
         * @brief Provide the XML data represented by the object.
         * @param xmlData A std::string with the XML data.
         */
        virtual void SetXmlData(const std::string& xmlData);

        /**
         * @brief Method for getting the XML data represented by the object.
         * @return A std::string with the XML data.
         */
        virtual const std::string& GetXmlData() const;

        /**
         * @brief Commit the XML data to the zipped .xlsx package.
         */
        virtual void WriteXMLData();

        /**
         * @brief Delete the XML file from the zipped .xlsx package.
         */
        virtual void DeleteXMLData();

        /**
         * @brief Get the path of the current file.
         * @return A string with the path of the file.
         * @note This method is final, i.e. it cannot be overridden.
         */
        virtual const std::string& FilePath() const final;


        //----------------------------------------------------------------------------------------------------------------------
        //           Protected Member Functions
        //----------------------------------------------------------------------------------------------------------------------

    protected:

        /**
         * @brief This method returns the underlying XMLDocument object.
         * @return A pointer to the XMLDocument object.
         * @note This method is final, i.e. it cannot be overridden.
         */
        virtual XMLDocument* XmlDocument() final;

        /**
         * @brief This method returns the underlying XMLDocument object.
         * @return A pointer to the const XMLDocument object.
         * @note This method is final, i.e. it cannot be overridden.
         */
        virtual const XMLDocument* XmlDocument() const final;

        /**
         * @brief The parseXMLData method is used to map or copy the XML data to the internal data structures.
         * @return true on success; otherwise false.
         * @note This is a pure virtual function, meaning that this must be implemented in derived classes.
         */
        virtual bool ParseXMLData() = 0;

        //----------------------------------------------------------------------------------------------------------------------
        //           Private Member Variables
        //----------------------------------------------------------------------------------------------------------------------


    private:

        std::string m_path; /**< */
        XLDocument& m_parentDocument; /**< */
        XMLDocument m_xmlDocument; /**< A pointer to the underlying XMLDocument resource*/
        mutable std::string m_xmlData; /**< A std::string with the XML data. This is only updated when GetXMLData() is called */

    };
}  // namespace OpenXLSX::Impl

#endif //OPENXLSX_IMPL_XLABSTRACTXMLFILE_H

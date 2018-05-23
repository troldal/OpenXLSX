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

#ifndef OPENXL_XLABSTRACTFILE_H
#define OPENXL_XLABSTRACTFILE_H

#include <iostream>
#include <memory>
#include <string>
#include <map>
#include "../Utilities/XML/XML.h"

namespace OpenXLSX
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
        friend class XLWorkbook;
        friend class XLColumn;

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief Constructor. Creates an object using the parent XLDocument object, the relative file path
         * and a data object as input.
         * @param root A string with the root path
         * @param filePath The path of the XML file, relative to the root.
         * @param xmlData An std::string object with the XML data to be represented by the object.
         */
        explicit XLAbstractXMLFile(XLDocument &parent,
                                   const std::string &filePath,
                                   const std::string &xmlData = "");

        /**
         * @brief Copy constructor. Default (shallow) implementation used.
         */
        XLAbstractXMLFile(const OpenXLSX::XLAbstractXMLFile &) = delete;

        /**
         * @brief Destructor. Default implementation used.
         */
        virtual ~XLAbstractXMLFile() = default;

        /**
         * @brief The assignment operator. The default implementation has been used.
         * @return A reference to the new object.
         */
        XLAbstractXMLFile &operator=(const OpenXLSX::XLAbstractXMLFile &) = delete;

        /**
         * @brief Provide the XML data represented by the object.
         * @param xmlData A std::string with the XML data.
         */
        virtual void SetXmlData(const std::string &xmlData);

        /**
         * @brief Method for getting the XML data represented by the object.
         * @return A std::string with the XML data.
         */
        virtual std::string GetXmlData() const;

        /**
         * @brief Commit the XML data to the zipped .xlsx package.
         */
        virtual void CommitXMLData();

        /**
         * @brief Delete the XML file from the zipped .xlsx package.
         */
        virtual void DeleteXMLData();

        /**
         * @brief Get the path of the current file.
         * @return A string with the path of the file.
         * @note This method is final, i.e. it cannot be overridden.
         */
        virtual const std::string &FilePath() const final;

        /**
         * @brief Prints the XML document to the standard output
         * @todo consider having a generic function for both saving and printing.
         */
        virtual void Print() const final;


//----------------------------------------------------------------------------------------------------------------------
//           Protected Member Functions
//----------------------------------------------------------------------------------------------------------------------

    protected:

        /**
         * @brief This method returns the underlying XMLDocument object.
         * @return A pointer to the XMLDocument object.
         * @note This method is final, i.e. it cannot be overridden.
         */
        virtual XMLDocument *XmlDocument() final;

        /**
         * @brief This method returns the underlying XMLDocument object.
         * @return A pointer to the const XMLDocument object.
         * @note This method is final, i.e. it cannot be overridden.
         */
        virtual const XMLDocument *XmlDocument() const final;

        /**
         * @brief Set the 'm_isModified' flag, meaning that the underlying XML file needs saving.
         */
        virtual void SetModified();

        /**
         * @brief Determine id the XML file has been modified or not.
         * @return true if the file has been modified; otherwise false.
         */
        virtual bool IsModified();

        /**
         * @brief The parseXMLData method is used to map or copy the XML data to the internal data structures.
         * @return true on success; otherwise false.
         * @note This is a pure virtual function, meaning that this must be implemented in derived classes.
         */
        virtual bool ParseXMLData() = 0;

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Variables
//----------------------------------------------------------------------------------------------------------------------

    protected:
        std::unique_ptr<XMLDocument> m_xmlDocument; /**< A pointer to the underlying XMLDocument resource*/

    private:

        XLDocument &m_parentDocument; /**< */
        std::string m_path; /**< */

        mutable std::map<std::string, XLAbstractXMLFile *> m_childXmlDocuments; /**< A std::map with the child XML documents. */
        mutable bool m_isModified; /**< A bool indicating if the document has been modified and needs saving. It is mutable, and can therefore be modified in const methods. */

    };
}

#endif //OPENXL_XLABSTRACTFILE_H

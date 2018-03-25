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
#include <string>
#include "XML/XMLDocument.h"

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
         * @param data An std::string object with the XML data to be represented by the object.
         */
        explicit XLAbstractXMLFile(const std::string &root,
                                   const std::string &filePath,
                                   const std::string &data = "");

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
         * @brief Get the path of the current file.
         * @return A string with the path of the file.
         * @note This method is final, i.e. it cannot be overridden.
         */
        virtual const std::string &FilePath() const final;

        /**
         * @brief Set the path of the file. This will be used when loading and saving the XML document.
         * @param filePath The path of the file.
         * @note This method is final, i.e. it cannot be overridden.
         * @todo Consider saving the file to the new path when setting.
         */
        virtual void SetFilePath(const std::string &filePath) final;

        /**
         * @brief Read the XML data using the XML library.
         * @return true if the reading was successful; otherwise false.
         * @note This method is final, i.e. it cannot be overridden.
         * @todo Currently, it always returns true. Fix, so that it checks the results.
         */
        virtual bool LoadXMLData() final;

        /**
         * @brief Saves the XML data to the file, using the m_filePath as the path. If the file exists, it will be overwritten.
         * @todo check that an existing file will indeed be overwritten.
        */
        virtual void SaveXMLData() const final;

        /**
         * @brief Prints the XML document to the standard output
         * @todo consider having a generic function for both saving and printing.
         */
        virtual void Print() const final;

        /**
         * @brief Get the parent XLDocument object
         * @return A pointer to the parent object
         */
        //virtual XLDocument &ParentDocument() const;


//----------------------------------------------------------------------------------------------------------------------
//           Protected Member Functions
//----------------------------------------------------------------------------------------------------------------------

    protected:



        /**
         * @brief This method returns the underlying XMLDocument object.
         * @return A pointer to the XMLDocument object.
         * @note This method is final, i.e. it cannot be overridden.
         */
        virtual XMLDocument &XmlDocument() final;

        /**
         * @brief
         * @return
         */
        virtual const XMLDocument &XmlDocument() const final;

        /**
         * @brief Set the 'isModified' flag, meaning that the underlying XML file needs saving.
         */
        virtual void SetModified();

        /**
         * @brief
         * @return
         */
        virtual bool IsModified();

        /**
         * @brief Delete the XML file represented by the object.
         * @return true if successful, otherwise false.
         * @todo Currently the method always returns true.
         */
        virtual bool DeleteFile();

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

        std::unique_ptr<XMLDocument> m_xmlDocument; /**< A pointer to the underlying XMLDocument resource*/
        std::string m_filePath; /**< A std::string with the full path to the XML file.*/
        std::string m_root;
        //XLDocument &m_parentDocument; /**< A reference to the parent XLDocument object*/

        mutable std::map<std::string, XLAbstractXMLFile *>
            m_childXmlDocuments; /**< A std::map with the child XML documents. */
        mutable bool
            m_isModified; /**< A bool indicating if the document has been modified and needs saving. It is mutable, and can therefore be modified in const methods. */

    };
}

#endif //OPENXL_XLABSTRACTFILE_H

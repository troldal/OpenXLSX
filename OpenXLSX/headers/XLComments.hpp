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

#ifndef OPENXLSX_XLCOMMENTS_HPP
#define OPENXLSX_XLCOMMENTS_HPP

// ===== External Includes ===== //
#include <cstdint>    // uint8_t, uint16_t, uint32_t
#include <ostream>    // std::basic_ostream

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
// #include "XLDocument.hpp"
#include "XLDrawing.hpp"   // XLVmlDrawing
#include "XLException.hpp"
#include "XLXmlData.hpp"
#include "XLXmlFile.hpp"

// TODO:
//   add to [Content_Types].xml:
//     <Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>

namespace OpenXLSX
{
    /**
     * @brief The XLComments class is the base class for worksheet comments
     */
    class OPENXLSX_EXPORT XLComments : public XLXmlFile
    {
        friend class XLWorksheet;   // for access to XLXmlFile::getXmlPath
    public:
        /**
         * @brief Constructor
         */
        XLComments();

        /**
         * @brief The constructor.
         * @param xmlData the source XML of the comments file
         */
        XLComments(XLXmlData* xmlData);

        /**
         * @brief The copy constructor.
         * @param other The object to be copied.
         */
        XLComments(const XLComments& other);

        /**
         * @brief
         * @param other
         */
        XLComments(XLComments&& other) noexcept;

        /**
         * @brief The destructor
         * @note The default destructor is used, since cleanup of pointer data members is not required.
         */
        ~XLComments() = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLComments& operator=(XLComments&& other) noexcept;

        /**
         * @brief Assignment operator
         * @return A reference to the new object.
         */
        XLComments& operator=(const XLComments&);

        /**
         * @brief associate the worksheet's VML drawing object with the comments so it can be modified from here
         * @param vmlDrawing the worksheet's previously created XLVmlDrawing object
         * @return true upon success
         */
        bool setVmlDrawing(XLVmlDrawing &vmlDrawing);

    private: // helper functions with repeating code
        XMLNode authorNode(uint16_t index) const;
        XMLNode commentNode(const std::string& cellRef) const;

    public:

        uint16_t authorCount() const;

        std::string author(uint16_t index) const;

        bool deleteAuthor(uint16_t index);

        uint16_t addAuthor(const std::string& authorName);

        /**
         * @brief get the amount of comments
         * @return the amount of comments for the worksheet
         */
        size_t count() const;

        uint16_t authorId(const std::string& cellRef) const;

        bool deleteComment(const std::string& cellRef);

        /**
         * @brief get the comment (if any) for the referenced cell
         * @param cellRef the cell address to check
         * @return the comment for this cell - an empty string if no comment is set
         */
        std::string get(std::string const& cellRef) const;

        /**
         * @brief set the comment for the referenced cell
         * @param cellRef the cell address to set
         * @param comment set this text as comment for the cell
         * @param authorId set this author
         * @return true upon success, false on failure
         */
        bool set(std::string const& cellRef, std::string const& comment, uint16_t authorId = 0);

        /**
         * @brief get the XLShape object for this comment
         */
        XLShape shape(std::string const& cellRef);

        /**
         * @brief Print the XML contents of this XLComments instance using the underlying XMLNode print function
         */
        void print(std::basic_ostream<char>& ostr) const;

    private:
        XMLNode m_authors{};
        XMLNode m_commentList{};
        std::unique_ptr<XLVmlDrawing> m_vmlDrawing;
        inline static const std::vector< std::string_view > m_nodeOrder = {      // comments XML node required child sequence
            "authors",
            "commentList"
        };
    };
}    // namespace OpenXLSX

#endif    // OPENXLSX_XLCOMMENTS_HPP

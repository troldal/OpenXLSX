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

#ifndef OPENXLSX_XLMERGECELLS_HPP
#define OPENXLSX_XLMERGECELLS_HPP

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#   pragma warning(push)
#   pragma warning(disable : 4251)
#   pragma warning(disable : 4275)
#endif // _MSC_VER

#include <deque>
#include <limits>     // std::numeric_limits
#include <memory>     // std::unique_ptr
#include <ostream>    // std::basic_ostream
#include <string>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLCellReference.hpp"
#include "XLXmlParser.hpp" // XMLNode, pugi node types

namespace OpenXLSX
{
    constexpr size_t XLMaxMergeCells = (std::numeric_limits< int32_t >::max)();    // pull request #261: wrapped max in parentheses to prevent expansion of windows.h "max" macro

    /**
     * @brief This class encapsulate the Excel concept of <mergeCells>. Each worksheet that has merged cells has a list of
     * (empty) <mergeCell> elements within that array, with a sole attribute ref="..." with ... being a range reference, e.g. A1:B5
     */
    class OPENXLSX_EXPORT XLMergeCells
    {
        //----------------------------------------------------------------------------------------------------------------------
        //           Public Member Functions
        //----------------------------------------------------------------------------------------------------------------------

    public:
        /**
         * @brief
         */
        XLMergeCells() = default;

        /**
         * @brief
         * @param node The <mergeCells> node of the worksheet document - must not be an empty node
         */
        explicit XLMergeCells(const XMLNode& node);

        /**
         * @brief Destructor
         */
        ~XLMergeCells();

        /**
         * @brief
         * @param other
         */
        XLMergeCells(const XLMergeCells& other)
        {
            m_mergeCellsNode = other.m_mergeCellsNode ? std::make_unique<XMLNode>( *other.m_mergeCellsNode ) : std::unique_ptr<XMLNode> {};
            m_referenceCache = other.m_referenceCache;
        }

        /**
         * @brief
         * @param other
         */
        XLMergeCells(XLMergeCells&& other)
        {
            m_mergeCellsNode = std::move( other.m_mergeCellsNode );
            m_referenceCache = std::move( other.m_referenceCache );
        }

        /**
         * @brief
         * @param other
         * @return
         */
        XLMergeCells& operator=(const XLMergeCells& other)
        {
            m_mergeCellsNode = other.m_mergeCellsNode ? std::make_unique<XMLNode>( *other.m_mergeCellsNode ) : std::unique_ptr<XMLNode> {};
            m_referenceCache = other.m_referenceCache;
            return *this;
        }

        /**
         * @brief
         * @param other
         * @return
         */
        XLMergeCells& operator=(XLMergeCells&& other)
        {
            m_mergeCellsNode = std::move( other.m_mergeCellsNode );
            m_referenceCache = std::move( other.m_referenceCache );
            return *this;
        }

        bool uninitialized() const { return ( !m_mergeCellsNode ); }

        /**
         * @brief get the index of a <mergeCell> entry by its reference
         * @param reference the reference to search for
         * @return -1 if no such reference exists, 0-based index otherwise
         */
        int32_t findMerge(const std::string& reference) const;

        /**
         * @brief test if a mergeCell with reference exists, equivalent to findMerge(reference) >= 0
         * @param reference the reference to find
         * @return true if reference exists, false otherwise
         */
        bool mergeExists(const std::string& reference) const;

        /**
         * @brief get the index of a <mergeCell> entry of which cellReference is a part
         * @param cellRef the cell reference (string or XLCellReference) to search for in the merged ranges
         * @return -1 if no such reference exists, 0-based index otherwise
         */
        int32_t findMergeByCell(const std::string& cellRef) const;
        int32_t findMergeByCell(XLCellReference cellRef) const;

        /**
         * @brief get the amount of entries in <mergeCells>
         * @return the count of defined cell merges
         */
        size_t count() const;

        /**
         * @brief return the cell reference string for the given index
         * @param index
         * @return
         */
        const char* merge(int32_t index) const;

        /**
         * @brief Operator overload: allow [] as shortcut access to merge
         */
        const char* operator[](int32_t index) const { return merge(index); }

        /**
         * @brief Append a new merge to the list of merges
         * @param reference The reference to append.
         * @return An int32_t with the index of the appended string
         * @throws XLInputException if the reference would overlap with an existing reference
         */
        int32_t appendMerge(const std::string& reference);

        /**
         * @brief Delete the merge at the given index.
         * @param index The index to delete
         * @note Previously obtained merge indexes will be invalidated when calling deleteMerge
         * @throws XLInputException if the index does not exist
         */
        void deleteMerge(int32_t index);

        /**
         * @brief print the XML contents of the mergeCells array using the underlying XMLNode print function
         */
        void print(std::basic_ostream<char>& ostr) const;

    private:
        std::unique_ptr<XMLNode> m_mergeCellsNode; /**< An XMLNode object with the mergeCells item */
        std::deque<std::string> m_referenceCache;
    };
}    // namespace OpenXLSX

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#   pragma warning(pop)
#endif // _MSC_VER

#endif    // OPENXLSX_XLMERGECELLS_HPP

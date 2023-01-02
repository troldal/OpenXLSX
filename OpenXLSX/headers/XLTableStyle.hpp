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

    Written by Akira SHIMAHARA

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

#ifndef OPENXLSX_XLTABLESTYLE_HPP
#define OPENXLSX_XLTABLESTYLE_HPP

#pragma warning(push)
//#pragma warning(disable : 4251)
//#pragma warning(disable : 4275)

// ===== External Includes ===== //
#include <string>
#include <memory>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    class XLTable;

    class OPENXLSX_EXPORT XLTableStyle
    {
    public:
        /**
         * @brief The constructor. 
         * @param dataNode XMLNode of the column
         */
      XLTableStyle(const XMLNode& dataNode, XLTable* table);

         /**
         * @brief Copy Constructor
         * @note The copy constructor is explicitly deleted
         */
      XLTableStyle(const XLTableStyle& other);

        /**
         * @brief Move Constructor
         * @note The move constructor has been explicitly deleted.
         */
       XLTableStyle(XLTableStyle&& other) noexcept;

        //XLtableStyle(constXLTableStyle&) = delete;
        //XLtableStyle& operator=(constXLTableStyle&) = delete;
        ~XLTableStyle();

        /**
         * @brief Copy assignment operator.
         * @note The copy assignment operator is explicitly deleted.
         */
       XLTableStyle& operator=(const XLTableStyle& other);

        /**
         * @brief Move assignment operator.
         * @note The move assignment operator has been explicitly deleted.
         */
       XLTableStyle& operator=(XLTableStyle&& other) noexcept;


        /**
         * @brief 
         * @return the style name
         */
        const std::string style() const;

        /**
         * @brief 
         * @return true id the colum stripes are shown
         */
        bool columnStripes() const;

         /**
         * @brief 
         * @param  show true to show the stripes, default is true
         */ 
        void showColumnStripes(bool show = true) const;

        /**
         * @brief 
         * @return true if the colum stripes are shown
         */
        bool rowStripes() const;

        /**
         * @brief 
         * @param  show true to show the stripes, default is true
         */ 
        void showRowStripes(bool show = true) const;

        /**
         * @brief 
         * @return true fd the first column is highligted
         */ 
        bool firstColumnHighlighted() const;

        /**
         * @brief 
         * @param  show true highlight the column, default is true
         */ 
        void showFirstColumnHighlighted(bool show = true) const;

        /**
         * @brief 
         * @return true fd the first column is highligted
         */ 
        bool lastColumnHighlighted() const;

        /**
         * @brief 
         * @param  show true highlight the column, default is true
         */ 
        void showLastColumnHighlighted(bool show = true) const;

        /**
         * @brief 
         * @param name the style name. If not availble, do nothing
         */
        void setStyle(const std::string& style) const;
    
    //----------------------------------------------------------------------------------------------------------------------
    //           Protected
    //----------------------------------------------------------------------------------------------------------------------  
    protected:
        /**
         * @brief return a vector of available table style
         * @param style style to be checked if available
         */
        bool isTableStyleAvailable(const std::string& style) const;
        

    private:
        std::unique_ptr<XMLNode>    m_dataNode;
        XLTable*                    m_table;
    };
}    // namespace OpenXLSX

//#pragma warning(pop)
#endif    // OPENXLSX_XLSTYLE_HPP

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

#ifndef OPENXLSX_XLNAMEDRANGE_HPP
#define OPENXLSX_XLNAMEDRANGE_HPP

#pragma warning(push)
#pragma warning(disable : 4251)
#pragma warning(disable : 4275)

// ===== External Includes ===== //
#include <string>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLCellRange.hpp"


namespace OpenXLSX
{
  /**
     * @brief This class derivate from cell range class,
     * it just add the method to creating, ....
     */
  class OPENXLSX_EXPORT XLNamedRange : public XLCellRange
  {
    public:
        /**
         * @brief Constructor 
         * @param rangeName Name of the range to be created
         * @param reference Reference of the cell/range to be named: Sheet1!$I$17:$I$19
         * @param localSheetId Id of the sheet where the name is defined, default is 0 (global)
         */
        XLNamedRange(const std::string& name,
                      const std::string& reference,
                      uint32_t localSheetId,
                      const XLCellRange& rng);
        /**
         * @brief Destructor [default]
         * @note This implements the default destructor.
         */
        ~XLNamedRange();

        /**
         * @brief The copy constructor
         * @param other The made range object to be copied and assigned.
         * @return A reference to the new object.
         */
        XLNamedRange(const XLNamedRange& other);

        /**
         * @brief The copy assignment operator [default]
         * @param other The named range to be copied and assigned.
         * @return A reference to the new object.
         * @throws A std::range_error if the source range and destination range are of different size and shape.
         * @note This implements the default copy assignment operator.
         */
        XLNamedRange& operator=(const XLNamedRange& other);

        /**
         * @brief The move assignment operator
         * @param other The named range to be moved and assigned.
         * @return A reference to the new object.
         * @note This implements the default move assignment operator.
         */
         XLNamedRange& operator=(XLNamedRange&& other) noexcept;

        /**
         * @brief  
         * @return The name of the named range.
         */
        const std::string& name() const;

         /**
         * @brief  
         * @return The reference of the range.
         */
        const std::string& reference() const;

         /**
         * @brief  
         * @return The name of the named range.
         */
        uint32_t localSheetId() const;

          /**
         * @brief  
         * @return The first cell of the range.
         */
        XLCell firstCell() const;
    
    private:
        uint32_t          m_localSheetId;
        std::string       m_name;
        std::string       m_reference;
  };   
}    // namespace OpenXLSX

#pragma warning(pop)
#endif    // OPENXLSX_XLNAMEDRANGE_HPP

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

#ifndef OPENXLSX_XLFORMULA_HPP
#define OPENXLSX_XLFORMULA_HPP

#pragma warning(push)
#pragma warning(disable : 4251)
#pragma warning(disable : 4275)

// ===== External Includes ===== //
#include <cstdint>
#include <iostream>
#include <string>
#include <variant>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLException.hpp"
#include "XLXmlParser.hpp"

// ========== CLASS AND ENUM TYPE DEFINITIONS ========== //
namespace OpenXLSX
{
    //---------- Forward Declarations ----------//
    class XLFormulaProxy;
    class XLCell;

    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLFormula
    {
        //---------- Friend Declarations ----------//

        friend bool          operator==(const XLFormula& lhs, const XLFormula& rhs);
        friend bool          operator!=(const XLFormula& lhs, const XLFormula& rhs);
        friend std::ostream& operator<<(std::ostream& os, const XLFormula& value);

    public:

        /**
         * @brief
         */
        XLFormula();

        /**
         * @brief
         * @tparam T
         * @param formula
         */
        template<typename T,
                 typename std::enable_if<std::is_constructible_v<T, char*>>::type* = nullptr>
        explicit XLFormula(T formula) {

            if constexpr (std::is_same_v<T, const char *>)
                m_formulaString = formula;
            else
                m_formulaString = formula.c_str();
        }

        /**
         * @brief
         * @param other
         */
        XLFormula(const XLFormula& other);

        /**
         * @brief
         * @param other
         */
        XLFormula(XLFormula&& other) noexcept;

        /**
         * @brief
         */
        ~XLFormula();

        /**
         * @brief
         * @param other
         * @return
         */
        XLFormula& operator=(const XLFormula& other);

        /**
         * @brief
         * @param other
         * @return
         */
        XLFormula& operator=(XLFormula&& other) noexcept;

        /**
         * @brief
         * @tparam T
         * @param formula
         * @return
         */
        template<typename T,
                 typename std::enable_if<std::is_constructible_v<T, char*>>::type* = nullptr>
        XLFormula& operator=(T formula) {
            XLFormula temp(formula);
            std::swap(*this, temp);
            return *this;
        }

        /**
         * @brief
         * @tparam T
         * @param formula
         */
        template<typename T,
                 typename std::enable_if<std::is_constructible_v<T, char*>>::type* = nullptr>
        void set(T formula) {
            *this = formula;
        }

        /**
         * @brief
         * @return
         */
        std::string get() const;

        /**
         * @brief
         * @tparam T
         * @return
         */
        template<typename T,
                 typename std::enable_if<std::is_constructible_v<T, const char*>>::type* = nullptr>
        operator T() const {
            return m_formulaString.c_str();
        }

        /**
         * @brief
         * @return
         */
        XLFormula& clear();

    private:
        std::string m_formulaString; /**< */
    };

    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLFormulaProxy
    {
        friend class XLCell;
        friend class XLFormula;

    public:

        /**
         * @brief
         */
        ~XLFormulaProxy();

        /**
         * @brief
         * @param other
         * @return
         */
        XLFormulaProxy& operator=(const XLFormulaProxy& other);

        /**
         * @brief
         * @tparam T
         * @param formula
         * @return
         */
        template<typename T,
                 typename std::enable_if<std::is_same_v<T, XLFormula> || std::is_constructible_v<T, const char*>>::type* = nullptr>
        XLFormulaProxy& operator=(T formula) {

            if constexpr (std::is_same_v<T, XLFormula>)
                setFormulaString(formula.get().c_str());
            else if constexpr (std::is_same_v<T, const char*>)
                setFormulaString(formula);
            else
                setFormulaString(formula.c_str());

            return *this;
        }

        /**
         * @brief
         * @tparam T
         * @param formula
         */
        template<typename T,
                 typename std::enable_if<std::is_constructible_v<T, char*>>::type* = nullptr>
        void set(T formula) {
            *this = formula;
        }

        /**
         * @brief
         * @return
         */
        std::string get() const;

        /**
         * @brief
         * @tparam T
         * @return
         */
        template<typename T,
                 typename std::enable_if<std::is_constructible_v<T, const char*>>::type* = nullptr>
        operator T() const {
            return get().c_str();
        }

        /**
         * @brief
         * @return
         */
        XLFormulaProxy& clear();

        /**
         * @brief
         * @return
         */
        operator XLFormula();    // NOLINT

    private:

        /**
         * @brief
         * @param cell
         * @param cellNode
         */
        XLFormulaProxy(XLCell* cell, XMLNode* cellNode);

        /**
         * @brief
         * @param other
         */
        XLFormulaProxy(const XLFormulaProxy& other);

        /**
         * @brief
         * @param other
         */
        XLFormulaProxy(XLFormulaProxy&& other) noexcept;

        /**
         * @brief
         * @param other
         * @return
         */
        XLFormulaProxy& operator=(XLFormulaProxy&& other) noexcept;

        /**
         * @brief
         * @param formulaString
         */
        void setFormulaString(const char* formulaString);

        /**
         * @brief
         * @return
         */
        XLFormula getFormula() const;

        //---------- Private Member Variables ---------- //
        XLCell*  m_cell;     /**< Pointer to the owning XLCell object. */
        XMLNode* m_cellNode; /**< Pointer to corresponding XML cell node. */
    };
} // namespace OpenXLSX

// ========== FRIEND FUNCTION IMPLEMENTATIONS ========== //
namespace OpenXLSX
{
    /**
     * @brief
     * @param lhs
     * @param rhs
     * @return
     */
    inline bool operator==(const XLFormula& lhs, const XLFormula& rhs)
    {
        return lhs.m_formulaString == rhs.m_formulaString;
    }

    /**
     * @brief
     * @param lhs
     * @param rhs
     * @return
     */
    inline bool operator!=(const XLFormula& lhs, const XLFormula& rhs)
    {
        return lhs.m_formulaString != rhs.m_formulaString;
    }

    /**
     * @brief
     * @param os
     * @param value
     * @return
     */
    inline std::ostream& operator<<(std::ostream& os, const XLFormula& value)
    {
        return os << value.m_formulaString;
    }
}    // namespace OpenXLSX

#endif    // OPENXLSX_XLFORMULA_HPP

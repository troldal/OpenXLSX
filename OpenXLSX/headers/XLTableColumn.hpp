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

#ifndef OPENXLSX_XLTABLECOLUMN_HPP
#define OPENXLSX_XLTABLECOLUMN_HPP

#pragma warning(push)
//#pragma warning(disable : 4251)
//#pragma warning(disable : 4275)

// ===== External Includes ===== //
#include <string>
#include <memory>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLXmlParser.hpp"
#include "XLCellRange.hpp"

namespace OpenXLSX
{
    class XLTable;


    /**
     * @brief The XLTableColumnFormulaProxy class is used for proxy (or placeholder) objects for XLTableColumnd formulas objects.
     * @details The XLTableColumnFormulaProxy class is used for proxy (or placeholder) objects forXLTableColumnd formulas objects.
     * The purpose is to enable implicit conversion during assignment operations. XLTableColumnd formulas objects
     * can not be constructed manually by the user, only through XLTableColumn objects.
     */
    class OPENXLSX_EXPORT XLTableColumnProxy
    {
        friend class XLTableColumn;

    public:
        //---------- Public Member Functions ----------//

        /**
         * @brief Destructor
         */
        ~XLTableColumnProxy();

        /**
         * @brief Copy assignment operator.
         * @param other XLCellValueProxy object to be copied.
         * @return A reference to the current object.
         */
        XLTableColumnProxy& operator=(const XLTableColumnProxy& other);

        /**
         * @brief Templated assignment operator
         * @tparam T The type of numberValue assigned to the object.
         * @param formula The formula to be set
         * @return A reference to the current object.
         */      
        virtual XLTableColumnProxy& operator=(const std::string& formula) = 0;

        void set(const std::string& formula)
        {
            *this = formula;
        }

        std::string get() const
        {
            return getFormula();
        }

        //XLTableColumnProxy& clear();

        /**
         * @brief Set the cell value to a error state.
         * @return A reference to the current object.
         */
        XLTableColumnProxy& setError(const std::string & error);

         /**
         * @brief Implicitly convert the XLTableColumnProxy object to a string object.
         * @return a string corresponding to the formula.
         */
        operator std::string() { return getFormula(); };  // NOLINT

        /**
         * @brief Implicitly convert the XLTableColumnProxy object to a string object.
         * @return a string corresponding to the formula.
         */
        operator std::string () const { return getFormula(); };

    protected:
        //---------- Private Member Functions ---------- //

        /**
         * @brief Constructor
         * @param attr Pointer to the corresponding XML attribute object.
         */
        XLTableColumnProxy(const XMLNode& dataNode, 
                        const std::string& attr, const XLTable& table);

        /**
         * @brief Copy constructor
         * @param other Object to be copied.
         */
        XLTableColumnProxy(const XLTableColumnProxy& other);

        /**
         * @brief Move constructor
         * @param other Object to be moved.
         */
        XLTableColumnProxy(XLTableColumnProxy&& other) noexcept;

        /**
         * @brief Move assignment operator
         * @param other Object to be moved
         * @return Reference to moved-to pbject.
         */
        XLTableColumnProxy& operator=(XLTableColumnProxy&& other) noexcept;

        /**
         * @brief Set the cell to a string value.
         * @param formula The value to be set.
         */
        virtual void setFormula(const std::string& formula) = 0;

        /**
         * @brief Get a copy of the XLCellValue object for the cell.
         * @return An XLCellValue object.
         */
        virtual std::string getFormula() const = 0;

          //---------- Private Member Variables ---------- //

        std::shared_ptr<XMLNode>    m_node; /**< Pointer to corresponding XML attribute */
        std::string                 m_attribute;
        const XLTable&              m_table;
    }; // Class

    class OPENXLSX_EXPORT XLTableColumnTotalProxy : public XLTableColumnProxy
    {
        friend class XLTableColumn;

    public:
        //---------- Public Member Functions ----------//

        /**
         * @brief Destructor
         */
        ~XLTableColumnTotalProxy()= default;

        /**
         * @brief Templated assignment operator
         * @param formula The formula to be set
         * @return A reference to the current object.
         */      
        XLTableColumnProxy& operator=(const std::string& formula);

        void clear();

    private:
        //---------- Private Member Functions ---------- //

        /**
         * @brief Constructor
         * @param attr Pointer to the corresponding XML attribute object.
         */
        XLTableColumnTotalProxy(const XMLNode& dataNode, 
                        const std::string& attr, const XLTable& table)
                        : XLTableColumnProxy(dataNode, attr,table) {};

        /**
         * @brief Set the cell to a string value.
         * @param formula The value to be set.
         */
        void setFormula(const std::string& formula);

        /**
         * @brief Get a copy of the XLCellValue object for the cell.
         * @return An XLCellValue object.
         */
        std::string getFormula() const;
    }; // Class 

     class OPENXLSX_EXPORT XLTableColumnFormulaProxy : public XLTableColumnProxy
    {
        friend class XLTableColumn;

    public:
        //---------- Public Member Functions ----------//

        /**
         * @brief Destructor
         */
        ~XLTableColumnFormulaProxy()= default;

        /**
         * @brief Templated assignment operator
         * @param formula The formula to be set
         * @return A reference to the current object.
         */      
        XLTableColumnProxy& operator=(const std::string& formula);

        void clear();

    private:
        //---------- Private Member Functions ---------- //

        /**
         * @brief Constructor
         * @param attr Pointer to the corresponding XML attribute object.
         */
        XLTableColumnFormulaProxy(const XMLNode& dataNode, 
                        const std::string& attr, const XLTable& table)
                        : XLTableColumnProxy(dataNode, attr,table) {};

        /**
         * @brief Set the cell to a string value.
         * @param formula The value to be set.
         */
        void setFormula(const std::string& formula);

        /**
         * @brief Get a copy of the XLCellValue object for the cell.
         * @return An XLCellValue object.
         */
        std::string getFormula() const;
    }; // Class 


    class OPENXLSX_EXPORT XLTableColumn
    {
    public:
        /**
         * @brief The constructor. 
         * @param dataNode XMLNode of the column
         */
        XLTableColumn(const XMLNode& dataNode, const XLTable& table);

         /**
         * @brief Copy Constructor
         * @note The copy constructor is explicitly deleted
         */
        XLTableColumn(const XLTableColumn& other);

        /**
         * @brief Move Constructor
         * @note The move constructor has been explicitly deleted.
         */
        XLTableColumn(XLTableColumn&& other) noexcept;

        //XLTableColumn(const XLTableColumn&) = delete;
        //XLTableColumn& operator=(const XLTableColumn&) = delete;
        ~XLTableColumn();

        /**
         * @brief Copy assignment operator.
         * @note The copy assignment operator is explicitly deleted.
         */
        XLTableColumn& operator=(const XLTableColumn& other);

        /**
         * @brief Move assignment operator.
         * @note The move assignment operator has been explicitly deleted.
         */
        XLTableColumn& operator=(XLTableColumn&& other) noexcept;


        /**
         * @brief 
         * @return the column name
         */
        std::string name() const;

        /**
         * @brief 
         * @param name set the column name
         */
        void setName(const std::string& name) const;

        /**
         * @brief the getter setter function
         * @return return a XLTableColumnProxy ref which could be implicitely convert to string
         */
        XLTableColumnProxy& totalsRowFormula();

        /**
         * @brief the getter setter function
         * @return return a XLTableColumnProxy ref which could be implicitely convert to string
         */
        const XLTableColumnProxy& totalsRowFormula() const;

         /**
         * @brief clear the total row function of this columns and trigger the sheet update
         */
        void clearTotalsRowFormula();

        /**
         * @brief the getter setter function
         * @return return a XLTableColumnProxy ref which could be implicitely convert to string
         */
        XLTableColumnFormulaProxy& columnFormula();

        /**
         * @brief the getter setter function
         * @return return a XLTableColumnFormulaProxy ref which could be implicitely convert to string
         */
        const XLTableColumnFormulaProxy& columnFormula() const;

         /**
         * @brief clear the total row function of this columns and trigger the sheet update
         */
        void clearColumnFormula();
        
        /**
         * @brief get the body range of the cell
         * @return an XLCellReference of the body range of the column
         */
        XLCellRange bodyRange() const;

    private:
        std::shared_ptr<XMLNode>    m_dataNode;
        const XLTable&              m_table;
        XLTableColumnTotalProxy     m_proxyTotal;
        XLTableColumnFormulaProxy   m_proxyColumn;

    };

    


}    // namespace OpenXLSX

//#pragma warning(pop)
#endif    // OPENXLSX_XLTABLECOLUMN_HPP

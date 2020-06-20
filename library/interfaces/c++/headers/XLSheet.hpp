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

#ifndef OPENXLSX_XLSHEET_HPP
#define OPENXLSX_XLSHEET_HPP

#include <string>

#include "openxlsx_export.h"
#include "XLDefinitions.hpp"

namespace OpenXLSX
{
    namespace Impl
    {
        class XLSheet;
    } // namespace Impl

    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLSheet
    {
    public:

        /**
         * @brief
         * @param sheet
         */
        explicit XLSheet(Impl::XLSheet& sheet);

        /**
          * @brief The copy constructor.
          * @param other The object to be copied.
          * @note The default copy constructor is used, i.e. only shallow copying of pointer data members.
          * @todo Can this method be deleted?
          */
        XLSheet(const XLSheet& other) = default;

        /**
         * @brief
         * @param other
         */
        XLSheet(XLSheet&& other) = default;

        /**
         * @brief The destructor
         * @note The default destructor is used, since cleanup of pointer data members is not required.
         */
        virtual ~XLSheet() = default;

        /**
         * @brief Assignment operator
         * @return A reference to the new object.
         * @note The default assignment operator is used, i.e. only shallow copying of pointer data members.
         * @todo Can this method be deleted?
         */
        XLSheet& operator=(const XLSheet& other) = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLSheet& operator=(XLSheet&& other) = default;

        /**
         * @brief Method to retrieve the name of the sheet.
         * @return A std::string with the sheet name.
         */
        virtual std::string const Name() const;

        /**
         * @brief Method for renaming the sheet.
         * @param name A std::string with the new name.
         */
        virtual void SetName(const std::string& name);

        /**
         * @brief Method for getting the current visibility state of the sheet.
         * @return An XLSheetState enum object, with the current sheet state.
         */
        virtual const XLSheetState& State() const;

        /**
         * @brief Method for setting the state of the sheet.
         * @param state An XLSheetState enum object with the new state.
         */
        virtual void SetState(XLSheetState state);

        /**
         * @brief Method to get the type of the sheet.
         * @return An XLSheetType enum object with the sheet type.
         */
        virtual const XLSheetType& Type() const;

        //**
        // * @brief Method for cloning the sheet.
        // * @param newName A std::string with the name of the clone
        // * @return A pointer to the cloned object.
        // * @note This is a pure abstract method. I.e. it is implemented in subclasses.
        // */
        //        virtual XLSheet* Clone(const std::string& newName) = 0;

        /**
         * @brief Method for getting the index of the sheet.
         * @return An int with the index of the sheet.
         */
        virtual unsigned int Index() const;

        /**
         * @brief Method for setting the index of the sheet. This effectively moves the sheet to a different position.
         */
        virtual void SetIndex(int index);

    protected:
        Impl::XLSheet* m_sheet; /**< */
    };
}  // namespace OpenXLSX

#endif //OPENXLSX_XLSHEET_HPP

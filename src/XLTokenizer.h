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

#ifndef OPENXLSX_XLTOKENIZER_H
#define OPENXLSX_XLTOKENIZER_H

#include <string>
#include <vector>

namespace OpenXLSX
{

    const std::string DEFAULT_DELIMITER = " \t\v\n\r\f";

    /**
     * @brief The XLTokenizer class is a convenience class for tokenizing a string, i.e. splitting it up in tokens,
     * separated by one or more defined separators.
     * It is based on the Tokenizer class developed by Song Ho Ahn (http://www.songho.ca/misc/tokenizer/tokenizer.html)
     */
    class XLTokenizer
    {
    public:

        /**
         * @brief
         */
        XLTokenizer();

        /**
         * @brief
         * @param str
         * @param delimiter
         */
        explicit XLTokenizer(const std::string& str, const std::string& delimiter=DEFAULT_DELIMITER);

        /**
         * @brief
         */
        ~XLTokenizer() = default;

        /**
         * @brief
         * @param str
         * @param delimiter
         */
        void set(const std::string& str, const std::string& delimiter=DEFAULT_DELIMITER);

        /**
         * @brief
         * @param str
         */
        void setString(const std::string& str);

        /**
         * @brief
         * @param delimiter
         */
        void setDelimiter(const std::string& delimiter);

        /**
         * @brief
         * @return
         */
        std::string next();

        /**
         * @brief
         * @return
         */
        std::vector<std::string> split();

    private:

        /**
         * @brief
         */
        void skipDelimiter();

        /**
         * @brief
         * @param c
         * @return
         */
        bool isDelimiter(char c);

        std::string buffer; /**< input string */
        std::string token; /**< output string */
        std::string delimiter; /**< delimiter string */
        std::string::const_iterator currPos; /**< string iterator pointing the current position */
    };
}

#endif //OPENXLSX_XLTOKENIZER_H

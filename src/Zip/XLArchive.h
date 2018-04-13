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

#ifndef OPENXL_XLZIPFILE_H
#define OPENXL_XLZIPFILE_H

#include <string>
#include <memory>

#include "libzip++.h"

namespace OpenXLSX
{

//======================================================================================================================
//========== XLArchive Class ===========================================================================================
//======================================================================================================================

    /**
     * @brief A wrapper around the libzippp library, exposing only the relevant functions.
     */
    class XLArchive final
    {

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief Creates and opens an XLArchive object with the given filename.
         * @param zipFileName The name of the file to open.
         * @todo ensure that the file is created, if it doesn't exist.
         */
        explicit XLArchive(const std::string &zipFileName);

        /**
         * @brief Copy constructor (deleted)
         */
        XLArchive(const XLArchive &) = delete;

        /**
         * @brief Closes the file and destroys the object
         */
        ~XLArchive();

        /**
         * @brief Copy assignment operator (deleted)
         * @return A reference to the new object.
         */
        XLArchive &operator=(const XLArchive &) = delete;

        /**
         * @brief Checks if a file with the given name exists in the archive.
         * @param fileName the name of the file to look for
         * @return true if the file exists; otherwise false
         */
        bool FileExists(const std::string &fileName) const;

        /**
         * @brief Get the file with the given name.
         * @param fileName The name of the file to get.
         * @return A std::string with the contents of the file.
         */
        std::string GetFile(const std::string &fileName) const;

        /**
         * @brief Add a file to the archive. If the file already exists, it will be overwritten.
         * @param fileName The name of the new file.
         * @param text The contents (as text) of the new file.
         */
        void AddFile(const std::string &fileName,
                     const std::string &text);

        /**
         * @brief Delete the file with the given name, from the archive.
         * @param fileName The name of the file to delete.
         */
        void DeleteFile(const std::string &fileName);

        /**
         * @brief Save a copy of the archive with a new name and opens the copy.
         * All subsequent file modifications will be applied to the new copy.
         * @param copyName The name of the copy.
         */
        void CopyArchive(const std::string &copyName);

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Variables
//----------------------------------------------------------------------------------------------------------------------

    private:
        std::unique_ptr<libzippp::ZipArchive> m_archive; /**<A pointer to the underlying libzippp archive.*/
        std::string m_fileName; /**<The name of the current archive*/

    };

}

#endif //OPENXL_XLZIPFILE_H

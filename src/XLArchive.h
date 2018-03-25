//
// Created by Troldal on 19/09/16.
//

#ifndef OPENXL_XLZIPFILE_H
#define OPENXL_XLZIPFILE_H

#include <string>
#include <memory>

#include "libzip++.h"

namespace RapidXLSX
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
         * All subsequent file modificantions will be applied to the new copy.
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

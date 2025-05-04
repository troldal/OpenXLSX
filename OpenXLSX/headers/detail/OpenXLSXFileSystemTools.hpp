#ifndef OPENXLSX_TOOLS_H
#define OPENXLSX_TOOLS_H

#include <random>       // std::random_device, std::mt19937, std::uniform_int_distribution
#include <string>       // std::string
#include <sys/stat.h>   // for stat, to test if a file exists and whether a file is a directory

// don't use "stat" directly because windows has compatibility-breaking defines
#if defined(_WIN32)    // moved below includes to make it absolutely clear that this is module-local
#    define STAT _stat                 // _stat should be available in standard environment on Windows
#    define STATSTRUCT struct _stat    // struct _stat also exists - split the two names in case the struct _stat must not be used on windows
#else
#    define STAT stat
#    define STATSTRUCT struct stat
#endif

namespace OpenXLSX
{
    /* 2025-05-04: moved a set of file system related utility functions to this new header file so that Unicode support on windows (nowide)
     * can be implemented in a central location */

    /**
     * @brief Test if path exists as either a file or a directory
     * @param path Check for existence of this
     * @return true if path exists as a file or directory
     */
    inline bool pathExists(const std::string& path)
    {
        STATSTRUCT info;
        if (STAT(path.c_str(), &info ) == 0)    // test if path exists
            return true;
        return false;
    }
#ifdef __GNUC__    // conditionally enable GCC specific pragmas to suppress unused function warning
#   pragma GCC diagnostic push
#   pragma GCC diagnostic ignored "-Wunused-function"
#endif // __GNUC__
    /**
     * @brief Test if fileName exists and is not a directory
     * @param fileName The path to check for existence (as a file)
     * @return true if fileName exists and is a file, otherwise false
     */
    inline bool fileExists(const std::string& fileName)
    {
        STATSTRUCT info;
        if (STAT(fileName.c_str(), &info ) == 0)    // test if path exists
            if ((info.st_mode & S_IFDIR) == 0)          // test if it is NOT a directory
                return true;
        return false;
    }
    inline bool isDirectory(const std::string& fileName)
    {
        STATSTRUCT info;
        if (STAT(fileName.c_str(), &info ) == 0)    // test if path exists
            if ((info.st_mode & S_IFDIR) != 0)          // test if it is a directory
                return true;
        return false;
    }
#ifdef __GNUC__    // conditionally enable GCC specific pragmas to suppress unused function warning
#   pragma GCC diagnostic pop
#endif // __GNUC__

    /**
     * @brief Generates a random filename, which is used to generate a temporary archive when modifying and saving
     * archive files.
     * @param length The length of the filename to create.
     * @return Returns the generated filename, appended with '.tmp'.
     */
    inline std::string GenerateRandomName(int length)
    {
        std::string letters = "abcdefghijklmnopqrstuvwxyz0123456789";

        std::random_device                 rand_dev;
        std::mt19937                       generator(rand_dev());
        std::uniform_int_distribution<int> distr(0, letters.size() - 1);

        std::string result;
        for (int i = 0; i < length; ++i) {
            result += letters[distr(generator)];
        }

        return result + ".tmp";
    }

    /**
     * @brief Generates a random filename with a call to GenerateRandomName, prepending the same path in which filename is located
     * @param filename determine "same path" from this file
     * @param length The length of the filename to create.
     * @return Returns the generated random filename in the same path as filename, appended with '.tmp'.
     * @note this function accounts for amigaos and windows specific drive / directory separators
     */
    inline std::string GenerateRandomNameInSamePath(std::string filename, int length)
    {
        if (filename.empty()) filename = "./"; // local folder

#       ifdef _WIN32
           std::replace( filename.begin(), filename.end(), '\\', '/' ); // pull request #210, alternate fix: fopen etc work fine with forward slashes
#       endif

        // ===== Determine path of the current file
        size_t pathPos = filename.rfind('/');

        // pull request #191, support AmigaOS style paths
#       ifdef __amigaos__
            constexpr const char * localFolder = "";    // local folder on AmigaOS can not be explicitly expressed in a path
            if (pathPos == std::string::npos) pathPos = filename.rfind(':'); // if no '/' found, attempt to find amiga drive root path
#       else
            constexpr const char * localFolder = "./"; // local folder on _WIN32 && __linux__ is .
#       endif
        std::string tempPath{};
        if (pathPos != std::string::npos) tempPath = filename.substr(0, pathPos + 1);
        else tempPath = localFolder; // prepend explicit identification of local folder in case path did not contain a folder

        // ===== Generate a random file name with the same path as the current file
        return tempPath + GenerateRandomName(length);
    }
}    // namespace OpenXLSX

#endif // OPENXLSX_TOOLS_H

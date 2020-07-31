//
// Created by Kenneth Balslev on 31/07/2020.
//

#ifndef OPENXLSX_XLZIPARCHIVE_HPP
#define OPENXLSX_XLZIPARCHIVE_HPP

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"

namespace Zippy
{
    class ZipArchive;
}    // namespace Zippy

namespace OpenXLSX
{
    class OPENXLSX_EXPORT XLZipArchive
    {
    public:
        XLZipArchive();

        ~XLZipArchive();

        bool isOpen();

        void open(const std::string& fileName);

        void close();

        void save(const std::string& path = "");

        void addEntry(const std::string& name, const std::string& data);

        void deleteEntry(const std::string& entryName);

        std::string getEntry(const std::string& name);

        bool hasEntry(const std::string& entryName);

    private:
        std::unique_ptr<Zippy::ZipArchive> m_archive;
    };
}    // namespace OpenXLSX
#endif    // OPENXLSX_XLZIPARCHIVE_HPP

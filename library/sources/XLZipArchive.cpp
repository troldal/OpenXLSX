//
// Created by Kenneth Balslev on 31/07/2020.
//

#include <zippy.hpp>

#include "XLZipArchive.hpp"

using namespace OpenXLSX;

OpenXLSX::XLZipArchive::XLZipArchive() : m_archive(std::make_unique<Zippy::ZipArchive>()) {}

XLZipArchive::~XLZipArchive() = default;

bool OpenXLSX::XLZipArchive::isOpen()
{
    return m_archive->IsOpen();
}

void OpenXLSX::XLZipArchive::open(const std::string& fileName)
{
    m_archive->Open(fileName);
}

void OpenXLSX::XLZipArchive::close()
{
    m_archive->Close();
}

void OpenXLSX::XLZipArchive::save(const std::string& path)
{
    m_archive->Save(path);
}

void OpenXLSX::XLZipArchive::addEntry(const std::string& name, const std::string& data)
{
    m_archive->AddEntry(name, data);
}

void OpenXLSX::XLZipArchive::deleteEntry(const std::string& entryName)
{
    m_archive->DeleteEntry(entryName);
}

std::string OpenXLSX::XLZipArchive::getEntry(const std::string& name)
{
    return m_archive->GetEntry(name).GetDataAsString();
}

bool OpenXLSX::XLZipArchive::hasEntry(const std::string& entryName)
{
    return m_archive->HasEntry(entryName);
}

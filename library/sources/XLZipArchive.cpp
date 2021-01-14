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

// ===== External Includes ===== //
#include <zippy.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLZipArchive.hpp"

using namespace OpenXLSX;

/**
 * @details
 */
OpenXLSX::XLZipArchive::XLZipArchive() : m_archive(nullptr) {}

/**
 * @details
 */
XLZipArchive::~XLZipArchive() = default;

/**
 * @details
 */
XLZipArchive::operator bool() const
{
    return m_archive != nullptr;
}

/**
 * @details
 */
bool OpenXLSX::XLZipArchive::isOpen()
{
    return m_archive && m_archive->IsOpen();
}

/**
 * @details
 */
void OpenXLSX::XLZipArchive::open(const std::string& fileName)
{
    m_archive = std::make_shared<Zippy::ZipArchive>();
    m_archive->Open(fileName);
}

/**
 * @details
 */
void OpenXLSX::XLZipArchive::close()
{
    m_archive->Close();
    m_archive = nullptr;
}

/**
 * @details
 */
void OpenXLSX::XLZipArchive::save(const std::string& path)
{
    m_archive->Save(path);
}

/**
 * @details
 */
void OpenXLSX::XLZipArchive::addEntry(const std::string& name, const std::string& data)
{
    m_archive->AddEntry(name, data);
}

/**
 * @details
 */
void OpenXLSX::XLZipArchive::deleteEntry(const std::string& entryName)
{
    m_archive->DeleteEntry(entryName);
}

/**
 * @details
 */
std::string OpenXLSX::XLZipArchive::getEntry(const std::string& name)
{
    return m_archive->GetEntry(name).GetDataAsString();
}

/**
 * @details
 */
bool OpenXLSX::XLZipArchive::hasEntry(const std::string& entryName)
{
    return m_archive->HasEntry(entryName);
}

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
#ifdef USE_LIBZIP
	#include "detail/LibZip.hpp"
#else
	#include "detail/Zippy.hpp"
#endif

// ===== OpenXLSX Includes ===== //
#include "XLZipArchive.hpp"
#include "XLException.hpp"

using namespace OpenXLSX;

/**
 * @details
 */
XLZipArchive::XLZipArchive() : m_archive(nullptr) {}

/**
 * @details CAUTION: shallow copy (explicit default constructor)
 */
XLZipArchive::XLZipArchive(const XLZipArchive& other)
 : m_archive(other.m_archive)
{}

/**
 * @details
 */
XLZipArchive::XLZipArchive(XLZipArchive&& other) noexcept
 : m_archive(std::move(other.m_archive))
{}

/**
 * @details
 */
XLZipArchive::~XLZipArchive() = default;

// /**
//  * @details CAUTION: shallow copy (explicit default copy assignment)
//  */
// XLZipArchive& XLZipArchive::operator=(const XLZipArchive& other)
// {
//     m_archive = other.m_archive;
//     return *this;
// }

/**
 * @details
 */
XLZipArchive& XLZipArchive::operator=(XLZipArchive&& other) noexcept
{
    m_archive = std::move(other.m_archive);
	 return *this;
}

/**
 * @details
 */
bool XLZipArchive::usesZippy()
{
#ifdef USE_LIBZIP
    return false;
#else
    return true;
#endif
}

/**
 * @details
 */
XLZipArchive::operator bool() const
{
    return isValid();
}

bool XLZipArchive::isValid() const { return m_archive != nullptr; }

/**
 * @details
 */
bool XLZipArchive::isOpen() const
{
    return m_archive && m_archive->IsOpen();
}

/**
 * @details
 */
void XLZipArchive::open(const std::string& fileName)
{
    if (isOpen()) throw XLInputError("XLZipArchive::open: archive is already open"); // prevent double open
    m_archive = std::make_shared<XLZipImplementation>();
    try {
        m_archive->Open(fileName);
    }
    catch( ... ) {    // catch all exceptions
        m_archive.reset();    // make m_archive invalid again
        throw;                // re-throw
    }
}

/**
 * @details
 */
void XLZipArchive::close()
{
    if (!m_archive) throw XLInputError("XLZipArchive::close: archive is not open"); // prevent SEGFAULT
    m_archive->Close();
    m_archive.reset();
}

/**
 * @details
 */
void XLZipArchive::commitChanges()
{
    if (!m_archive) throw XLInputError("XLZipArchive::commitChanges: archive is not open"); // prevent SEGFAULT
#ifdef USE_LIBZIP
    m_archive->CommitChanges(); // libzip needs a zip archive close & reopen
#else
    // no-op for zippy
#endif
}

/**
 * @details
 */
void XLZipArchive::save(const std::string& path) // NOLINT
{
    if (!m_archive) throw XLInputError("XLZipArchive::save: archive is not open"); // prevent SEGFAULT
    m_archive->Save(path);
}

/**
 * @details
 */
void XLZipArchive::addEntry(const std::string& name, const std::string& data) // NOLINT
{
    if (!m_archive) throw XLInputError("XLZipArchive::addEntry: archive is not open"); // prevent SEGFAULT
    m_archive->AddEntry(name, data);
}

/**
 * @details
 */
void XLZipArchive::addEntryAndCommit(const std::string& name, const std::string& data) // NOLINT
{
    addEntry(name, data);
    commitChanges();       // ensure that entry name is directly available for calls to getEntry
}

/**
 * @details
 */
void XLZipArchive::deleteEntry(const std::string& entryName) // NOLINT
{
    if (!m_archive) throw XLInputError("XLZipArchive::deleteEntry: archive is not open"); // prevent SEGFAULT
    m_archive->DeleteEntry(entryName);
}

/**
 * @details
 */
std::string XLZipArchive::getEntry(const std::string& name) const {
    if (!m_archive) throw XLInputError("XLZipArchive::getEntry: archive is not open"); // prevent SEGFAULT
#ifdef USE_LIBZIP
    return m_archive->GetEntryDataAsString(name);
#else
    return m_archive->GetEntry(name).GetDataAsString();
#endif
}

/**
 * @details
 */
bool XLZipArchive::hasEntry(const std::string& entryName) const {
    if (!m_archive) throw XLInputError("XLZipArchive::hasEntry: archive is not open"); // prevent SEGFAULT
    return m_archive->HasEntry(entryName);
}

/**
 * @details
 */
ssize_t XLZipArchive::entryCount() const
{
#ifdef USE_LIBZIP
    return m_archive->EntryCount();
#else
    return 0; // TODO: can Zippy support this?
#endif
}

/**
 * @details
 */
std::string XLZipArchive::entryName( int index ) const
{
#ifdef USE_LIBZIP
    return m_archive->EntryName(index);
#else
    return ""; // TODO: can Zippy support this?
#endif
}

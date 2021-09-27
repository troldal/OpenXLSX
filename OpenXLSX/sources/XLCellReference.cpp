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
#include <array>
#include <cmath>
#ifdef CHARCONV_ENABLED
#    include <charconv>
#endif

// ===== OpenXLSX Includes ===== //
#include "XLCellReference.hpp"
#include "XLConstants.hpp"
#include "XLException.hpp"

using namespace OpenXLSX;

constexpr uint8_t alphabetSize = 26;

constexpr uint8_t asciiOffset = 64;

namespace {
    bool addressIsValid(uint32_t row, uint16_t column) {
        return !(row < 1 || row > OpenXLSX::MAX_ROWS || column < 1 || column > OpenXLSX::MAX_COLS);
    }
} // namespace

/**
 * @details The constructor creates a new XLCellReference from a string, e.g. 'A1'. If there's no input,
 * the default reference will be cell A1.
 */
XLCellReference::XLCellReference(const std::string& cellAddress)
{
    if (!cellAddress.empty()) setAddress(cellAddress);
    if (!addressIsValid(m_row, m_column)) {
        m_row = 1;
        m_column = 1;
        m_cellAddress = "A1";
        throw XLCellAddressError("Cell reference is invalid");
    }
}

/**
 * @details This constructor creates a new XLCellReference from a given row and column number, e.g. 1,1 (=A1)
 * @todo consider swapping the arguments.
 */
XLCellReference::XLCellReference(uint32_t row, uint16_t column){
    if (!addressIsValid(row, column)) throw XLCellAddressError("Cell reference is invalid");
    setRowAndColumn(row, column);
}

/**
 * @details This constructor creates a new XLCellReference from a row number and the column name (e.g. 1, A)
 * @todo consider swapping the arguments.
 */
XLCellReference::XLCellReference(uint32_t row, const std::string& column) {
    if (!addressIsValid(row, columnAsNumber(column))) throw XLCellAddressError("Cell reference is invalid");
    setRowAndColumn(row, columnAsNumber(column));
}

/**
 * @details
 */
XLCellReference::XLCellReference(const XLCellReference& other) = default;

/**
 * @details
 */
XLCellReference::XLCellReference(XLCellReference&& other) noexcept = default;

/**
 * @details
 */
XLCellReference::~XLCellReference() = default;

/**
 * @details
 */
XLCellReference& XLCellReference::operator=(const XLCellReference& other) = default;

/**
 * @details
 */
XLCellReference& XLCellReference::operator=(XLCellReference&& other) noexcept = default;

/**
 * @details
 */
 XLCellReference& XLCellReference::operator++()
{
    if (m_column < MAX_COLS) {
        setColumn(m_column + 1);
    }
    else if (m_column == MAX_COLS && m_row < MAX_ROWS) {
        m_column = 1;
        setRow(m_row + 1);
    }
    else if (m_column == MAX_COLS && m_row == MAX_ROWS) {
        m_column = 1;
        m_row = 1;
        m_cellAddress = "A1";
    }

    return *this;
}

/**
 * @details
 */
 XLCellReference XLCellReference::operator++(int) { // NOLINT
    auto oldRef(*this);
    ++(*this);
    return oldRef;
}

/**
 * @details
 */
 XLCellReference& XLCellReference::operator--()
{
    if (m_column > 1) {
        setColumn(m_column - 1);
    }
    else if (m_column == 1 && m_row > 1) {
        m_column = MAX_COLS;
        setRow(m_row - 1);
    }
    else if (m_column == 1 && m_row == 1) {
        m_column = MAX_COLS;
        m_row = MAX_ROWS;
        m_cellAddress = "XFD1048576";
    }
    return *this;
}

/**
 * @details
 */
 XLCellReference XLCellReference::operator--(int) {// NOLINT
    auto oldRef(*this);
    --(*this);
    return oldRef;
}

/**
 * @details Returns the m_row property.
 */
uint32_t XLCellReference::row() const
{
    return m_row;
}

/**
 * @details Sets the row of the XLCellReference objects. If the number is larger than 16384 (the maximum),
 * the row is set to 16384.
 */
void XLCellReference::setRow(uint32_t row)
{
    if(!addressIsValid(row, m_column)) throw XLCellAddressError("Cell reference is invalid");

    m_row = row;
    m_cellAddress = columnAsString(m_column) + rowAsString(m_row);
}

/**
 * @details Returns the m_column property.
 */
uint16_t XLCellReference::column() const
{
    return m_column;
}

/**
 * @details Sets the column of the XLCellReference object. If the number is larger than 1048576 (the maximum),
 * the column is set to 1048576.
 */
void XLCellReference::setColumn(uint16_t column)
{
    if(!addressIsValid(m_row, column)) throw XLCellAddressError("Cell reference is invalid");

    m_column = column;
    m_cellAddress = columnAsString(m_column) + rowAsString(m_row);
}

/**
 * @details Sets row and column of the XLCellReference object. Checks that row and column is less than
 * or equal to the maximum row and column numbers allowed by Excel.
 */
void XLCellReference::setRowAndColumn(uint32_t row, uint16_t column)
{
    if (!addressIsValid(row, column)) throw XLCellAddressError("Cell reference is invalid");

    m_row = row;
    m_column = column;
    m_cellAddress = columnAsString(m_column) + rowAsString(m_row);
}

/**
 * @details Returns the m_cellAddress property.
 */
std::string XLCellReference::address() const
{
    return m_cellAddress;
}

/**
 * @details Sets the address of the XLCellReference object, e.g. 'B2'. Checks that row and column is less than
 * or equal to the maximum row and column numbers allowed by Excel.
 */
void XLCellReference::setAddress(const std::string& address)
{
    auto coordinates = coordinatesFromAddress(address);
    m_row         = coordinates.first;
    m_column      = coordinates.second;
    m_cellAddress = address;
}

/**
 * @details
 */
std::string XLCellReference::rowAsString(uint32_t row)
{
#ifdef CHARCONV_ENABLED
    std::array<char, 7> str {};  // NOLINT
    auto*               p = std::to_chars(str.data(), str.data() + str.size(), row).ptr;
    return std::string{str.data(), static_cast<uint16_t>(p - str.data())};
#else
    std::string result;
    while (row != 0) {
        int rem = row % 10;
        result += (rem > 9) ? (rem - 10) + 'a' : rem + '0';
        row = row / 10;
    }

    for (int i = 0; i < result.length() / 2; i++) std::swap(result[i], result[result.length() - i - 1]);

    return result;
#endif
}

/**
 * @details
 */
uint32_t XLCellReference::rowAsNumber(const std::string& row)
{
#ifdef CHARCONV_ENABLED
    uint32_t value = 0;
    std::from_chars(row.data(), row.data() + row.size(), value); // NOLINT
    return value;
#else
    return stoul(row);
#endif
}

/**
 * @details Helper method to calculate the column letter from column number.
 */
std::string XLCellReference::columnAsString(uint16_t column)
{
    std::string result;

    // ===== If there is one letter in the Column Name:
    if (column <= alphabetSize) result += char(column + asciiOffset);

    // ===== If there are two letters in the Column Name:
    else if (column > alphabetSize && column <= alphabetSize * (alphabetSize + 1)) {
        result += char((column - (alphabetSize + 1)) / alphabetSize + asciiOffset + 1);
        result += char((column - (alphabetSize + 1)) % alphabetSize + asciiOffset + 1);
    }

    // ===== If there is three letters in the Column Name:
    else {
        result += char((column - 703) / (alphabetSize * alphabetSize) + asciiOffset + 1);  // NOLINT
        result += char(((column - 703) / alphabetSize) % alphabetSize + asciiOffset + 1);  // NOLINT
        result += char((column - 703) % alphabetSize + asciiOffset + 1);  // NOLINT
    }

    return result;
}

/**
 * @details Helper method to calculate the column number from column letter.
 */
uint16_t XLCellReference::columnAsNumber(const std::string& column)
{
    uint16_t result = 0;

    for (int16_t i = static_cast<int16_t>(column.size() - 1), j = 0; i >= 0; --i, ++j) { // NOLINT
        result += static_cast<uint16_t>((column[static_cast<uint64_t>(i)] - asciiOffset) * std::pow(alphabetSize, j));
    }

    return result;
}

/**
 * @details Helper method for calculating the coordinates from the cell address.
 * @todo Consider checking if the given address is valid.
 */
XLCoordinates XLCellReference::coordinatesFromAddress(const std::string& address)
{
    uint64_t letterCount = 0;
    for (auto letter : address) {
        if (letter >= 65)  // NOLINT
            ++letterCount;
        else if (letter <= 57)  // NOLINT
            break;
    }

    auto numberCount = address.size() - letterCount;

    return std::make_pair(rowAsNumber(address.substr(letterCount, numberCount)), columnAsNumber(address.substr(0, letterCount)));
}
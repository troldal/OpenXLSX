//
// Created by Troldal on 07/09/16.
//

#include "XLCellReference_Impl.hpp"

#include <XLDefinitions.hpp>
#include <iostream>
#include <string>
#include <array>

#ifdef CHARCONV_ENABLED
#include <charconv>
#include <XLException.hpp>

#endif

using namespace std;
using namespace OpenXLSX;

/**
 * @details The constructor creates a new XLCellReference from a string, e.g. 'A1'. If there's no input,
 * the default reference will be cell A1.
 */
Impl::XLCellReference::XLCellReference(const std::string& cellAddress)
        : m_row(1),
          m_column(1),
          m_cellAddress("A1") {

    if (!cellAddress.empty())
        SetAddress(cellAddress);
}

/**
 * @details This constructor creates a new XLCellReference from a given row and column number, e.g. 1,1 (=A1)
 * @todo consider swapping the arguments.
 */
Impl::XLCellReference::XLCellReference(uint32_t row, uint16_t column)
        : m_row(row),
          m_column(column),
          m_cellAddress(ColumnAsString(column) + RowAsString(row)) {

    if (m_row < 1 || m_row > maxRows || m_column < 1 || m_column > maxCols) {
        m_row = 0;
        m_column = 0;
        m_cellAddress = "";
    }
}

/**
 * @details This constructor creates a new XLCellReference from a row number and the column name (e.g. 1, A)
 * @todo consider swapping the arguments.
 */
Impl::XLCellReference::XLCellReference(uint32_t row, const std::string& column)
        : m_row(row),
          m_column(ColumnAsNumber(column)),
          m_cellAddress(column + RowAsString(row)) {

    if (m_row < 1 || m_row > maxRows || m_column < 1 || m_column > maxCols) {
        m_row = 0;
        m_column = 0;
        m_cellAddress = "";
    }
}

/**
 * @details Returns the m_row property.
 */
uint32_t Impl::XLCellReference::Row() const {

    return m_row;
}

/**
 * @details Sets the row of the XLCellReference objects. If the number is larger than 16384 (the maximum),
 * the row is set to 16384.
 */
void Impl::XLCellReference::SetRow(uint32_t row) {

    if (row < 1)
        m_row = 1;
    else if (row > maxRows)
        m_row = maxRows;
    else
        m_row = row;

    m_cellAddress = ColumnAsString(m_column) + RowAsString(m_row);
}

/**
 * @details Returns the m_column property.
 */
uint16_t Impl::XLCellReference::Column() const {

    return m_column;
}

/**
 * @details Sets the column of the XLCellReference object. If the number is larger than 1048576 (the maximum),
 * the column is set to 1048576.
 */
void Impl::XLCellReference::SetColumn(uint16_t column) {

    if (column < 1)
        m_column = 1;
    else if (column > maxCols)
        m_column = maxCols;
    else
        m_column = column;

    m_cellAddress = ColumnAsString(m_column) + RowAsString(m_row);
}

/**
 * @details Sets row and column of the XLCellReference object. Checks that row and column is less than
 * or equal to the maximum row and column numbers allowed by Excel.
 */
void Impl::XLCellReference::SetRowAndColumn(uint32_t row, uint16_t column) {

    if (row < 1)
        m_row = 1;
    else if (row > maxRows)
        m_row = maxRows;
    else
        m_row = row;

    if (column < 1)
        m_column = 1;
    else if (column > maxCols)
        m_column = maxCols;
    else
        m_column = column;

    m_cellAddress = ColumnAsString(m_column) + RowAsString(m_row);
}

/**
 * @details Returns the m_cellAddress property.
 */
std::string Impl::XLCellReference::Address() const {

    return m_cellAddress;
}

/**
 * @details Sets the address of the XLCellReference object, e.g. 'B2'. Checks that row and column is less than
 * or equal to the maximum row and column numbers allowed by Excel.
 */
void Impl::XLCellReference::SetAddress(const std::string& address) {

    auto coordinates = CoordinatesFromAddress(address);
    if (coordinates.first < 1 || coordinates.first > maxRows || coordinates.second < 1
            || coordinates.second > maxCols) {
        m_row = 0;
        m_column = 0;
        m_cellAddress = "";
    }
    else {
        m_row = coordinates.first;
        m_column = coordinates.second;
        m_cellAddress = address;
    }
}

/**
 * @details
 *
 * @todo Find out why std::to_chars causes a linker error when using LLVM/Clang. As a workaround, a custom conversion
 * algorithm has been implemented for clang (std::to_string is too slow!)
 */
std::string Impl::XLCellReference::RowAsString(uint32_t row) {

#ifdef CHARCONV_ENABLED
    std::array<char, 7> str {};
    auto *p = std::to_chars(str.data(), str.data() + str.size(), row).ptr;
    return string(str.data(), p - str.data());
#else
    string result;
    while (row != 0) {
        int rem = row % 10;
        result += (rem > 9) ? (rem - 10) + 'a' : rem + '0';
        row = row / 10;
    }

    for (int i = 0; i < result.length() / 2; i++)
        std::swap(result[i], result[result.length() - i - 1]);

    return result;
#endif

}

/**
 * @details
 */
uint32_t Impl::XLCellReference::RowAsNumber(const std::string& row) {

#ifdef CHARCONV_ENABLED
    unsigned long value = 0;
    std::from_chars(row.data(), row.data() + row.size(), value);
    return value;
#else
    return stoul(row);
#endif
}

/**
 * @details Helper method to calculate the column letter from column number.
 */
std::string Impl::XLCellReference::ColumnAsString(uint16_t column) {

    string result;

    if (column <= 1)
        return std::string("A");
    else if (column >= maxCols)
        result += std::string("XFD");

        // ===== If there is one letter in the Column Name:
    else if (column <= alphabetSize)
        result += char(column + asciiOffset);

        // ===== If there are two letters in the Column Name:
    else if (column > alphabetSize && column <= 702) {
        result += char((column - (alphabetSize + 1)) / alphabetSize + asciiOffset + 1);
        result += char((column - (alphabetSize + 1)) % alphabetSize + asciiOffset + 1);
    }

        // ===== If there is three letters in the Column Name:
    else {
        result += char((column - 703) / (alphabetSize * alphabetSize) + asciiOffset + 1);
        result += char(((column - 703) / alphabetSize) % alphabetSize + asciiOffset + 1);
        result += char((column - 703) % alphabetSize + asciiOffset + 1);
    }

    return result;
}

/**
 * @details Helper method to calculate the column number from column letter.
 */
uint16_t Impl::XLCellReference::ColumnAsNumber(const std::string& column) {

    unsigned int length = column.size();
    unsigned int result = 0;

    // ===== If the length of the string is only one character, look up the number in the alphabet.
    if (length == 1) {
        result += (column.at(0) - asciiOffset);
    }

        // ===== If the string is two characters long...
    else if (length == 2) {
        result += (column.at(0) - asciiOffset) * alphabetSize;
        result += (column.at(1) - asciiOffset);
    }

        // ===== If the string is three characters long...
    else if (length == 3) {
        result += (column.at(0) - asciiOffset) * alphabetSize * alphabetSize;
        result += (column.at(1) - asciiOffset) * alphabetSize;
        result += (column.at(2) - asciiOffset);
    }

    return result;
}

/**
 * @details Helper method for calculating the coordinates from the cell address.
 */
Impl::XLCoordinates Impl::XLCellReference::CoordinatesFromAddress(const std::string& address) {

    int letterCount = 0;

    for (auto letter : address) {
        if (isalpha(letter) != 0)
            ++letterCount;
    }

    int numberCount = address.size() - letterCount;

    if (letterCount < 1 || letterCount > 3 || numberCount < 1 || numberCount > 7)
        throw XLException("Cell address \"" + address + "\" is invalid!");

    return make_pair(RowAsNumber(address.substr(letterCount, numberCount)),
            ColumnAsNumber(address.substr(0, letterCount)));
}

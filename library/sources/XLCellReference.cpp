//
// Created by Troldal on 07/09/16.
//

#include "XLCellReference.hpp"

#include "XLDefinitions.hpp"
#include "XLException.hpp"

#include <array>
#include <cmath>

#ifdef CHARCONV_ENABLED
#    include <charconv>
#endif

using namespace std;
using namespace OpenXLSX;

/**
 * @details The constructor creates a new XLCellReference from a string, e.g. 'A1'. If there's no input,
 * the default reference will be cell A1.
 */
XLCellReference::XLCellReference(const std::string& cellAddress) : m_row(1), m_column(1), m_cellAddress("A1")
{
    if (!cellAddress.empty()) SetAddress(cellAddress);
}

/**
 * @details This constructor creates a new XLCellReference from a given row and column number, e.g. 1,1 (=A1)
 * @todo consider swapping the arguments.
 */
XLCellReference::XLCellReference(uint32_t row, uint16_t column)
    : m_row(row),
      m_column(column),
      m_cellAddress(ColumnAsString(column) + RowAsString(row))
{}

/**
 * @details This constructor creates a new XLCellReference from a row number and the column name (e.g. 1, A)
 * @todo consider swapping the arguments.
 */
XLCellReference::XLCellReference(uint32_t row, const std::string& column)
    : m_row(row),
      m_column(ColumnAsNumber(column)),
      m_cellAddress(column + RowAsString(row))
{}

/**
 * @details Returns the m_row property.
 */
uint32_t XLCellReference::Row() const
{
    return m_row;
}

/**
 * @details Sets the row of the XLCellReference objects. If the number is larger than 16384 (the maximum),
 * the row is set to 16384.
 */
void XLCellReference::SetRow(uint32_t row)
{
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
uint16_t XLCellReference::Column() const
{
    return m_column;
}

/**
 * @details Sets the column of the XLCellReference object. If the number is larger than 1048576 (the maximum),
 * the column is set to 1048576.
 */
void XLCellReference::SetColumn(uint16_t column)
{
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
void XLCellReference::SetRowAndColumn(uint32_t row, uint16_t column)
{
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
std::string XLCellReference::Address() const
{
    return m_cellAddress;
}

/**
 * @details Sets the address of the XLCellReference object, e.g. 'B2'. Checks that row and column is less than
 * or equal to the maximum row and column numbers allowed by Excel.
 */
void XLCellReference::SetAddress(const std::string& address)
{
    auto coordinates = CoordinatesFromAddress(address);

    m_row         = coordinates.first;
    m_column      = coordinates.second;
    m_cellAddress = address;
}

/**
 * @details
 *
 * @todo Find out why std::to_chars causes a linker error when using LLVM/Clang. As a workaround, a custom conversion
 * algorithm has been implemented for clang (std::to_string is too slow!)
 */
std::string XLCellReference::RowAsString(uint32_t row)
{
#ifdef CHARCONV_ENABLED
    std::array<char, 7> str {};
    auto*               p = std::to_chars(str.data(), str.data() + str.size(), row).ptr;
    return std::string(str.data(), p - str.data());
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
uint32_t XLCellReference::RowAsNumber(const std::string& row)
{
#ifdef CHARCONV_ENABLED
    uint32_t value = 0;
    std::from_chars(row.data(), row.data() + row.size(), value);
    return value;
#else
    return stoul(row);
#endif
}

/**
 * @details Helper method to calculate the column letter from column number.
 */
std::string XLCellReference::ColumnAsString(uint16_t column)
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
        result += char((column - 703) / (alphabetSize * alphabetSize) + asciiOffset + 1);
        result += char(((column - 703) / alphabetSize) % alphabetSize + asciiOffset + 1);
        result += char((column - 703) % alphabetSize + asciiOffset + 1);
    }

    return result;
}

/**
 * @details Helper method to calculate the column number from column letter.
 */
uint16_t XLCellReference::ColumnAsNumber(const std::string& column)
{
    uint16_t result = 0;

    for (int32_t i = column.size() - 1, j = 0; i >= 0; --i, ++j) {
        result += (column[i] - asciiOffset) * std::pow(alphabetSize, j);
    }

    return result;
}

/**
 * @details Helper method for calculating the coordinates from the cell address.
 */
XLCoordinates XLCellReference::CoordinatesFromAddress(const std::string& address)
{
    int letterCount = 0;
    for (auto letter : address) {
        if (letter >= 65)
            ++letterCount;
        else if (letter <= 57)
            break;
    }

    int numberCount = address.size() - letterCount;

    return std::make_pair(RowAsNumber(address.substr(letterCount, numberCount)), ColumnAsNumber(address.substr(0, letterCount)));
}
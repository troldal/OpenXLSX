//
// Created by Troldal on 07/09/16.
//

#include "XLCellReference.h"

#include <iostream>

using namespace std;
using namespace OpenXLSX::Impl;

const unordered_map<int, string> XLCellReference::s_alphabet = {{0, "A"},
                                                                {1, "B"},
                                                                {2, "C"},
                                                                {3, "D"},
                                                                {4, "E"},
                                                                {5, "F"},
                                                                {6, "G"},
                                                                {7, "H"},
                                                                {8, "I"},
                                                                {9, "J"},
                                                                {10, "K"},
                                                                {11, "L"},
                                                                {12, "M"},
                                                                {13, "N"},
                                                                {14, "O"},
                                                                {15, "P"},
                                                                {16, "Q"},
                                                                {17, "R"},
                                                                {18, "S"},
                                                                {19, "T"},
                                                                {20, "U"},
                                                                {21, "V"},
                                                                {22, "W"},
                                                                {23, "X"},
                                                                {24, "Y"},
                                                                {25, "Z"}};

unordered_map<string, unsigned int> XLCellReference::s_columnNumbers = {};
unordered_map<unsigned int, string> XLCellReference::s_columnNames = {};
unordered_map<unsigned long, std::string> XLCellReference::s_rowNames = {};
unordered_map<std::string, unsigned long> XLCellReference::s_rowNumbers = {};

/**
 * @details The constructor creates a new XLCellReference from a string, e.g. 'A1'. If there's no input,
 * the default reference will be cell A1.
 */
XLCellReference::XLCellReference(const std::string &cellAddress)
    : m_row(1),
      m_column(1),
      m_cellAddress("A1"),
      m_valid(true)
{
    if (cellAddress != "") SetAddress(cellAddress);
}

/**
 * @details This constructor creates a new XLCellReference from a given row and column number, e.g. 1,1 (=A1)
* @todo consider swapping the arguments.
 */
XLCellReference::XLCellReference(unsigned long row,
                                 unsigned int column)
    : m_row(row),
      m_column(column),
      m_cellAddress(ColumnAsString(column) +
          RowAsString(row)),
      m_valid(false)
{
    if (m_row < 1 || m_row > 1048576 || m_column < 1 || m_column > 16384) {
        m_row = 0;
        m_column = 0;
        m_cellAddress = "";
        m_valid = false;
    }
    else {
        m_valid = true;
    }
}

/**
 * @details This constructor creates a new XLCellReference from a row number and the column name (e.g. 1, A)
 * @todo consider swapping the arguments.
 */
XLCellReference::XLCellReference(unsigned long row,
                                 const std::string &column)
    : m_row(row),
      m_column(ColumnAsNumber(column)),
      m_cellAddress(
          column + RowAsString(row)),
      m_valid(false)
{
    if (m_row < 1 || m_row > 1048576 || m_column < 1 || m_column > 16384) {
        m_row = 0;
        m_column = 0;
        m_cellAddress = "";
        m_valid = false;
    }
    else {
        m_valid = true;
    }
}

/**
 * @details Returns the m_row property.
 */
unsigned long XLCellReference::Row() const
{
    return m_row;
}

/**
 * @details Sets the row of the XLCellReference objects. If the number is larger than 16384 (the maximum),
 * the row is set to 16384.
 */
void XLCellReference::SetRow(unsigned long row)
{

    if (row < 1)
        m_row = 1;
    else if (row > 1048576)
        m_row = 1048576;
    else
        m_row = row;

    m_cellAddress = ColumnAsString(m_column) + RowAsString(m_row);
}

/**
 * @details Returns the m_column property.
 */
unsigned int XLCellReference::Column() const
{
    return m_column;
}

/**
 * @details Sets the column of the XLCellReference object. If the number is larger than 1048576 (the maximum),
 * the column is set to 1048576.
 */
void XLCellReference::SetColumn(unsigned int column)
{

    if (column < 1)
        m_column = 1;
    else if (column > 16384)
        m_column = 16384;
    else
        m_column = column;

    m_cellAddress = ColumnAsString(m_column) + RowAsString(m_row);
}

/**
 * @details Sets row and column of the XLCellReference object. Checks that row and column is less than
 * or equal to the maximum row and column numbers allowed by Excel.
 */
void XLCellReference::SetRowAndColumn(unsigned long row,
                                      unsigned int column)
{

    if (row < 1)
        m_row = 1;
    else if (row > 1048576)
        m_row = 1048576;
    else
        m_row = row;

    if (column < 1)
        m_column = 1;
    else if (column > 16384)
        m_column = 16384;
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
void XLCellReference::SetAddress(const std::string &address)
{
    auto coordinates = CoordinatesFromAddress(address);
    if (coordinates.first < 1 || coordinates.first > 1048576 || coordinates.second < 1 || coordinates.second > 16384) {
        m_row = 0;
        m_column = 0;
        m_cellAddress = "";
        m_valid = false;
    }
    else {
        m_row = coordinates.first;
        m_column = coordinates.second;
        m_cellAddress = address;
        m_valid = true;
    }
}

/**
 * @details
 */
std::string XLCellReference::RowAsString(unsigned long row)
{

    auto iter = s_rowNames.find(row);

    if (iter == s_rowNames.end()) {
        s_rowNames[row] = to_string(row);
        return s_rowNames.at(row);
    }
    else {
        return iter->second;
    }
}

/**
 * @details
 */
unsigned long XLCellReference::RowAsNumber(const std::string &row)
{
    auto iter = s_rowNumbers.find(row);

    if (iter == s_rowNumbers.end()) {
        s_rowNumbers[row] = stol(row);
        return s_rowNumbers.at(row);
    }
    else {
        return iter->second;
    }
}

/**
 * @details Helper method to calculate the column letter from column number.
 */
std::string XLCellReference::ColumnAsString(unsigned int column)
{

    // If the Column Name has already been stored, look it up.
    auto iter = s_columnNames.find(column);
    if (iter != s_columnNames.end()) return iter->second;

    // Otherwise calculate it
    string result;

    if (column <= 1) return std::string("A");
    else if (column >= 16384) result = std::string("XFD");

        // If there is one letter in the Column Name:
    else if (column <= 26) result = s_alphabet.at(column - 1);

        // If there are two letters in the Column Name:
    else if (column > 26 && column <= 702) {
        int lastLetter = (column - 27) % 26;
        int firstLetter = (column - 27) / 26;

        result = s_alphabet.at(firstLetter) + s_alphabet.at(lastLetter);
    }

        // If there is three letters in the Column Name:
    else {
        int lastLetter = (column - 703) % 26;
        int middleLetter = ((column - 703) / 26) % 26;
        int firstLetter = (column - 703) / 676;

        result = s_alphabet.at(firstLetter) + s_alphabet.at(middleLetter) + s_alphabet.at(lastLetter);
    }

    // Store the results for later use
    s_columnNames[column] = result;
    s_columnNumbers[result] = column;
    return result;
}

/**
 * @details Helper method to calculate the column number from column letter.
 */
unsigned int XLCellReference::ColumnAsNumber(const std::string &column)
{

    // If the Column number has already been stored, look it up.
    auto iter = s_columnNumbers.find(column);
    if (iter != s_columnNumbers.end()) return iter->second;

    // Otherwise, calculate it
    unsigned int length = column.size();
    unsigned int result = 0;

    // If the length of the string is only one character, look up the number in the alphabet.
    if (length == 1) {
        for (auto letter : s_alphabet) {
            if (letter.second == column) {
                result += letter.first + 1;
                break;
            }
        }
    }

    // If the string is two characters long...
    if (length == 2) {
        for (auto letter : s_alphabet) {
            if (letter.second == column.substr(0, 1)) {
                result += (letter.first + 1) * 26;
                break;
            }
        }

        for (auto letter : s_alphabet) {
            if (letter.second == column.substr(1, 1)) {
                result += (letter.first + 1);
                break;
            }
        }
    }
    // If the string is three characters long...
    if (length == 3) {
        for (auto letter : s_alphabet) {
            if (letter.second == column.substr(0, 1)) {
                result += (letter.first + 1) * 26 * 26;
                break;
            }
        }

        for (auto letter : s_alphabet) {
            if (letter.second == column.substr(1, 1)) {
                result += (letter.first + 1) * 26;
                break;
            }
        }

        for (auto letter : s_alphabet) {
            if (letter.second == column.substr(2, 1)) {
                result += (letter.first + 1);
                break;
            }
        }
    }

    // Store the results for later use
    s_columnNumbers[column] = result;
    s_columnNames[result] = column;
    return result;
}

/**
 * @details Helper method for calculating the coordinates from the cell address.
 */
std::pair<unsigned long, unsigned int> XLCellReference::CoordinatesFromAddress(const std::string &address)
{
    int letterCount = 0;

    for (auto letter : address) {
        if (isalpha(letter)) ++letterCount;
    }

    int numberCount = address.size() - letterCount;

    if (letterCount < 1 || letterCount > 3 || numberCount < 1 || numberCount > 7)
        // If the Address is invalid, return 0,0
        //return pair<unsigned long, unsigned int>(0, 0);
        return make_pair(0, 0);
    else {
        unsigned int column = ColumnAsNumber(address.substr(0, letterCount));
        unsigned long row = RowAsNumber(address.substr(letterCount, numberCount));

        //return pair<unsigned long, unsigned int>(Row, Column);
        return make_pair(row, column);
    }

}

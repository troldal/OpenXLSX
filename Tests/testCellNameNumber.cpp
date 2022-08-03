/**
 * @file testCellNameNumber.cpp
 * @author Oliver Ofenloch (57812959+ofenloch@users.noreply.github.com)
 * @brief
 * @version 0.1
 * @date 2022-08-03
 *
 * @copyright Copyright (c) 2022
 *
 */

#include <OpenXLSX.hpp>
#include <catch.hpp>
#include <fstream>

TEST_CASE("Cell Name & Number Tests", "[CellNameNumberConversion]")
{
    SECTION("Consistency Test with Selected Values")
    {
        std::multimap<uint16_t, std::string> sanctionedValues;
        sanctionedValues.insert(std::pair(1, "A"));
        sanctionedValues.insert(std::pair(1, "a"));
        sanctionedValues.insert(std::pair(2, "B"));
        sanctionedValues.insert(std::pair(26, "Z"));
        sanctionedValues.insert(std::pair(27, "AA"));
        sanctionedValues.insert(std::pair(28, "AB"));
        sanctionedValues.insert(std::pair(52, "AZ"));
        sanctionedValues.insert(std::pair(53, "BA"));
        sanctionedValues.insert(std::pair(129, "DY"));
        sanctionedValues.insert(std::pair(130, "DZ"));
        sanctionedValues.insert(std::pair(131, "EA"));
        sanctionedValues.insert(std::pair(132, "EB"));
        sanctionedValues.insert(std::pair(702, "ZZ"));
        sanctionedValues.insert(std::pair(703, "AAA"));
        sanctionedValues.insert(std::pair(704, "AAB"));
        sanctionedValues.insert(std::pair(728, "AAZ"));
        sanctionedValues.insert(std::pair(729, "ABA"));
        sanctionedValues.insert(std::pair(1023, "AMI"));
        sanctionedValues.insert(std::pair(1023, "ami"));
        sanctionedValues.insert(std::pair(1024, "AMJ"));
        sanctionedValues.insert(std::pair(1024, "amj"));
        for (auto column : sanctionedValues) {
            const uint16_t    columnNumberOrig = column.first;
            const std::string columnNameOrig   = column.second;
            const uint16_t    columnNumberNew  = OpenXLSX::XLWorksheet::columnNameToNumber(columnNameOrig);
            std::string       columnNameNew    = OpenXLSX::XLWorksheet::columnNumberToName(columnNumberOrig);

            std::string columnNameOrigUpper   = column.second;
            // convert columnNameNew to upper case (see
            // https://stackoverflow.com/questions/313970/how-to-convert-an-instance-of-stdstring-to-lower-case)
            std::transform(columnNameOrigUpper.begin(), columnNameOrigUpper.end(), columnNameOrigUpper.begin(), [](unsigned char c) {
                return std::toupper(c);
            });

            REQUIRE(columnNameNew == columnNameOrigUpper);
            REQUIRE(columnNumberNew == columnNumberOrig);
        }
    }

    SECTION("Number To Name Conversion")
    {
        for (auto cellNumber : { 1, 5, 25, 26, 27, 52, 53, 54, 100, 200, 512, 1023, 1024, 1025 }) {
            const std::string cellName    = OpenXLSX::XLWorksheet::columnNumberToName(cellNumber);
            const uint16_t    cellNumber2 = OpenXLSX::XLWorksheet::columnNameToNumber(cellName);
            REQUIRE(cellNumber2 == cellNumber);
        }
    }

    SECTION("Name To Number Conversion")
    {
        for (auto cellName : { "A", "E", "Z", "AA", "AB", "AZ", "BA", "BZ", "CA", "CV", "GR", "ZA", "AMI", "AMJ" }) {
            const uint16_t    cellNumber = OpenXLSX::XLWorksheet::columnNameToNumber(cellName);
            const std::string cellName2  = OpenXLSX::XLWorksheet::columnNumberToName(cellNumber);
            REQUIRE(cellName2 == cellName);
        }
    }

    SECTION("Number To Name Conversion")
    {
        for (auto cellNumber : { 1, 5, 25, 26, 27, 52, 53, 54, 100, 200, 512, 1023, 1024, 1025 }) {
            const std::string cellName    = OpenXLSX::XLWorksheet::columnNumberToName(cellNumber);
            const uint16_t    cellNumber2 = OpenXLSX::XLWorksheet::columnNameToNumber(cellName);
            REQUIRE(cellNumber2 == cellNumber);
        }
    }
}
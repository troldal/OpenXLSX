//
// Created by Troldal on 2019-01-21.
//

#include "catch.hpp"
#include <OpenXLSX.hpp>

using namespace OpenXLSX;

TEST_CASE("Test 08: Testing of XLCell/XLCellValue objects")
{
    XLDocument doc;
    doc.open("./TestCell.xlsx");
    auto wbk = doc.workbook();
    auto wks = wbk.worksheet("Sheet1");

    SECTION("CellValue (String)")
    {
        wks.cell("A1").value() = "Hello OpenXLSX!";
        REQUIRE(wks.cell("A1").value().get<std::string>() == "Hello OpenXLSX!");
        REQUIRE(wks.cell(1, 1).value().get<std::string>() == "Hello OpenXLSX!");
        REQUIRE(wks.cell(XLCellReference(1, 1)).value().get<std::string>() == "Hello OpenXLSX!");
        REQUIRE(wks.cell(1, 1).valueType() == XLValueType::String);
    }

    SECTION("CellValue (Integer)")
    {
        wks.cell("B2").value() = 42;
        REQUIRE(wks.cell("B2").value().get<int>() == 42);
        REQUIRE(wks.cell(2, 2).value().get<int>() == 42);
        REQUIRE(wks.cell(XLCellReference(2, 2)).value().get<int>() == 42);
        REQUIRE(wks.cell(2, 2).valueType() == XLValueType::Integer);
    }

    SECTION("CellValue (Float)")
    {
        wks.cell("C3").value() = 3.14159;
        REQUIRE(wks.cell("C3").value().get<double>() == 3.14159);
        REQUIRE(wks.cell(3, 3).value().get<double>() == 3.14159);
        REQUIRE(wks.cell(XLCellReference(3, 3)).value().get<double>() == 3.14159);
        REQUIRE(wks.cell(3, 3).valueType() == XLValueType::Float);
    }

    SECTION("CellValue (Boolean)")
    {
        wks.cell("D4").value() = true;
        REQUIRE(wks.cell("D4").value().get<bool>() == true);
        REQUIRE(wks.cell(4, 4).value().get<bool>() == true);
        REQUIRE(wks.cell(XLCellReference(4, 4)).value().get<bool>() == true);
        REQUIRE(wks.cell(4, 4).valueType() == XLValueType::Boolean);
    }

    SECTION("CellValue (Empty)")
    {
        REQUIRE(wks.cell(5, 5).valueType() == XLValueType::Empty);
    }

    SECTION("CellRange")
    {
        auto rng               = wks.range(XLCellReference(5, 5), XLCellReference(7, 9));
        rng.Cell(1, 1).Value() = "Range";
        REQUIRE(wks.cell(XLCellReference(5, 5)).value().get<std::string>() == "Range");
        REQUIRE(rng.numRows() == 3);
        REQUIRE(rng.numColumns() == 5);

        rng.Transpose(true);
        REQUIRE(rng.numRows() == 5);
        REQUIRE(rng.numColumns() == 3);
    }

    doc.save();
}
//
// Created by Troldal on 2019-01-21.
//

#include "catch.hpp"
#include <OpenXLSX.h>

using namespace OpenXLSX;

TEST_CASE("Test 08: Testing of XLCell/XLCellValue objects") {

XLDocument doc;
doc.OpenDocument("./TestCell.xlsx");
auto wbk = doc.Workbook();
auto wks = wbk.Worksheet("Sheet1");

SECTION("CellValue (String)") {
wks.Cell("A1").
Value() = "Hello OpenXLSX!";
REQUIRE(wks
.Cell("A1").
Value()
.
Get<std::string>()
== "Hello OpenXLSX!");
REQUIRE(wks
.Cell(1, 1).
Value()
.
Get<std::string>()
== "Hello OpenXLSX!");
REQUIRE(wks
.
Cell(XLCellReference(1, 1)
).
Value()
.
Get<std::string>()
== "Hello OpenXLSX!");
REQUIRE(wks
.Cell(1, 1).
ValueType()
== XLValueType::String);
}

SECTION("CellValue (Integer)") {
wks.Cell("B2").
Value() = 42;
REQUIRE(wks
.Cell("B2").
Value()
.
Get<int>()
== 42);
REQUIRE(wks
.Cell(2, 2).
Value()
.
Get<int>()
== 42);
REQUIRE(wks
.
Cell(XLCellReference(2, 2)
).
Value()
.
Get<int>()
== 42);
REQUIRE(wks
.Cell(2, 2).
ValueType()
== XLValueType::Integer);
}

SECTION("CellValue (Float)") {
wks.Cell("C3").
Value() = 3.14159;
REQUIRE(wks
.Cell("C3").
Value()
.
Get<double>()
== 3.14159);
REQUIRE(wks
.Cell(3, 3).
Value()
.
Get<double>()
== 3.14159);
REQUIRE(wks
.
Cell(XLCellReference(3, 3)
).
Value()
.
Get<double>()
== 3.14159);
REQUIRE(wks
.Cell(3, 3).
ValueType()
== XLValueType::Float);
}

SECTION("CellValue (Boolean)") {
wks.Cell("D4").
Value() = true;
REQUIRE(wks
.Cell("D4").
Value()
.
Get<bool>()
== true);
REQUIRE(wks
.Cell(4, 4).
Value()
.
Get<bool>()
== true);
REQUIRE(wks
.
Cell(XLCellReference(4, 4)
).
Value()
.
Get<bool>()
== true);
REQUIRE(wks
.Cell(4, 4).
ValueType()
== XLValueType::Boolean);
}

SECTION("CellValue (Empty)") {
REQUIRE(wks
.Cell(5, 5).
ValueType()
== XLValueType::Empty);
}

SECTION("CellRange") {
auto rng = wks.Range(XLCellReference(5, 5), XLCellReference(7, 9));
rng.Cell(1, 1).
Value() = "Range";
REQUIRE(wks
.
Cell(XLCellReference(5, 5)
).
Value()
.
Get<std::string>()
== "Range");
REQUIRE(rng
.
NumRows()
== 3);
REQUIRE(rng
.
NumColumns()
== 5);

rng.Transpose(true);
REQUIRE(rng
.
NumRows()
== 5);
REQUIRE(rng
.
NumColumns()
== 3);
}

doc.
SaveDocument();

}
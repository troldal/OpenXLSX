//
// Created by Troldal on 2019-01-20.
//

#include "catch.hpp"
#include <OpenXLSX.hpp>

using namespace OpenXLSX;

TEST_CASE("Test 07: Testing of XLCellReference objects")
{
    XLCellReference ref(3, 3);

    SECTION("Address")
    {
        REQUIRE(ref.address() == "C3");
    }

    SECTION("Column")
    {
        REQUIRE(ref.column() == 3);
    }

    SECTION("Row")
    {
        REQUIRE(ref.row() == 3);
    }

    SECTION("SetAddress")
    {
        ref.setAddress("A4");
        REQUIRE(ref.address() == "A4");
        REQUIRE(ref.row() == 4);
        REQUIRE(ref.column() == 1);
    }

    SECTION("SetRowAndColumn")
    {
        ref.setRowAndColumn(2, 5);
        REQUIRE(ref.address() == "E2");
    }
}
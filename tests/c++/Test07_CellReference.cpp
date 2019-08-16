//
// Created by Troldal on 2019-01-20.
//

#include "catch.hpp"
#include <OpenXLSX.h>

using namespace OpenXLSX;

TEST_CASE("Test 07: Testing of XLCellReference objects") {

XLCellReference ref(3, 3);

SECTION("Address") {
REQUIRE(ref
.
Address()
== "C3");
}

SECTION("Column") {
REQUIRE(ref
.
Column()
== 3);
}

SECTION("Row") {
REQUIRE(ref
.
Row()
== 3);
}

SECTION("SetAddress") {
ref.SetAddress("A4");
REQUIRE(ref
.
Address()
== "A4");
REQUIRE(ref
.
Row()
== 4);
REQUIRE(ref
.
Column()
== 1);
}

SECTION("SetRowAndColumn") {
ref.SetRowAndColumn(2, 5);
REQUIRE(ref
.
Address()
== "E2");
}

}
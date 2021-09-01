//
// Created by Kenneth Balslev on 29/08/2021.
//

#include <OpenXLSX.hpp>
#include <catch.hpp>
#include <fstream>

using namespace OpenXLSX;

TEST_CASE("XLDateTime Tests", "[XLFormula]")
{


    SECTION("Conversion")
    {
        XLDateTime dt1(1.5);
        XLDateTime dt2(dt1.tm());
        REQUIRE(dt2.serial() == Approx(1.5));
    }
}


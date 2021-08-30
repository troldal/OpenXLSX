//
// Created by Kenneth Balslev on 29/08/2021.
//

#include <OpenXLSX.hpp>
#include <catch.hpp>
#include <fstream>

using namespace OpenXLSX;

TEST_CASE("XLDateTime Tests", "[XLFormula]")
{
    SECTION("Default Constructor")
    {
        XLDateTime dt;

        REQUIRE(dt.test(1700) == 365);
        REQUIRE(dt.test(1800) == 365);
        REQUIRE(dt.test(1900) == 366); // Due to a historical bug, year 1900 is considered a leap year.
        REQUIRE(dt.test(2100) == 365);
        REQUIRE(dt.test(2200) == 365);
        REQUIRE(dt.test(2300) == 365);
        REQUIRE(dt.test(2500) == 365);
        REQUIRE(dt.test(2600) == 365);

        REQUIRE(dt.test(1600) == 366);
        REQUIRE(dt.test(2000) == 366);
        REQUIRE(dt.test(2400) == 366);

        REQUIRE(dt.test(1992) == 366);
        REQUIRE(dt.test(2000) == 366);
        REQUIRE(dt.test(1990) == 365);

    }

    SECTION("Conversion")
    {
        XLDateTime dt1(15);
        XLDateTime dt2(dt1.timepoint());
        REQUIRE(dt2.serial() == Approx(15));

    }
}


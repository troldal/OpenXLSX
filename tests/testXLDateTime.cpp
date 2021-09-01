//
// Created by Kenneth Balslev on 29/08/2021.
//

#include <OpenXLSX.hpp>
#include <catch.hpp>
#include <fstream>

using namespace OpenXLSX;

TEST_CASE("XLDateTime Tests", "[XLFormula]")
{

    SECTION("Default construction")
    {
        XLDateTime dt;

        REQUIRE(dt.serial() == Approx(1.0));

        auto tm = dt.tm();
        REQUIRE(tm.tm_year == 0);
        REQUIRE(tm.tm_mon == 0);
        REQUIRE(tm.tm_mday == 1);
        REQUIRE(tm.tm_yday == 0);
        REQUIRE(tm.tm_wday == 0);
        REQUIRE(tm.tm_hour == 0);
        REQUIRE(tm.tm_min == 0);
        REQUIRE(tm.tm_sec == 0);
    }

    SECTION("Constructor (serial number)")
    {
        REQUIRE_THROWS(XLDateTime(0.0));

        XLDateTime dt (6069.86742);

        REQUIRE(dt.serial() == Approx(6069.86742));

        auto tm = dt.tm();
        REQUIRE(tm.tm_year == 16);
        REQUIRE(tm.tm_mon == 7);
        REQUIRE(tm.tm_mday == 12);
//        REQUIRE(tm.tm_yday == 0);
        REQUIRE(tm.tm_wday == 6);
        REQUIRE(tm.tm_hour == 20);
        REQUIRE(tm.tm_min == 49);
        REQUIRE(tm.tm_sec == 5);
    }

    SECTION("Constructor (std::tm object)")
    {
        std::tm tmo;

        tmo.tm_year = -1;
        tmo.tm_mon = 0;
        tmo.tm_mday = 0;
        tmo.tm_yday = 0;
        tmo.tm_wday = 0;
        tmo.tm_hour = 0;
        tmo.tm_min = 0;
        tmo.tm_sec = 0;
        REQUIRE_THROWS(XLDateTime(tmo));

        tmo.tm_year = 89;
        tmo.tm_mon = 13;
        REQUIRE_THROWS(XLDateTime(tmo));

        tmo.tm_mon = -1;
        REQUIRE_THROWS(XLDateTime(tmo));

        tmo.tm_mon = 11;
        tmo.tm_mday = 0;
        REQUIRE_THROWS(XLDateTime(tmo));

        tmo.tm_mday = 32;
        REQUIRE_THROWS(XLDateTime(tmo));

        tmo.tm_mday = 30;
        tmo.tm_hour = 17;
        tmo.tm_min = 6;
        tmo.tm_sec = 28;
        tmo.tm_wday = 6;

        XLDateTime dt(tmo);

        REQUIRE(dt.serial() == Approx(32872.7128));
    }

    SECTION("Copy Constructor")
    {
        XLDateTime dt (6069.86742);
        XLDateTime dt2 = dt;

        auto tm = dt2.tm();
        REQUIRE(tm.tm_year == 16);
        REQUIRE(tm.tm_mon == 7);
        REQUIRE(tm.tm_mday == 12);
        REQUIRE(tm.tm_wday == 6);
        REQUIRE(tm.tm_hour == 20);
        REQUIRE(tm.tm_min == 49);
        REQUIRE(tm.tm_sec == 5);
    }

    SECTION("Move Constructor")
    {
        XLDateTime dt (6069.86742);
        XLDateTime dt2 = std::move(dt);

        auto tm = dt2.tm();
        REQUIRE(tm.tm_year == 16);
        REQUIRE(tm.tm_mon == 7);
        REQUIRE(tm.tm_mday == 12);
        REQUIRE(tm.tm_wday == 6);
        REQUIRE(tm.tm_hour == 20);
        REQUIRE(tm.tm_min == 49);
        REQUIRE(tm.tm_sec == 5);
    }

    SECTION("Copy Assignment")
    {
        XLDateTime dt (6069.86742);
        XLDateTime dt2;
        dt2 = dt;

        auto tm = dt2.tm();
        REQUIRE(tm.tm_year == 16);
        REQUIRE(tm.tm_mon == 7);
        REQUIRE(tm.tm_mday == 12);
        REQUIRE(tm.tm_wday == 6);
        REQUIRE(tm.tm_hour == 20);
        REQUIRE(tm.tm_min == 49);
        REQUIRE(tm.tm_sec == 5);
    }

    SECTION("Move Assignment")
    {
        XLDateTime dt (6069.86742);
        XLDateTime dt2;
        dt2 = std::move(dt);

        auto tm = dt2.tm();
        REQUIRE(tm.tm_year == 16);
        REQUIRE(tm.tm_mon == 7);
        REQUIRE(tm.tm_mday == 12);
        REQUIRE(tm.tm_wday == 6);
        REQUIRE(tm.tm_hour == 20);
        REQUIRE(tm.tm_min == 49);
        REQUIRE(tm.tm_sec == 5);
    }

    SECTION("Serial Assignment")
    {
        XLDateTime dt (1.0);
        dt = 6069.86742;

        auto tm = dt.tm();
        REQUIRE(tm.tm_year == 16);
        REQUIRE(tm.tm_mon == 7);
        REQUIRE(tm.tm_mday == 12);
        REQUIRE(tm.tm_wday == 6);
        REQUIRE(tm.tm_hour == 20);
        REQUIRE(tm.tm_min == 49);
        REQUIRE(tm.tm_sec == 5);
    }

    SECTION("std::tm Assignment")
    {
        std::tm tmo;
        tmo.tm_year = 16;
        tmo.tm_mon = 7;
        tmo.tm_mday = 12;
        tmo.tm_wday = 6;
        tmo.tm_hour = 20;
        tmo.tm_min = 49;
        tmo.tm_sec = 5;

        XLDateTime dt (1.0);
        dt = tmo;

        REQUIRE(dt.serial() == Approx(6069.86742));
    }

    SECTION("Implicit conversion")
    {
        std::tm tmo;
        tmo.tm_year = 16;
        tmo.tm_mon = 7;
        tmo.tm_mday = 12;
        tmo.tm_wday = 6;
        tmo.tm_hour = 20;
        tmo.tm_min = 49;
        tmo.tm_sec = 5;

        XLDateTime dt (1.0);
        dt = tmo;

        double serial = dt;
        std::tm result = dt;

        REQUIRE(serial == Approx(6069.86742));
        REQUIRE(result.tm_year == 16);
        REQUIRE(result.tm_mon == 7);
        REQUIRE(result.tm_mday == 12);
        REQUIRE(result.tm_wday == 6);
        REQUIRE(result.tm_hour == 20);
        REQUIRE(result.tm_min == 49);
        REQUIRE(result.tm_sec == 5);
    }
}


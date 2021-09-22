//
// Created by I22696 on 21-09-2021.
//

#include <OpenXLSX.hpp>
#include <catch.hpp>
#include <fstream>

using namespace OpenXLSX;

TEST_CASE("XLColor", "[XLColor]")
{
    SECTION("Constructors") {

        XLColor color;
        REQUIRE(color.alpha() == 255);
        REQUIRE(color.red() == 0);
        REQUIRE(color.green() == 0);
        REQUIRE(color.blue() == 0);
        REQUIRE(color.hex() == "ff000000");

        color = XLColor(128, 128, 128);
        REQUIRE(color.alpha() == 255);
        REQUIRE(color.red() == 128);
        REQUIRE(color.green() == 128);
        REQUIRE(color.blue() == 128);
        REQUIRE(color.hex() == "ff808080");

        color = XLColor(128, 128, 128, 128);
        REQUIRE(color.alpha() == 128);
        REQUIRE(color.red() == 128);
        REQUIRE(color.green() == 128);
        REQUIRE(color.blue() == 128);
        REQUIRE(color.hex() == "80808080");

        color = XLColor("80808080");
        REQUIRE(color.alpha() == 128);
        REQUIRE(color.red() == 128);
        REQUIRE(color.green() == 128);
        REQUIRE(color.blue() == 128);
        REQUIRE(color.hex() == "80808080");

        color = XLColor("ffffff");
        REQUIRE(color.alpha() == 255);
        REQUIRE(color.red() == 255);
        REQUIRE(color.green() == 255);
        REQUIRE(color.blue() == 255);
        REQUIRE(color.hex() == "ffffffff");

        color.set("808080");
        REQUIRE(color.alpha() == 255);
        REQUIRE(color.red() == 128);
        REQUIRE(color.green() == 128);
        REQUIRE(color.blue() == 128);
        REQUIRE(color.hex() == "ff808080");

        color.set(255, 255, 255);
        REQUIRE(color.alpha() == 255);
        REQUIRE(color.red() == 255);
        REQUIRE(color.green() == 255);
        REQUIRE(color.blue() == 255);
        REQUIRE(color.hex() == "ffffffff");

        color.set(0, 0, 0, 0);
        REQUIRE(color.alpha() == 0);
        REQUIRE(color.red() == 0);
        REQUIRE(color.green() == 0);
        REQUIRE(color.blue() == 0);
        REQUIRE(color.hex() == "00000000");

        REQUIRE_THROWS(XLColor("ffff"));
        REQUIRE_THROWS(XLColor("ffffffffff"));

    }
}
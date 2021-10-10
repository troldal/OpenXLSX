//
// Created by Kenneth Balslev on 27/08/2021.
//

#include <OpenXLSX.hpp>
#include <catch.hpp>

using namespace OpenXLSX;

TEST_CASE("XLSheet Tests", "[XLSheet]")
{
    SECTION("XLSheet") {

        XLDocument doc;
        doc.create("./testXLSheet.xlsx");

        auto sht = doc.workbook().sheet(1);
        REQUIRE(sht.name() == "Sheet1");

    }
}
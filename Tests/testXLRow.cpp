//
// Created by Kenneth Balslev on 27/08/2021.
//

#include <OpenXLSX.hpp>
#include <catch.hpp>
#include <fstream>

using namespace OpenXLSX;

TEST_CASE("XLRow Tests", "[XLRow]")
{
    SECTION("XLRow") {

        XLDocument doc;
        doc.create("./testXLRow.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        auto row_3 = wks.row(3);
        REQUIRE(row_3.rowNumber() == 3);

        XLRow copy1 = row_3;
        auto height = copy1.height();
        REQUIRE(copy1.height() == height);

        XLRow copy2 = std::move(copy1);
        auto descent = copy2.descent();
        REQUIRE(copy2.descent() == descent);

        XLRow copy3;
        copy3 = copy2;
        copy3.setHeight(height * 3);
        REQUIRE(copy3.height() == height * 3);

        XLRow copy4;
        copy4 = std::move(copy3);
        copy4.setDescent(descent * 2);
        REQUIRE(copy4.descent() == descent * 2);

        REQUIRE(copy4.isHidden() == false);
        copy4.setHidden(true);
        REQUIRE(copy4.isHidden() == true);


        doc.save();
    }
}
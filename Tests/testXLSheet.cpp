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
        doc.workbook().addWorksheet("Sheet2");

        auto wks1 = doc.workbook().sheet(1);
        REQUIRE(wks1.name() == "Sheet1");

        wks1.setName("OtherName");
        REQUIRE(wks1.name() == "OtherName");

        REQUIRE(wks1.visibility() == XLSheetState::Visible);
//        wks1.setVisibility(XLSheetState::Hidden);
        //REQUIRE(wks1.visibility() == XLSheetState::Hidden);

        auto wks2 = doc.workbook().sheet("Sheet2");
        wks2.setVisibility(XLSheetState::Hidden);

        doc.save();

    }
}
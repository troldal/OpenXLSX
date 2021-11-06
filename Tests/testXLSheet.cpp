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

        auto wks1 = doc.workbook().sheet(1);
        wks1.setName("VeryHidden");

        doc.workbook().addWorksheet("Hidden");
        auto wks2 = doc.workbook().sheet("Hidden");

        doc.workbook().addWorksheet("Visible");
        auto wks3 = doc.workbook().sheet("Visible");


        REQUIRE(wks1.visibility() == XLSheetState::Visible);
        REQUIRE(wks2.visibility() == XLSheetState::Visible);
        REQUIRE(wks3.visibility() == XLSheetState::Visible);

        wks1.setVisibility(XLSheetState::VeryHidden);
        REQUIRE(wks1.visibility() == XLSheetState::VeryHidden);

        wks2.setVisibility(XLSheetState::Hidden);
        REQUIRE(wks2.visibility() == XLSheetState::Hidden);

        REQUIRE_THROWS(wks3.setVisibility(XLSheetState::Hidden));
        wks3.setVisibility(XLSheetState::Visible);

        doc.save();

    }
}
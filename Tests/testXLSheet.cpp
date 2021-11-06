//
// Created by Kenneth Balslev on 27/08/2021.
//

#include <OpenXLSX.hpp>
#include <catch.hpp>

using namespace OpenXLSX;

TEST_CASE("XLSheet Tests", "[XLSheet]")
{
    SECTION("XLSheet Visibility") {

        XLDocument doc;
        doc.create("./testXLSheet1.xlsx");

        auto wks1 = doc.workbook().sheet(1);
        wks1.setName("VeryHidden");
        REQUIRE(wks1.name() == "VeryHidden");

        doc.workbook().addWorksheet("Hidden");
        auto wks2 = doc.workbook().sheet("Hidden");
        REQUIRE(wks2.name() == "Hidden");

        doc.workbook().addWorksheet("Visible");
        auto wks3 = doc.workbook().sheet("Visible");
        REQUIRE(wks3.name() == "Visible");


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

    SECTION("XLSheet Tab Color") {

        XLDocument doc;
        doc.create("./testXLSheet2.xlsx");

        auto wks1 = doc.workbook().sheet(1);
        wks1.setName("Sheet1");
        REQUIRE(wks1.name() == "Sheet1");

        doc.workbook().addWorksheet("Sheet2");
        auto wks2 = doc.workbook().sheet("Sheet2");
        REQUIRE(wks2.name() == "Sheet2");

        doc.workbook().addWorksheet("Sheet3");
        auto wks3 = doc.workbook().sheet("Sheet3");
        REQUIRE(wks3.name() == "Sheet3");

        wks1.setColor(XLColor(255, 0, 0));
        REQUIRE(wks1.color() == XLColor(255, 0, 0));
        REQUIRE_FALSE(wks1.color() == XLColor(0, 0, 0));

        wks2.setColor(XLColor(0, 255, 0));
        REQUIRE(wks2.color() == XLColor(0, 255, 0));
        REQUIRE_FALSE(wks2.color() == XLColor(0, 0, 0));

        wks3.setColor(XLColor(0, 0, 255));
        REQUIRE(wks3.color() == XLColor(0, 0, 255));
        REQUIRE_FALSE(wks3.color() == XLColor(0, 0, 0));

        doc.save();
    }
}
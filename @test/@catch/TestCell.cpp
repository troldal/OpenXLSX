//
// Created by Troldal on 2019-01-21.
//

#include "catch.hpp"
#include <OpenXLSX.h>

using namespace OpenXLSX;

TEST_CASE( "Testing of XLCell/XLCellValue objects") {
    XLDocument doc;
    doc.OpenDocument("./TestCell.xlsx");
    auto wbk = doc.Workbook();
    auto wks = wbk.Worksheet("Sheet1");

    SECTION( "CellValue (String)" ) {
        wks.Cell("A1").Value() = "Hello OpenXLSX!";
        REQUIRE(wks.Cell("A1").Value().Get<std::string>() == "Hello OpenXLSX!");
    }

    SECTION( "CellValue (Integer)" ) {
        wks.Cell("B2").Value() = 42;
        REQUIRE(wks.Cell("B2").Value().Get<int>() == 42);
    }

    SECTION( "CellValue (Float)" ) {
        wks.Cell("C3").Value() = 3.14159;
        REQUIRE(wks.Cell("C3").Value().Get<double>() == 3.14159);
    }

    SECTION( "CellValue (Boolean)" ) {
        wks.Cell("D4").Value() = true;
        REQUIRE(wks.Cell("D4").Value().Get<bool>() == true);
    }

    doc.SaveDocument();

}
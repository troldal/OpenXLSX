//
// Created by Troldal on 2019-01-12.
//

#include "catch.hpp"
#include <fstream>
#include <OpenXLSX.h>

using namespace OpenXLSX;

TEST_CASE( "Testing of XLWorkbook objects") {
    XLDocument doc;
    doc.CreateDocument("./WorkbookTests.xlsx");
    auto wbk = doc.Workbook();

    SECTION( "SheetCount" ) {
        REQUIRE(wbk.SheetCount() == 1);
    }

    SECTION( "IndexOfSheet" ) {
        REQUIRE(wbk.IndexOfSheet("Sheet1") == 1);
    }

    SECTION( "WorksheetCount" ) {
        REQUIRE(wbk.WorksheetCount() == 1);
    }

    SECTION( "ChartsheetCount" ) {
        REQUIRE(wbk.ChartsheetCount() == 0);
    }

    SECTION( "SheetExists" ) {
        REQUIRE(wbk.SheetExists("Sheet1") == true);
    }

    SECTION( "WorksheetExists" ) {
        REQUIRE(wbk.WorksheetExists("Sheet1") == true);
    }

    SECTION( "ChartsheetExists" ) {
        REQUIRE(wbk.ChartsheetExists("Sheet1") == false);
    }

}
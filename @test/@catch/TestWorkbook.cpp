//
// Created by Troldal on 2019-01-12.
//

#include "catch.hpp"
#include <fstream>
#include <OpenXLSX.h>

using namespace OpenXLSX;

TEST_CASE( "Testing of XLWorkbook objects") {
    XLDocument doc;
    doc.OpenDocument("./TestWorkbook.xlsx");
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

    SECTION( "AddWorksheet" ) {
        REQUIRE(wbk.SheetExists("MySheet") == false);
        wbk.AddWorksheet("MySheet");
        doc.SaveDocument();
        REQUIRE(wbk.SheetExists("MySheet") == true);
    }

    SECTION( "CloneWorksheet" ) {
        REQUIRE(wbk.SheetExists("MyClonedSheet") == false);
        wbk.CloneWorksheet("MySheet", "MyClonedSheet");
        doc.SaveDocument();
        REQUIRE(wbk.SheetExists("MyClonedSheet") == true);
    }

    SECTION( "AddChartsheet" ) {
        REQUIRE(wbk.SheetExists("MyChartSheet") == false);
        wbk.AddChartsheet("MyChartSheet");
        doc.SaveDocument();
        REQUIRE(wbk.SheetExists("MyChartSheet") == true);
    }

    SECTION( "MoveSheet" ) {
        REQUIRE(wbk.IndexOfSheet("MyClonedSheet") == 3);
        wbk.MoveSheet("MyClonedSheet", 1);
        doc.SaveDocument();
        REQUIRE(wbk.IndexOfSheet("MyClonedSheet") == 1);
    }

    SECTION( "DeleteSheet" ) {
        REQUIRE(wbk.SheetExists("MySheet") == true);
        wbk.DeleteSheet("MySheet");
        doc.SaveDocument();
        REQUIRE(wbk.SheetExists("MySheet") == false);
    }

}
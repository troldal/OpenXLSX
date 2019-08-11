//
// Created by Troldal on 2019-01-12.
//

#include <catch.hpp>
#include <fstream>
#include <OpenXLSX.h>

using namespace OpenXLSX;

/**
 * @brief
 *
 * @details
 */
TEST_CASE("C++ Interface Test 03: Testing of XLWorkbook objects") {

XLDocument mdoc;
mdoc.OpenDocument("./TestWorkbook.xlsx");
const XLDocument cdoc = mdoc;

auto mwbk = mdoc.Workbook();
auto cwbk = cdoc.Workbook();

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02A: SheetCount()") {
REQUIRE(mwbk
.
SheetCount()
== 1);
REQUIRE(cwbk
.
SheetCount()
== 1);
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02B: IndexOfSheet()") {
REQUIRE(mwbk
.IndexOfSheet("Sheet1") == 1);
REQUIRE(cwbk
.IndexOfSheet("Sheet1") == 1);
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02C: WorksheetCount()") {
REQUIRE(mwbk
.
WorksheetCount()
== 1);
REQUIRE(cwbk
.
WorksheetCount()
== 1);
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02D: ChartsheetCount()") {
REQUIRE(mwbk
.
ChartsheetCount()
== 0);
REQUIRE(cwbk
.
ChartsheetCount()
== 0);
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02E: SheetExists()") {
REQUIRE(mwbk
.SheetExists("Sheet1"));
REQUIRE(cwbk
.SheetExists("Sheet1"));
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02F: WorksheetExists()") {
REQUIRE(mwbk
.WorksheetExists("Sheet1"));
REQUIRE(cwbk
.WorksheetExists("Sheet1"));
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02G: ChartsheetExists()") {
REQUIRE(!mwbk.ChartsheetExists("Sheet1"));
REQUIRE(!cwbk.ChartsheetExists("Sheet1"));
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02H: AddWorksheet()") {
REQUIRE(!mwbk.SheetExists("MySheet"));
mwbk.AddWorksheet("MySheet");
mdoc.
SaveDocument();
REQUIRE(mwbk
.SheetExists("MySheet"));

// ===== Should not compile as cwbk is const =====//
cwbk.AddWorksheet("MySheetConst");
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02I: CloneWorksheet()") {
REQUIRE(!mwbk.SheetExists("MyClonedSheet"));
mwbk.CloneWorksheet("MySheet", "MyClonedSheet");
mdoc.
SaveDocument();
REQUIRE(mwbk
.SheetExists("MyClonedSheet"));

// ===== Should not compile as cwbk is const =====//
//cwbk.CloneWorksheet("MySheetConst", "MyClonedSheetConst");
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02J: AddChartsheet()") {
REQUIRE(!mwbk.SheetExists("MyChartSheet"));
mwbk.AddChartsheet("MyChartSheet");
mdoc.
SaveDocument();
REQUIRE(mwbk
.SheetExists("MyChartSheet"));

// ===== Should not compile as cwbk is const =====//
//cwbk.AddChartsheet("MyChartsheetConst");
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02K: MoveSheet()") {
REQUIRE(mwbk
.IndexOfSheet("MyClonedSheet") == 3);
mwbk.MoveSheet("MyClonedSheet", 1);
mdoc.
SaveDocument();
REQUIRE(mwbk
.IndexOfSheet("MyClonedSheet") == 1);

// ===== Should not compile as cwbk is const =====//
//cwbk.MoveSheet("MyClonedSheetConst");
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02L: Copy construction and assignment") {
REQUIRE(mwbk
.SheetExists("MySheet"));
auto mcopy = mwbk;
REQUIRE(mcopy
.SheetExists("MySheet"));

REQUIRE(cwbk
.SheetExists("Sheet1"));
auto ccopy = cwbk;
REQUIRE(ccopy
.SheetExists("Sheet1"));

mcopy = ccopy;
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02N: DeleteSheet()") {
REQUIRE(mwbk
.SheetExists("MySheet"));
mwbk.DeleteSheet("MySheet");
mdoc.
SaveDocument();
REQUIRE(!mwbk.SheetExists("MySheet"));

// ===== Should not compile as cwbk is const =====//
//cwbk.DeleteSheet()"MySheet";
}

}
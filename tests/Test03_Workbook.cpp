//
// Created by Troldal on 2019-01-12.
//

#include <OpenXLSX.hpp>
#include <catch.hpp>
#include <fstream>

using namespace OpenXLSX;

/**
 * @brief
 *
 * @details
 */
TEST_CASE("C++ Interface Test 03: Testing of XLWorkbook objects")
{
    XLDocument mdoc;
    mdoc.open("./TestWorkbook.xlsx");
    const XLDocument cdoc = mdoc;

    auto mwbk = mdoc.workbook();
    auto cwbk = cdoc.workbook();

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02A: SheetCount()")
    {
        REQUIRE(mwbk.sheetCount() == 1);
        REQUIRE(cwbk.sheetCount() == 1);
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02B: IndexOfSheet()")
    {
        REQUIRE(mwbk.indexOfSheet("Sheet1") == 1);
        REQUIRE(cwbk.indexOfSheet("Sheet1") == 1);
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02C: WorksheetCount()")
    {
        REQUIRE(mwbk.worksheetCount() == 1);
        REQUIRE(cwbk.worksheetCount() == 1);
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02D: ChartsheetCount()")
    {
        REQUIRE(mwbk.chartsheetCount() == 0);
        REQUIRE(cwbk.chartsheetCount() == 0);
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02E: SheetExists()")
    {
        REQUIRE(mwbk.sheetExists("Sheet1"));
        REQUIRE(cwbk.sheetExists("Sheet1"));
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02F: WorksheetExists()")
    {
        REQUIRE(mwbk.worksheetExists("Sheet1"));
        REQUIRE(cwbk.worksheetExists("Sheet1"));
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02G: ChartsheetExists()")
    {
        REQUIRE(!mwbk.chartsheetExists("Sheet1"));
        REQUIRE(!cwbk.chartsheetExists("Sheet1"));
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02H: AddWorksheet()")
    {
        REQUIRE(!mwbk.sheetExists("MySheet"));
        mwbk.addWorksheet("MySheet");
        mdoc.save();
        REQUIRE(mwbk.sheetExists("MySheet"));

        // ===== Should not compile as cwbk is const =====//
        cwbk.addWorksheet("MySheetConst");
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02I: CloneWorksheet()")
    {
        REQUIRE(!mwbk.sheetExists("MyClonedSheet"));
        mwbk.CloneWorksheet("MySheet", "MyClonedSheet");
        mdoc.save();
        REQUIRE(mwbk.sheetExists("MyClonedSheet"));

        // ===== Should not compile as cwbk is const =====//
        // cwbk.CloneWorksheet("MySheetConst", "MyClonedSheetConst");
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02J: AddChartsheet()")
    {
        REQUIRE(!mwbk.sheetExists("MyChartSheet"));
        mwbk.AddChartsheet("MyChartSheet");
        mdoc.save();
        REQUIRE(mwbk.sheetExists("MyChartSheet"));

        // ===== Should not compile as cwbk is const =====//
        // cwbk.AddChartsheet("MyChartsheetConst");
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02K: MoveSheet()")
    {
        REQUIRE(mwbk.indexOfSheet("MyClonedSheet") == 3);
        mwbk.MoveSheet("MyClonedSheet", 1);
        mdoc.save();
        REQUIRE(mwbk.indexOfSheet("MyClonedSheet") == 1);

        // ===== Should not compile as cwbk is const =====//
        // cwbk.MoveSheet("MyClonedSheetConst");
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02L: Copy construction and assignment")
    {
        REQUIRE(mwbk.sheetExists("MySheet"));
        auto mcopy = mwbk;
        REQUIRE(mcopy.sheetExists("MySheet"));

        REQUIRE(cwbk.sheetExists("Sheet1"));
        auto ccopy = cwbk;
        REQUIRE(ccopy.sheetExists("Sheet1"));

        mcopy = ccopy;
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02N: DeleteSheet()")
    {
        REQUIRE(mwbk.sheetExists("MySheet"));
        mwbk.deleteSheet("MySheet");
        mdoc.save();
        REQUIRE(!mwbk.sheetExists("MySheet"));

        // ===== Should not compile as cwbk is const =====//
        // cwbk.DeleteSheet()"MySheet";
    }
}
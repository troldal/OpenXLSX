//
// Created by Troldal on 2019-01-12.
//

#include "catch.hpp"
#include <OpenXLSX.hpp>

using namespace OpenXLSX::Impl;

TEST_CASE("Test 04: Testing of XLSheet objects")
{
    XLDocument doc;
    doc.OpenDocument("./TestSheet.xlsx");
    auto wbk = doc.Workbook();

    SECTION("Name")
    {
        auto sheet = wbk.Sheet(1);
        REQUIRE(sheet.Name() == "Sheet1");
    }

    SECTION("SetName")
    {
        auto sheet = wbk.Sheet(1);
        sheet.SetName("MySheet");
        REQUIRE(sheet.Name() == "MySheet");
    }

    SECTION("State")
    {
        auto sheet = wbk.Sheet(1);
        REQUIRE(sheet.State() == XLSheetState::Visible);
        REQUIRE(sheet.State() != XLSheetState::Hidden);
    }

    SECTION("SetState")
    {
        wbk.AddWorksheet("Sheet2");
        auto sheet = wbk.Sheet(2);
        sheet.SetState(XLSheetState::Hidden);
        REQUIRE(sheet.State() != XLSheetState::Visible);
        REQUIRE(sheet.State() == XLSheetState::Hidden);
    }

    SECTION("Type")
    {
        auto sheet = wbk.Sheet(1);
        REQUIRE(sheet.Type() == XLSheetType::WorkSheet);
        REQUIRE(sheet.Type() != XLSheetType::ChartSheet);
    }

    SECTION("Index")
    {
        auto sheet = wbk.Sheet(1);
        REQUIRE(sheet.Index() == 1);
        REQUIRE(sheet.Index() != 2);
    }

    SECTION("SetIndex")
    {
        auto sheet = wbk.Sheet(2);
        sheet.SetIndex(0);
        REQUIRE(sheet.Index() == 2);
        REQUIRE(sheet.Index() != 1);
    }

    doc.SaveDocument();
}

TEST_CASE("Test 05: Testing of XLWorksheet objects")
{
    XLDocument doc;
    doc.OpenDocument("./Testworksheet.xlsx");
    auto wbk = doc.Workbook();

    SECTION("RowCount")
    {
        auto sheet = wbk.Worksheet("Sheet1");
        REQUIRE(sheet.RowCount() == 1);
    }

    SECTION("ColumnCount")
    {
        auto sheet = wbk.Worksheet("Sheet1");
        REQUIRE(sheet.ColumnCount() == 1);
    }

    SECTION("FirstCell")
    {
        auto sheet = wbk.Worksheet("Sheet1");
        auto ref   = sheet.FirstCell();
        REQUIRE(ref.Address() == "A1");
    }

    doc.SaveDocument();
}
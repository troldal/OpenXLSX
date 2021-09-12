//
// Created by Troldal on 2019-01-12.
//

#include <OpenXLSX.hpp>
#include <catch.hpp>
#include <fstream>

using namespace OpenXLSX;

TEST_CASE("XLCell Tests", "[XLCell]")
{
    SECTION("Default Constructor")
    {
        XLCell cell;

        REQUIRE_FALSE(cell);
        REQUIRE_THROWS(cell.cellReference());
        REQUIRE_FALSE(cell.hasFormula());
        REQUIRE_THROWS(cell.formula());
        REQUIRE_THROWS(cell.formula().set("=1+1"));
        REQUIRE_THROWS(cell.value());

    }

    SECTION("Create from worksheet")
    {
        XLDocument doc;
        doc.create("./testXLCell.xlsx");
        XLWorksheet wks = doc.workbook().sheet(1);
        auto cell = wks.cell("A1");
        cell.value() = 42;

        REQUIRE(cell);
        REQUIRE(cell.cellReference().address() == "A1");
        REQUIRE_FALSE(cell.hasFormula());
        REQUIRE(cell.value().get<int>() == 42);

    }

    SECTION("Copy constructor")
    {
        XLDocument doc;
        doc.create("./testXLCell.xlsx");
        XLWorksheet wks = doc.workbook().sheet(1);
        auto cell = wks.cell("A1");
        cell.value() = 42;

        auto copy = cell;

        REQUIRE(copy);
        REQUIRE(copy.cellReference().address() == "A1");
        REQUIRE_FALSE(copy.hasFormula());
        REQUIRE(copy.value().get<int>() == 42);

    }

    SECTION("Move constructor")
    {
        XLDocument doc;
        doc.create("./testXLCell.xlsx");
        XLWorksheet wks = doc.workbook().sheet(1);
        auto cell = wks.cell("A1");
        cell.value() = 42;

        XLCell copy = std::move(cell);

        REQUIRE(copy);
        REQUIRE(copy.cellReference().address() == "A1");
        REQUIRE_FALSE(copy.hasFormula());
        REQUIRE(copy.value().get<int>() == 42);

    }

    SECTION("Copy assignment operator")
    {
        XLDocument doc;
        doc.create("./testXLCell.xlsx");
        XLWorksheet wks = doc.workbook().sheet(1);
        auto cell = wks.cell("A1");
        cell.value() = 42;

        XLCell copy;
        copy = cell;

        REQUIRE(copy);
        REQUIRE(copy.cellReference().address() == "A1");
        REQUIRE_FALSE(copy.hasFormula());
        REQUIRE(copy.value().get<int>() == 42);

    }

    SECTION("Move assignment operator")
    {
        XLDocument doc;
        doc.create("./testXLCell.xlsx");
        XLWorksheet wks = doc.workbook().sheet(1);
        auto cell = wks.cell("A1");
        cell.value() = 42;

        XLCell copy;
        copy = std::move(cell);

        REQUIRE(copy);
        REQUIRE(copy.cellReference().address() == "A1");
        REQUIRE_FALSE(copy.hasFormula());
        REQUIRE(copy.value().get<int>() == 42);

    }

    SECTION("Setters and Getters")
    {
        XLDocument doc;
        doc.create("./testXLCell.xlsx");
        XLWorksheet wks = doc.workbook().sheet(1);
        auto cell = wks.cell("A1");
        cell.formula().set("=1+1");

        REQUIRE(cell.hasFormula());
        REQUIRE(cell.formula().get() == "=1+1");



    }

}
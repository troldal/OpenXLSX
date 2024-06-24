//
// Created by Troldal on 2019-01-12.
//

#include <OpenXLSX.hpp>
#include <catch.hpp>
#include <filesystem>
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

        const auto copy = cell;
        REQUIRE_FALSE(copy);
        REQUIRE_THROWS(copy.cellReference());
        REQUIRE_FALSE(copy.hasFormula());
        REQUIRE_THROWS(copy.formula());
        REQUIRE_THROWS(copy.value());


    }

    SECTION("Create from worksheet")
    {
        std::filesystem::path fileName{"./testXLCell.xlsx"};
        std::filesystem::remove(fileName);
        XLDocument doc;
        doc.create(fileName);
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
        std::filesystem::path fileName{"./testXLCell.xlsx"};
        std::filesystem::remove(fileName);
        XLDocument doc;
        doc.create(fileName);
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
        std::filesystem::path fileName{"./testXLCell.xlsx"};
        std::filesystem::remove(fileName);
        XLDocument doc;
        doc.create(fileName);
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
        std::filesystem::path fileName{"./testXLCell.xlsx"};
        std::filesystem::remove(fileName);
        XLDocument doc;
        doc.create(fileName);
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
        std::filesystem::path fileName{"./testXLCell.xlsx"};
        std::filesystem::remove(fileName);
        XLDocument doc;
        doc.create(fileName);
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
        std::filesystem::path fileName{"./testXLCell.xlsx"};
        std::filesystem::remove(fileName);
        XLDocument doc;
        doc.create(fileName);
        XLWorksheet wks = doc.workbook().sheet(1);
        auto cell = wks.cell("A1");
        cell.formula().set("=1+1");
        cell.value() = 42;

        REQUIRE(cell.hasFormula());
        REQUIRE(cell.formula().get() == "=1+1");
        REQUIRE(cell.value().get<int>() == 42);


        const auto copy = cell;
        REQUIRE(copy.hasFormula());
        REQUIRE(copy.formula().get() == "=1+1");
        REQUIRE(copy.value().get<int>() == 42);

    }

    SECTION("Relational operators")
    {
        std::filesystem::path fileName{"./testXLCell.xlsx"};
        std::filesystem::remove(fileName);
        auto doc = XLDocument();
        doc.create(fileName);
        auto wks = doc.workbook().worksheet("Sheet1");

        auto cell1 = wks.cell("B2");
        auto cell2 = wks.cell("B2");
        auto cell3 = wks.cell("C3");

        REQUIRE(cell1 == cell2);
        REQUIRE_FALSE(cell1 == cell3);
        REQUIRE(cell1 != cell3);
        REQUIRE_FALSE(cell1 != cell2);
    }

    SECTION("Offset function")
    {
        std::filesystem::path fileName{"./testXLCell.xlsx"};
        std::filesystem::remove(fileName);
        auto doc = XLDocument();
        doc.create(fileName);
        auto wks = doc.workbook().worksheet("Sheet1");

        auto cell1 = wks.cell("B2");
        cell1.value() = "Value1";

        auto cell2 = cell1.offset(1,1);
        cell2.value() = "Value2";

        REQUIRE(cell1.value().get<std::string>() == "Value1");
        REQUIRE(cell2.value().get<std::string>() == "Value2");
        REQUIRE(cell2.cellReference().address() == "C3");

        doc.save();
    }

}
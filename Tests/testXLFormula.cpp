//
// Created by Kenneth Balslev on 27/08/2021.
//

#include <OpenXLSX.hpp>
#include <catch.hpp>
#include <fstream>

using namespace OpenXLSX;

TEST_CASE("XLFormula Tests", "[XLFormula]")
{
    SECTION("Default Constructor")
    {
        XLFormula formula;

        REQUIRE(formula.get().empty());
    }

    SECTION("String Constructor")
    {
        XLFormula formula1("BLAH1");
        REQUIRE(formula1.get() == "BLAH1");

        std::string s = "BLAH2";
        XLFormula formula2(s);
        REQUIRE(formula2.get() == "BLAH2");
    }

    SECTION("Copy Constructor")
    {
        XLFormula formula1("BLAH1");
        REQUIRE(formula1.get() == "BLAH1");

        XLFormula formula2 = formula1;
        REQUIRE(formula2.get() == "BLAH1");

    }

    SECTION("Move Constructor")
    {
        XLFormula formula1("BLAH1");
        REQUIRE(formula1.get() == "BLAH1");

        XLFormula formula2 = std::move(formula1);
        REQUIRE(formula2.get() == "BLAH1");

    }

    SECTION("Copy assignment")
    {
        XLFormula formula1("BLAH1");
        REQUIRE(formula1.get() == "BLAH1");

        XLFormula formula2;
        formula2 = formula1;
        REQUIRE(formula2.get() == "BLAH1");
    }

    SECTION("Move assignment")
    {
        XLFormula formula1("BLAH1");
        REQUIRE(formula1.get() == "BLAH1");

        XLFormula formula2;
        formula2 = std::move(formula1);
        REQUIRE(formula2.get() == "BLAH1");

    }

    SECTION("Clear")
    {
        XLFormula formula1("BLAH1");
        REQUIRE(formula1.get() == "BLAH1");

        formula1.clear();
        REQUIRE(formula1.get().empty());
    }

    SECTION("String assignment")
    {
        XLFormula formula1;
        formula1 = "BLAH1";
        REQUIRE(formula1.get() == "BLAH1");

        XLFormula formula2;
        formula2 = std::string("BLAH2");
        REQUIRE(formula2.get() == "BLAH2");
    }

    SECTION("String setter")
    {
        XLFormula formula1;
        formula1.set("BLAH1");
        REQUIRE(formula1.get() == "BLAH1");

        XLFormula formula2;
        formula2.set(std::string("BLAH2"));
        REQUIRE(formula2.get() == "BLAH2");
    }

    SECTION("Implicit conversion")
    {
        XLFormula formula1;
        formula1.set("BLAH1");
        REQUIRE(formula1.get() == "BLAH1");

        auto result = std::string(formula1);
        REQUIRE(result == "BLAH1");
    }

    SECTION("FormulaProxy")
    {
        XLDocument doc;
        doc.create("./testXLFormula.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.cell("A1").formula() = "=1+1";
        wks.cell("B2").formula() = wks.cell("A1").formula();
        REQUIRE(wks.cell("B2").formula() == XLFormula("=1+1"));

        XLFormula form = wks.cell("B2").formula();
        REQUIRE(form == XLFormula("=1+1"));

        REQUIRE(wks.cell("A1").hasFormula());
        wks.cell("A1").formula().clear();
        REQUIRE_FALSE(wks.cell("A1").hasFormula());
        REQUIRE(wks.cell("B2").formula() == XLFormula("=1+1"));

    }
}
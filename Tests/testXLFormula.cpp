//
// Created by Kenneth Balslev on 27/08/2021.
//

#include <OpenXLSX.hpp>
#include <catch.hpp>
#include <filesystem>
#include <fstream>

using namespace OpenXLSX;

/**
 * @brief convert a cell's string value to a proper XLFormula
 *
 * @param cell the cell with a string that should be used as formula
 * 
 * @todo this should be a mathod of XLCell
 * 
 * @todo the name could be better ...
 */
void evaluateFormulaString(OpenXLSX::XLCell cell)
{
    std::string sFormula { cell.value().get<std::string>() };
    if (sFormula[0] == '=') {
        // remove the '='
        sFormula.erase(0, 1);
    }
    OpenXLSX::XLFormula formula(sFormula);
    cell.value().clear();    // This is vital! If we don't clear the cell's value, Excel won't open the file.
    cell.formula() = formula;
}

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
        std::filesystem::path fileName { "./testXLFormula.xlsx" };
        std::filesystem::remove(fileName);
        XLDocument doc;
        doc.create(fileName);
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

    SECTION("StringToFormula")
    {
        std::filesystem::path fileName { "./testXLFormula.xlsx" };
        std::filesystem::remove(fileName);
        XLDocument doc;
        doc.create(fileName);
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.cell("A1").value() = "Values";
        wks.cell("B1").value() = "Calculations";
        wks.cell("C1").value() = "More Values";

        for (uint16_t i = 2; i <= 10; i++) {
            std::string cellNameA       = "A" + std::to_string(i);
            std::string cellNameB       = "B" + std::to_string(i);
            std::string cellNameC       = "C" + std::to_string(i);
            wks.cell(cellNameA).value() = double(i + i * 0.10);
            wks.cell(cellNameB).value() = std::string("= " + cellNameA + " * " + cellNameC);
            wks.cell(cellNameC).value() = double(i * 5 + i * 1.1);
        }

        // This works: The cells F5 and G7 have a valid formula and the
        // resulting Excel file is opened by Excel without any complaints.
        const std::string sFormula_F5 { "2*A2" };
        const std::string sFormula_G7 { sFormula_F5 + "+3*A3" };
        wks.cell("F5").formula() = sFormula_F5;
        wks.cell("G7").formula() = sFormula_G7;
        REQUIRE(wks.cell("F5").formula() == XLFormula(sFormula_F5));
        REQUIRE(wks.cell("G7").formula() == XLFormula(sFormula_G7));

        // Now let's try to convert the string in cell B4 to a proper formula:
        OpenXLSX::XLCell cell = wks.cell("B4");
        std::string      sFormula_B4 { cell.value().get<std::string>() };
        if (sFormula_B4[0] == '=') {
            sFormula_B4.erase(0, 1);
        }
        OpenXLSX::XLFormula formula(sFormula_B4);
        cell.value().clear();    // This is vital! If we don't clear the cell's value, Excel won't open the file.
        cell.formula() = formula;
        REQUIRE(wks.cell("B4").formula() == formula);

        // OK, try something like "=$C2+C$3*$B$4"
        wks.cell("E9").value() = "=$C2+C$3*$B$4";
        evaluateFormulaString(wks.cell("E9"));
        REQUIRE(wks.cell("E9").formula() == XLFormula("$C2+C$3*$B$4"));

        doc.save();
        doc.close();
    }
}
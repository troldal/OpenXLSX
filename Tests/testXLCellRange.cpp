//
// Created by Kenneth Balslev on 19/09/2021.
//

#include <OpenXLSX.hpp>
#include <catch.hpp>
#include <fstream>

using namespace OpenXLSX;

TEST_CASE("XLCellRange Tests", "[XLCellRange]")
{
    XLDocument doc;
    doc.create("./testXLCellRange.xlsx");
    auto wks = doc.workbook().worksheet("Sheet1");

    SECTION("Constructor")
    {
        auto rng = wks.range(XLCellReference("B2"), XLCellReference("D4"));
        for (auto cl : rng) cl.value() = "Value";
        doc.save();

        REQUIRE(wks.cell("B2").value().get<std::string>() == "Value");
        REQUIRE(wks.cell("C2").value().get<std::string>() == "Value");
        REQUIRE(wks.cell("D2").value().get<std::string>() == "Value");
        REQUIRE(wks.cell("B3").value().get<std::string>() == "Value");
        REQUIRE(wks.cell("C3").value().get<std::string>() == "Value");
        REQUIRE(wks.cell("D3").value().get<std::string>() == "Value");
        REQUIRE(wks.cell("B4").value().get<std::string>() == "Value");
        REQUIRE(wks.cell("C4").value().get<std::string>() == "Value");
        REQUIRE(wks.cell("D4").value().get<std::string>() == "Value");
    }

    SECTION("Copy Constructor")
    {
        auto rng = wks.range(XLCellReference("B2"), XLCellReference("D4"));
        for (auto cl : rng) cl.value() = "Value";
        doc.save();

        XLCellRange rng2 = rng;
        for (auto cl : rng2) cl.value() = "Value2";
        doc.save();

        REQUIRE(wks.cell("B2").value().get<std::string>() == "Value2");
        REQUIRE(wks.cell("C2").value().get<std::string>() == "Value2");
        REQUIRE(wks.cell("D2").value().get<std::string>() == "Value2");
        REQUIRE(wks.cell("B3").value().get<std::string>() == "Value2");
        REQUIRE(wks.cell("C3").value().get<std::string>() == "Value2");
        REQUIRE(wks.cell("D3").value().get<std::string>() == "Value2");
        REQUIRE(wks.cell("B4").value().get<std::string>() == "Value2");
        REQUIRE(wks.cell("C4").value().get<std::string>() == "Value2");
        REQUIRE(wks.cell("D4").value().get<std::string>() == "Value2");
    }

    SECTION("Move Constructor")
    {
        auto rng = wks.range(XLCellReference("B2"), XLCellReference("D4"));
        for (auto cl : rng) cl.value() = "Value";
        doc.save();

        XLCellRange rng2 {std::move(rng)};
        for (auto cl : rng2) cl.value() = "Value3";
        doc.save();

        REQUIRE(wks.cell("B2").value().get<std::string>() == "Value3");
        REQUIRE(wks.cell("C2").value().get<std::string>() == "Value3");
        REQUIRE(wks.cell("D2").value().get<std::string>() == "Value3");
        REQUIRE(wks.cell("B3").value().get<std::string>() == "Value3");
        REQUIRE(wks.cell("C3").value().get<std::string>() == "Value3");
        REQUIRE(wks.cell("D3").value().get<std::string>() == "Value3");
        REQUIRE(wks.cell("B4").value().get<std::string>() == "Value3");
        REQUIRE(wks.cell("C4").value().get<std::string>() == "Value3");
        REQUIRE(wks.cell("D4").value().get<std::string>() == "Value3");
    }

    SECTION("Copy Assignment")
    {
        auto rng = wks.range(XLCellReference("B2"), XLCellReference("D4"));
        for (auto cl : rng) cl.value() = "Value";
        doc.save();

        XLCellRange rng2 = wks.range();
        rng2 = rng;
        for (auto cl : rng2) cl.value() = "Value4";
        doc.save();

        REQUIRE(wks.cell("B2").value().get<std::string>() == "Value4");
        REQUIRE(wks.cell("C2").value().get<std::string>() == "Value4");
        REQUIRE(wks.cell("D2").value().get<std::string>() == "Value4");
        REQUIRE(wks.cell("B3").value().get<std::string>() == "Value4");
        REQUIRE(wks.cell("C3").value().get<std::string>() == "Value4");
        REQUIRE(wks.cell("D3").value().get<std::string>() == "Value4");
        REQUIRE(wks.cell("B4").value().get<std::string>() == "Value4");
        REQUIRE(wks.cell("C4").value().get<std::string>() == "Value4");
        REQUIRE(wks.cell("D4").value().get<std::string>() == "Value4");
    }

    SECTION("Move Assignment")
    {
        auto rng = wks.range(XLCellReference("B2"), XLCellReference("D4"));
        for (auto cl : rng) cl.value() = "Value";
        doc.save();

        XLCellRange rng2 = wks.range();
        rng2 = std::move(rng);
        for (auto cl : rng2) cl.value() = "Value5";
        doc.save();

        REQUIRE(wks.cell("B2").value().get<std::string>() == "Value5");
        REQUIRE(wks.cell("C2").value().get<std::string>() == "Value5");
        REQUIRE(wks.cell("D2").value().get<std::string>() == "Value5");
        REQUIRE(wks.cell("B3").value().get<std::string>() == "Value5");
        REQUIRE(wks.cell("C3").value().get<std::string>() == "Value5");
        REQUIRE(wks.cell("D3").value().get<std::string>() == "Value5");
        REQUIRE(wks.cell("B4").value().get<std::string>() == "Value5");
        REQUIRE(wks.cell("C4").value().get<std::string>() == "Value5");
        REQUIRE(wks.cell("D4").value().get<std::string>() == "Value5");
    }

    SECTION("Functions")
    {
        auto rng = wks.range(XLCellReference("B2"), XLCellReference("D4"));
        for(auto iter = rng.begin(); iter != rng.end(); iter++)
            iter->value() = "Value";

        doc.save();

        REQUIRE(rng.numColumns() == 3);
        REQUIRE(rng.numRows() == 3);

        rng.clear();

        REQUIRE(wks.cell("B2").value().get<std::string>().empty());
        REQUIRE(wks.cell("C2").value().get<std::string>().empty());
        REQUIRE(wks.cell("D2").value().get<std::string>().empty());
        REQUIRE(wks.cell("B3").value().get<std::string>().empty());
        REQUIRE(wks.cell("C3").value().get<std::string>().empty());
        REQUIRE(wks.cell("D3").value().get<std::string>().empty());
        REQUIRE(wks.cell("B4").value().get<std::string>().empty());
        REQUIRE(wks.cell("C4").value().get<std::string>().empty());
        REQUIRE(wks.cell("D4").value().get<std::string>().empty());
    }

    SECTION("XLCellIterator")
    {
        auto rng = wks.range(XLCellReference("B2"), XLCellReference("D4"));
        for(auto iter = rng.begin(); iter != rng.end(); iter++)
            iter->value() = "Value";

        auto begin = rng.begin();
        REQUIRE(begin->cellReference().address() == "B2");

        auto iter2 = ++begin;
        REQUIRE(iter2->cellReference().address() == "C2");

        XLCellIterator iter3 = std::move(++begin) ;
        REQUIRE(iter3->cellReference().address() == "D2");

        auto iter4 = rng.begin();
        iter4 = ++iter3;
        REQUIRE(iter4->cellReference().address() == "B3");

        auto iter5 = rng.begin();
        iter5 = std::move(++iter3);
        REQUIRE(iter5->cellReference().address() == "C3");

        auto it1 = rng.begin();
        auto it2 = rng.begin();
        auto it3 = rng.end();

        REQUIRE(it1 == it2);
        REQUIRE_FALSE(it1 != it2);
        REQUIRE(it1 != it3);
        REQUIRE_FALSE(it1 == it3);

        REQUIRE(std::distance(it1, it3) == 9);

    }

}
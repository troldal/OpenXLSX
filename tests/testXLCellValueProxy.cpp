//
// Created by Kenneth Balslev on 26/08/2021.
//

#include <OpenXLSX.hpp>
#include <catch.hpp>
#include <fstream>

using namespace OpenXLSX;

TEST_CASE("XLCellValueProxy Tests", "[XLCellValue]")
{
    SECTION("XLCellValueProxy conversion to XLCellValue (XLCellValue constructor)")
    {
        XLDocument doc;
        doc.create("./testXLCellValueProxy.xlsx");
        XLWorksheet wks = doc.workbook().sheet(1);

        {
            wks.cell("A1").value() = "Hello OpenXLSX!";
            XLCellValue value                  = wks.cell("A1").value();
            REQUIRE(value.type() == XLValueType::String);
            REQUIRE(value.typeAsString() == "string");
            REQUIRE(value.get<std::string>() == "Hello OpenXLSX!");
            REQUIRE_THROWS(value.get<int>());
            REQUIRE_THROWS(value.get<double>());
            REQUIRE_THROWS(value.get<bool>());
        }

        {
            wks.cell("A1").value() = 3.14159;
            XLCellValue value                  = wks.cell("A1").value();
            REQUIRE(value.type() == XLValueType::Float);
            REQUIRE(value.typeAsString() == "float");
            REQUIRE_THROWS(value.get<std::string>());
            REQUIRE_THROWS(value.get<int>());
            REQUIRE(value.get<double>() == 3.14159);
            REQUIRE_THROWS(value.get<bool>());
        }

        {
            wks.cell("A1").value() = 42;
            XLCellValue value      = wks.cell("A1").value();
            REQUIRE(value.type() == XLValueType::Integer);
            REQUIRE(value.typeAsString() == "integer");
            REQUIRE_THROWS(value.get<std::string>());
            REQUIRE(value.get<int>() == 42);
            REQUIRE_THROWS(value.get<double>());
            REQUIRE_THROWS(value.get<bool>());
        }

        {
            wks.cell("A1").value() = true;
            XLCellValue value = wks.cell("A1").value();
            REQUIRE(value.type() == XLValueType::Boolean);
            REQUIRE(value.typeAsString() == "boolean");
            REQUIRE_THROWS(value.get<std::string>());
            REQUIRE_THROWS(value.get<int>());
            REQUIRE_THROWS(value.get<double>());
            REQUIRE(value.get<bool>() == true);
        }
    }

    SECTION("XLCellValueProxy conversion to XLCellValue (XLCellValue assignment operator)")
    {
        XLCellValue value;
        XLDocument doc;
        doc.create("./testXLCellValueProxy.xlsx");
        XLWorksheet wks = doc.workbook().sheet(1);

        wks.cell("A1").value() = "Hello OpenXLSX!";
        value = wks.cell("A1").value();
        REQUIRE(value.type() == XLValueType::String);
        REQUIRE(value.typeAsString() == "string");
        REQUIRE(value.get<std::string>() == "Hello OpenXLSX!");
        REQUIRE_THROWS(value.get<int>());
        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());

        wks.cell("A1").value() = 3.14159;
        value = wks.cell("A1").value();
        REQUIRE(value.type() == XLValueType::Float);
        REQUIRE(value.typeAsString() == "float");
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE_THROWS(value.get<int>());
        REQUIRE(value.get<double>() == 3.14159);
        REQUIRE_THROWS(value.get<bool>());

        wks.cell("A1").value() = 42;
        value = wks.cell("A1").value();
        REQUIRE(value.type() == XLValueType::Integer);
        REQUIRE(value.typeAsString() == "integer");
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE(value.get<int>() == 42);
        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());

        wks.cell("A1").value() = true;
        value = wks.cell("A1").value();
        REQUIRE(value.type() == XLValueType::Boolean);
        REQUIRE(value.typeAsString() == "boolean");
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE_THROWS(value.get<int>());
        REQUIRE_THROWS(value.get<double>());
        REQUIRE(value.get<bool>() == true);
    }

    SECTION("XLCellValueProxy copy assignment")
    {
        XLCellValue value;
        XLDocument doc;
        doc.create("./testXLCellValueProxy.xlsx");
        XLWorksheet wks = doc.workbook().sheet(1);

        wks.cell("A1").value() = "Hello OpenXLSX!";
        wks.cell("A2").value() = wks.cell("A1").value();

        REQUIRE(wks.cell("A2").value().type() == XLValueType::String);
        REQUIRE(wks.cell("A2").value().typeAsString() == "string");
        REQUIRE(wks.cell("A2").value().get<std::string>() == "Hello OpenXLSX!");
        REQUIRE_THROWS(wks.cell("A2").value().get<int>());
        REQUIRE_THROWS(wks.cell("A2").value().get<double>());
        REQUIRE_THROWS(wks.cell("A2").value().get<bool>());
    }
}
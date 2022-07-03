//
// Created by Troldal on 2019-01-12.
//

#include <OpenXLSX.hpp>
#include <catch.hpp>
#include <fstream>

using namespace OpenXLSX;

/**
 * @brief The purpose of this test case is to test the creation of XLDocument objects. Each section section
 * tests document creation using a different method. In addition, saving, closing and copying is tested.
 */
TEST_CASE("XLCellValue Tests", "[XLCellValue]")
{
    SECTION("Default Constructor")
    {
        XLCellValue value;

        REQUIRE(value.type() == XLValueType::Empty);
        REQUIRE(value.typeAsString() == "empty");
        REQUIRE(value.get<std::string>().empty());
        REQUIRE_THROWS(value.get<int>());
        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Float Constructor")
    {
        XLCellValue value(3.14159);

        REQUIRE(value.type() == XLValueType::Float);
        REQUIRE(value.typeAsString() == "float");
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE_THROWS(value.get<int>());
        REQUIRE(value.get<double>() == 3.14159);
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Integer Constructor")
    {
        XLCellValue value(42);

        REQUIRE(value.type() == XLValueType::Integer);
        REQUIRE(value.typeAsString() == "integer");
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE(value.get<int>() == 42);
        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Boolean Constructor")
    {
        XLCellValue value(true);

        REQUIRE(value.type() == XLValueType::Boolean);
        REQUIRE(value.typeAsString() == "boolean");
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE_THROWS(value.get<int>());
        REQUIRE_THROWS(value.get<double>());
        REQUIRE(value.get<bool>() == true);
    }

    SECTION("String Constructor")
    {
        XLCellValue value("Hello OpenXLSX!");

        REQUIRE(value.type() == XLValueType::String);
        REQUIRE(value.typeAsString() == "string");
        REQUIRE(value.get<std::string>() == "Hello OpenXLSX!");
        REQUIRE_THROWS(value.get<int>());
        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Copy Constructor")
    {
        XLCellValue value("Hello OpenXLSX!");
        auto        copy = value;

        REQUIRE(copy.type() == XLValueType::String);
        REQUIRE(value.typeAsString() == "string");
        REQUIRE(copy.get<std::string>() == "Hello OpenXLSX!");
        REQUIRE_THROWS(copy.get<int>());
        REQUIRE_THROWS(copy.get<double>());
        REQUIRE_THROWS(copy.get<bool>());
    }

    SECTION("Move Constructor")
    {
        XLCellValue value("Hello OpenXLSX!");
        auto        copy = std::move(value);

        REQUIRE(copy.type() == XLValueType::String);
        REQUIRE(value.typeAsString() == "string");
        REQUIRE(copy.get<std::string>() == "Hello OpenXLSX!");
        REQUIRE_THROWS(copy.get<int>());
        REQUIRE_THROWS(copy.get<double>());
        REQUIRE_THROWS(copy.get<bool>());
    }

    SECTION("Copy Assignment")
    {
        XLCellValue value("Hello OpenXLSX!");
        XLCellValue copy(1);
        copy = value;

        REQUIRE(copy.type() == XLValueType::String);
        REQUIRE(value.typeAsString() == "string");
        REQUIRE(copy.get<std::string>() == "Hello OpenXLSX!");
        REQUIRE_THROWS(copy.get<int>());
        REQUIRE_THROWS(copy.get<double>());
        REQUIRE_THROWS(copy.get<bool>());
    }

    SECTION("Move Assignment")
    {
        XLCellValue value("Hello OpenXLSX!");
        XLCellValue copy(1);
        copy = std::move(value);

        REQUIRE(copy.type() == XLValueType::String);
        REQUIRE(value.typeAsString() == "string");
        REQUIRE(copy.get<std::string>() == "Hello OpenXLSX!");
        REQUIRE_THROWS(copy.get<int>());
        REQUIRE_THROWS(copy.get<double>());
        REQUIRE_THROWS(copy.get<bool>());
    }

    SECTION("Float Assignment")
    {
        XLCellValue value;
        value = 3.14159;

        REQUIRE(value.type() == XLValueType::Float);
        REQUIRE(value.typeAsString() == "float");
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE_THROWS(value.get<int>());
        REQUIRE(value.get<double>() == 3.14159);
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Integer Assignment")
    {
        XLCellValue value;
        value = 42;

        REQUIRE(value.type() == XLValueType::Integer);
        REQUIRE(value.typeAsString() == "integer");
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE(value.get<int>() == 42);
        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Boolean Assignment")
    {
        XLCellValue value;
        value = true;

        REQUIRE(value.type() == XLValueType::Boolean);
        REQUIRE(value.typeAsString() == "boolean");
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE_THROWS(value.get<int>());
        REQUIRE_THROWS(value.get<double>());
        REQUIRE(value.get<bool>() == true);
    }

    SECTION("String Assignment")
    {
        XLCellValue value;
        value = "Hello OpenXLSX!";

        REQUIRE(value.type() == XLValueType::String);
        REQUIRE(value.typeAsString() == "string");
        REQUIRE(value.get<std::string>() == "Hello OpenXLSX!");
        REQUIRE_THROWS(value.get<int>());
        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("XLCellValueProxy Assignment")
    {
        XLCellValue value;
        XLDocument doc;
        doc.create("./testXLCellValue.xlsx");
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

    SECTION("Set Float ")
    {
        XLCellValue value;
        value.set(3.14159);

        REQUIRE(value.type() == XLValueType::Float);
        REQUIRE(value.typeAsString() == "float");
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE_THROWS(value.get<int>());
        REQUIRE(value.get<double>() == 3.14159);
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Set Nan Float")
    {
        XLCellValue value;
        value.set(std::numeric_limits<double>::quiet_NaN());

        REQUIRE(value.type() == XLValueType::Error);
        REQUIRE(value.typeAsString() == "error");
        REQUIRE(value.get<std::string>() == "#NUM!");
        REQUIRE_THROWS(value.get<int>());
//        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());
    }

        SECTION("Set Inf Float")
    {
        XLCellValue value;
        value.set(std::numeric_limits<double>::infinity());

        REQUIRE(value.type() == XLValueType::Error);
        REQUIRE(value.typeAsString() == "error");
        REQUIRE(value.get<std::string>() == "#NUM!");
        REQUIRE_THROWS(value.get<int>());
//        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Set Integer")
    {
        XLCellValue value;
        value.set(42);

        REQUIRE(value.type() == XLValueType::Integer);
        REQUIRE(value.typeAsString() == "integer");
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE(value.get<int>() == 42);
        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Set Boolean")
    {
        XLCellValue value;
        value.set(true);

        REQUIRE(value.type() == XLValueType::Boolean);
        REQUIRE(value.typeAsString() == "boolean");
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE_THROWS(value.get<int>());
        REQUIRE_THROWS(value.get<double>());
        REQUIRE(value.get<bool>() == true);
    }

    SECTION("Set String")
    {
        XLCellValue value;
        value.set("Hello OpenXLSX!");

        REQUIRE(value.type() == XLValueType::String);
        REQUIRE(value.typeAsString() == "string");
        REQUIRE(value.get<std::string>() == "Hello OpenXLSX!");
        REQUIRE_THROWS(value.get<int>());
        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Clear")
    {
        XLCellValue value;
        value.set("Hello OpenXLSX!");
        value.clear();

        REQUIRE(value.type() == XLValueType::Empty);
        REQUIRE(value.typeAsString() == "empty");
        REQUIRE(value.get<std::string>().empty());
        REQUIRE_THROWS(value.get<int>());
        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Set Error")
    {
        XLCellValue value;
        value.set("Hello OpenXLSX!");
        value.setError("#N/A");

        REQUIRE(value.type() == XLValueType::Error);
        REQUIRE(value.typeAsString() == "error");
        REQUIRE(value.get<std::string>() == "#N/A");
        REQUIRE_THROWS(value.get<int>());
//        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Implicit conversion to supported types")
    {
        XLCellValue value;

        value ="Hello OpenXLSX!";
        auto result1 = value.get<std::string>();
        REQUIRE(result1 == "Hello OpenXLSX!");
        REQUIRE_THROWS(static_cast<int>(value));
        REQUIRE_THROWS(static_cast<double>(value));
        REQUIRE_THROWS(static_cast<bool>(value));

        value = 42;
        auto result2 = static_cast<int>(value);
        REQUIRE(result2 == 42);
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE_THROWS(static_cast<double>(value));
        REQUIRE_THROWS(static_cast<bool>(value));

        value = 3.14159;
        auto result3 = static_cast<double>(value);
        REQUIRE(result3 == 3.14159);
        REQUIRE_THROWS(static_cast<int>(value));
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE_THROWS(static_cast<bool>(value));

        value = true;
        auto result4 = static_cast<bool>(value);
        REQUIRE(result4 == true);
        REQUIRE_THROWS(static_cast<int>(value));
        REQUIRE_THROWS(static_cast<double>(value));
        REQUIRE_THROWS(value.get<std::string>());

    }
}

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
        REQUIRE(value.get<std::string>().empty());
        REQUIRE_THROWS(value.get<int>());
        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Float Constructor")
    {
        XLCellValue value(3.14159);

        REQUIRE(value.type() == XLValueType::Float);
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE_THROWS(value.get<int>());
        REQUIRE(value.get<double>() == 3.14159);
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Integer Constructor")
    {
        XLCellValue value(42);

        REQUIRE(value.type() == XLValueType::Integer);
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE(value.get<int>() == 42);
        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Boolean Constructor")
    {
        XLCellValue value(true);

        REQUIRE(value.type() == XLValueType::Boolean);
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE_THROWS(value.get<int>());
        REQUIRE_THROWS(value.get<double>());
        REQUIRE(value.get<bool>() == true);
    }

    SECTION("String Constructor")
    {
        XLCellValue value("Hello OpenXLSX!");

        REQUIRE(value.type() == XLValueType::String);
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
        REQUIRE(value.get<std::string>() == "Hello OpenXLSX!");
        REQUIRE_THROWS(value.get<int>());
        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Set Float ")
    {
        XLCellValue value;
        value.set(3.14159);

        REQUIRE(value.type() == XLValueType::Float);
        REQUIRE_THROWS(value.get<std::string>());
        REQUIRE_THROWS(value.get<int>());
        REQUIRE(value.get<double>() == 3.14159);
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Set Integer")
    {
        XLCellValue value;
        value.set(42);

        REQUIRE(value.type() == XLValueType::Integer);
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
        REQUIRE(value.get<std::string>().empty());
        REQUIRE_THROWS(value.get<int>());
        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());
    }

    SECTION("Set Error")
    {
        XLCellValue value;
        value.set("Hello OpenXLSX!");
        value.setError();

        REQUIRE(value.type() == XLValueType::Error);
        REQUIRE(value.get<std::string>().empty());
        REQUIRE_THROWS(value.get<int>());
        REQUIRE_THROWS(value.get<double>());
        REQUIRE_THROWS(value.get<bool>());
    }
}
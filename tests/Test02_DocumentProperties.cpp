//
// Created by Troldal on 2019-01-13.
//

#include "catch.hpp"
#include <OpenXLSX.hpp>
#include <fstream>

#include <ctime>
#include <iomanip>
#include <iostream>
#include <sstream>

using namespace OpenXLSX;

/**
 * @brief
 *
 * @details
 */
TEST_CASE("C++ Interface Test 02: Testing of Document Properties")
{
    XLDocument  doc;
    std::string file = "./TestDocumentProperties.xlsx";
    doc.open(file);

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02A: GetProperty")
    {
        REQUIRE(doc.property(XLProperty::Title).empty());
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02B: Test SetProperty - Title")
    {
        doc.setProperty(XLProperty::Title, "TitleTest");
        doc.save();
        doc.close();
        doc.open(file);
        REQUIRE(doc.property(XLProperty::Title) == "TitleTest");
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02C: Test SetProperty - Title (Duplicate)")
    {
        doc.setProperty(XLProperty::Title, "TitleTest");
        doc.save();
        doc.close();
        doc.open(file);
        REQUIRE(doc.property(XLProperty::Title) == "TitleTest");
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02D: Test SetProperty - Subject")
    {
        doc.setProperty(XLProperty::Subject, "SubjectTest");
        doc.save();
        doc.close();
        doc.open(file);
        REQUIRE(doc.property(XLProperty::Subject) == "SubjectTest");
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02E: Test SetProperty - Creator")
    {
        doc.setProperty(XLProperty::Creator, "CreatorTest");
        doc.save();
        doc.close();
        doc.open(file);
        REQUIRE(doc.property(XLProperty::Creator) == "CreatorTest");
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02F: Test SetProperty - Keywords")
    {
        doc.setProperty(XLProperty::Keywords, "A, B, C");
        doc.save();
        doc.close();
        doc.open(file);
        REQUIRE(doc.property(XLProperty::Keywords) == "A, B, C");
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02G: Test SetProperty - Description")
    {
        doc.setProperty(XLProperty::Description, "DescriptionTest");
        doc.save();
        doc.close();
        doc.open(file);
        REQUIRE(doc.property(XLProperty::Description) == "DescriptionTest");
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02H: Test SetProperty - LastModifiedBy")
    {
        doc.setProperty(XLProperty::LastModifiedBy, "LastModifiedByTest");
        doc.save();
        doc.close();
        doc.open(file);
        REQUIRE(doc.property(XLProperty::LastModifiedBy) == "LastModifiedByTest");
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02I: Test SetProperty - LastPrinted")
    {
        std::time_t t  = std::time(nullptr);
        std::tm     tm = *std::gmtime(&t);

        std::stringstream ss;
        ss << std::put_time(&tm, "%FT%TZ");
        auto datetime = ss.str();

        doc.setProperty(XLProperty::LastPrinted, datetime);
        doc.save();
        doc.close();
        doc.open(file);
        REQUIRE(doc.property(XLProperty::LastPrinted) == datetime);
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02J: Test SetProperty - CreationDate")
    {
        std::time_t t  = std::time(nullptr);
        std::tm     tm = *std::gmtime(&t);

        std::stringstream ss;
        ss << std::put_time(&tm, "%FT%TZ");
        auto datetime = ss.str();

        doc.setProperty(XLProperty::CreationDate, datetime);
        doc.save();
        doc.close();
        doc.open(file);
        REQUIRE(doc.property(XLProperty::CreationDate) == datetime);
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02K: Test SetProperty - ModificationDate")
    {
        std::time_t t  = std::time(nullptr);
        std::tm     tm = *std::gmtime(&t);

        std::stringstream ss;
        ss << std::put_time(&tm, "%FT%TZ");
        auto datetime = ss.str();

        doc.setProperty(XLProperty::ModificationDate, datetime);
        doc.save();
        doc.close();
        doc.open(file);
        REQUIRE(doc.property(XLProperty::ModificationDate) == datetime);
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02L: Test SetProperty - Category")
    {
        doc.setProperty(XLProperty::Category, "CategoryTest");
        doc.save();
        doc.close();
        doc.open(file);
        REQUIRE(doc.property(XLProperty::Category) == "CategoryTest");
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02M: Test SetProperty - Application")
    {
        doc.setProperty(XLProperty::Application, "ApplicationTest");
        doc.save();
        doc.close();
        doc.open(file);
        REQUIRE(doc.property(XLProperty::Application) == "ApplicationTest");
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02N: Test SetProperty - DocSecurity")
    {
        doc.setProperty(XLProperty::DocSecurity, "4");
        doc.save();
        doc.close();
        doc.open(file);
        REQUIRE(doc.property(XLProperty::DocSecurity) == "4");
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02O: Test SetProperty - ScaleCrop")
    {
        doc.setProperty(XLProperty::ScaleCrop, "false");
        doc.save();
        doc.close();
        doc.open(file);
        REQUIRE(doc.property(XLProperty::ScaleCrop) == "false");
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02P: Test SetProperty - Manager")
    {
        doc.setProperty(XLProperty::Manager, "ManagerTest");
        doc.save();
        doc.close();
        doc.open(file);
        REQUIRE(doc.property(XLProperty::Manager) == "ManagerTest");
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02Q: Test SetProperty - Company")
    {
        doc.setProperty(XLProperty::Company, "CompanyTest");
        doc.save();
        doc.close();
        doc.open(file);
        REQUIRE(doc.property(XLProperty::Company) == "CompanyTest");
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02R: Test SetProperty - LinksUpToDate")
    {
        doc.setProperty(XLProperty::LinksUpToDate, "false");
        doc.save();
        doc.close();
        doc.open(file);
        REQUIRE(doc.property(XLProperty::LinksUpToDate) == "false");
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02S: Test SetProperty - SharedDoc")
    {
        doc.setProperty(XLProperty::SharedDoc, "false");
        doc.save();
        doc.close();
        doc.open(file);
        REQUIRE(doc.property(XLProperty::SharedDoc) == "false");
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02T: Test SetProperty - HyperlinkBase")
    {
        doc.setProperty(XLProperty::HyperlinkBase, "HyperlinkBaseTest");
        doc.save();
        doc.close();
        doc.open(file);
        REQUIRE(doc.property(XLProperty::HyperlinkBase) == "HyperlinkBaseTest");
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02U: Test SetProperty - HyperlinksChanged")
    {
        doc.setProperty(XLProperty::HyperlinksChanged, "false");
        doc.save();
        doc.close();
        doc.open(file);
        REQUIRE(doc.property(XLProperty::HyperlinksChanged) == "false");
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02V: Test SetProperty - AppVersion")
    {
        doc.setProperty(XLProperty::AppVersion, "12.0300");
        doc.save();
        doc.close();
        doc.open(file);
        REQUIRE(doc.property(XLProperty::AppVersion) == "12.0300");
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("DeleteProperty - Keywords")
    {
        doc.deleteProperty(XLProperty::Keywords);
        doc.save();
        doc.close();
        doc.open(file);
        REQUIRE(doc.property(XLProperty::Keywords).empty());
    }

    doc.close();
}
//
// Created by Troldal on 2019-01-12.
//

#include "catch.hpp"
#include <OpenXLSX.hpp>
#include <fstream>

using namespace OpenXLSX;

/**
 * @brief The purpose of this test case is to test the creation of XLDocument objects. Each section section
 * tests document creation using a different method. In addition, saving, closing and copying is tested.
 */
TEST_CASE("C++ Interface Test 01: Creation of Excel Documents")
{
    std::string file    = "./TestDocumentCreation.xlsx";
    std::string newfile = "./TestDocumentCreationNew.xlsx";

    /**
     * @test Create new document using the CreateDocument method.
     *
     * @details Creates an empty document and creates the excel file using the CreateDocument() member function.
     * Success is tested by checking if the file have been created on disk and that the DocumentName member function
     * returns the correct file name.
     */
    SECTION("Section 01A: Create new using CreateDocument()")
    {
        XLDocument doc;
        doc.create(file);
        std::ifstream f(file);
        REQUIRE(f.good());
        REQUIRE(doc.name() == file);
    }

    /**
     * @brief Open an existing document using the constructor.
     *
     * @details Opens an existing document by passing the file name to the constructor.
     * Success is tested by checking that the DocumentName member function returns the correct file name.
     */
    SECTION("Section 01B: Open existing using Constructor")
    {
        XLDocument doc(file);
        REQUIRE(doc.name() == file);
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Section 01C: Open existing using openDocument()")
    {
        XLDocument doc;
        doc.open(file);
        REQUIRE(doc.name() == file);
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Section 01D: Save document using Save()")
    {
        XLDocument doc(file);

        doc.save();
        std::ifstream n(file);
        REQUIRE(n.good());
        REQUIRE(doc.name() == file);
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Section 01E: Save document using SaveDocumentAs()")
    {
        XLDocument doc(file);

        doc.saveAs(newfile);
        std::ifstream n(newfile);
        REQUIRE(n.good());
        REQUIRE(doc.name() == newfile);
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Section 01F: Copy construction")
    {
        XLDocument doc(file);
        XLDocument copy = doc;

        REQUIRE(copy.name() == doc.name());
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Section 01G: Copy assignment")
    {
        XLDocument doc(file);
        XLDocument copy;
        copy = doc;

        REQUIRE(copy.name() == doc.name());
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Section 01H: Move construction")
    {
        XLDocument doc(file);
        XLDocument copy = std::move(doc);

        REQUIRE(copy.name() == file);
        REQUIRE_THROWS(doc.name() == file);
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Section 01I: Move assignment")
    {
        XLDocument doc(file);
        XLDocument copy;
        copy = std::move(doc);

        REQUIRE(copy.name() == file);
        REQUIRE_THROWS(doc.name() == file);
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Section 01J: Close and Reopen")
    {
        XLDocument doc;
        doc.create(file);
        doc.close();
        REQUIRE_THROWS(doc.name() == file);

        doc.open(file);
        REQUIRE(doc.name() == file);
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Section 01K: Reopen without closing")
    {
        XLDocument doc;
        doc.create(file);

        doc.open(newfile);
        REQUIRE(doc.name() == newfile);
    }

    /**
     * @brief
     *
     * @details
     */
    SECTION("Section 01L: Open document as const")
    {
        const XLDocument doc(file);
        REQUIRE(doc.name() == file);
    }
}
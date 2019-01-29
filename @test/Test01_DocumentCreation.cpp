//
// Created by Troldal on 2019-01-12.
//

#include "catch.hpp"
#include <fstream>
#include <OpenXLSX.h>

using namespace OpenXLSX;

/**
 * @brief The purpose of this test case is to test the creation of XLDocument objects. Each section section
 * tests document creation using a different method.
 */
TEST_CASE("Test 01: Creation of Excel Documents") {

    std::string file = "./TestDocumentCreation.xlsx";
    std::string newfile = "./TestDocumentCreationNew.xlsx";

    /**
     * @test Create new document using the CreateDocument method.
     */
    SECTION("Section 01A: Create new using CreateDocument()") {
        XLDocument doc;
        doc.CreateDocument(file);
        std::ifstream f(file);
        REQUIRE(f.good() == true);
        REQUIRE(doc.DocumentName() == file);
    }

    SECTION("Section 01B: Open existing using Constructor") {
        XLDocument doc(file);
        REQUIRE(doc.DocumentName() == file);
    }

    SECTION("Section 01C: Open existing using OpenDocument()") {
        XLDocument doc;
        doc.OpenDocument(file);
        REQUIRE(doc.DocumentName() == file);
    }

    SECTION("Section 01D: Save document using Save()") {
        XLDocument doc(file);

        doc.SaveDocument();
        std::ifstream n(file);
        REQUIRE(n.good() == true);
        REQUIRE(doc.DocumentName() == file);
    }

    SECTION("Section 01D: Save document using SaveAs()") {
        XLDocument doc(file);

        doc.SaveDocumentAs(newfile);
        std::ifstream n(newfile);
        REQUIRE(n.good() == true);
        REQUIRE(doc.DocumentName() == newfile);
    }

}
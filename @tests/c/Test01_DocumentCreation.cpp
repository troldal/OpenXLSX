//
// Created by Troldal on 2019-01-12.
//

#include "catch.hpp"
#include <fstream>

extern "C" {
#include <XL_Document.h>
}

/**
 * @brief The purpose of this test case is to test the creation of XLDocument objects. Each section section
 * tests document creation using a different method. In addition, saving, closing and copying is tested.
 */
TEST_CASE("C Interface Test 01: Creation of Excel Documents") {

    std::string file    = "./TestDocumentCreation.xlsx";
    std::string newfile = "./TestDocumentCreationNew.xlsx";

    /**
     * @test Create new document using the CreateDocument method.
     *
     * @details Creates an empty document and creates the excel file using the CreateDocument() member function.
     * Success is tested by checking if the file have been created on disk and that the DocumentName member function
     * returns the correct file name.
     */
    SECTION("Section 01A: Create new using XL_CreateDocument()") {

        XL_Document* doc = XL_CreateDocument(file.c_str());
        std::ifstream f(file);
        REQUIRE(f.good() == true);
        REQUIRE(XL_DocumentName(doc) == file);

        XL_CloseDocument(doc);
        XL_DestroyDocument(doc);
    }

        /**
         * @brief
         */
    SECTION("Section 01B: Open existing using XL_OpenDocument()") {
        XL_Document* doc = XL_OpenDocument(file.c_str());
        REQUIRE(XL_DocumentName(doc) == file);
    }

        /**
         * @brief
         */
    SECTION("Section 01C: Save document using XL_SaveDocument()") {
        XL_Document* doc = XL_OpenDocument(file.c_str());
        XL_SaveDocument(doc);

        std::ifstream n(file);
        REQUIRE(n.good() == true);
        REQUIRE(XL_DocumentName(doc) == file);
    }

        /**
         * @brief
         */
    SECTION("Section 01D: Save document using XL_SaveDocumentAs()") {
        XL_Document* doc = XL_OpenDocument(file.c_str());
        XL_SaveDocumentAs(doc, newfile.c_str());

        std::ifstream n(newfile);
        REQUIRE(n.good() == true);
        REQUIRE(XL_DocumentName(doc) == newfile);
    }


        /**
         * @brief
         */
    SECTION("Section 01E: Close and Reopen") {
        XL_Document* doc = XL_CreateDocument(file.c_str());
        XL_CloseDocument(doc);
        XL_DestroyDocument(doc);

        doc = XL_OpenDocument(file.c_str());
        REQUIRE(XL_DocumentName(doc) == file);
    }
}
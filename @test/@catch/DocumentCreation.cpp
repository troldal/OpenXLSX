//
// Created by Troldal on 2019-01-12.
//

#include "catch.hpp"
#include <fstream>
#include <OpenXLSX.h>

using namespace OpenXLSX;

TEST_CASE( "Creation of Excel Documents", "[create]" ) {

    SECTION( "Create new document" ) {
        XLDocument doc;
        doc.CreateDocument("./MyTest.xlsx");
        std::ifstream f("./MyTest.xlsx");
        REQUIRE(f.good() == true);
        REQUIRE(doc.DocumentName() == "./MyTest.xlsx");
    }


    SECTION( "Open existing document" ) {
        XLDocument doc("./MyTest.xlsx");
        std::ifstream f("./MyTest.xlsx");
        REQUIRE(f.good() == true);
        REQUIRE(doc.DocumentName() == "./MyTest.xlsx");
    }

}
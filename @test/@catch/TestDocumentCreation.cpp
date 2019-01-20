//
// Created by Troldal on 2019-01-12.
//

#include "catch.hpp"
#include <fstream>
#include <OpenXLSX.h>

using namespace OpenXLSX;

TEST_CASE( "Creation of Excel Documents") {

    std::string file = "./TestDocumentCreation.xlsx";
    std::string newfile = "./TestDocumentCreationNew.xlsx";

    SECTION( "CreateDocument" ) {
        XLDocument doc;
        doc.CreateDocument(file);
        std::ifstream f(file);
        REQUIRE(f.good() == true);
        REQUIRE(doc.DocumentName() == file);
    }

    SECTION( "OpenDocument" ) {
        // Test opening via constructor
        XLDocument doc(file);
        REQUIRE(doc.DocumentName() == file);

        // Test opening via OpenDocument function
        doc.CloseDocument();
        doc.OpenDocument(file);
        REQUIRE(doc.DocumentName() == file);
    }

    SECTION( "Save" ) {
        XLDocument doc(file);

        doc.SaveDocument();
        doc.SaveDocumentAs(newfile);
        std::ifstream n(newfile);
        REQUIRE(n.good() == true);
        REQUIRE(doc.DocumentName() == newfile);
    }

}
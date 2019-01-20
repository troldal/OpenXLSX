//
// Created by Troldal on 2019-01-13.
//

#include "catch.hpp"
#include <fstream>
#include <OpenXLSX.h>

using namespace OpenXLSX;

TEST_CASE( "Testing of Document Properties", "[create]" ) {

    XLDocument doc;
    std::string file = "./TestDocumentProperties.xlsx";
    doc.OpenDocument(file);

    SECTION( "GetProperty" ) {
        REQUIRE(doc.GetProperty(XLProperty::Title).empty());
    }

    SECTION( "SetProperty - Title" ) {
        doc.SetProperty(XLProperty::Title, "TitleTest");
        doc.SaveDocument();
        doc.CloseDocument();
        doc.OpenDocument(file);
        REQUIRE(doc.GetProperty(XLProperty::Title) == "TitleTest");
    }

    // Duplication of above to check that OpenXLSX can handle
    // re-setting of already existing properties.
    SECTION( "SetProperty - Title (Duplicate)" ) {
        doc.SetProperty(XLProperty::Title, "TitleTest");
        doc.SaveDocument();
        doc.CloseDocument();
        doc.OpenDocument(file);
        REQUIRE(doc.GetProperty(XLProperty::Title) == "TitleTest");
    }

    SECTION( "SetProperty - Subject" ) {
        doc.SetProperty(XLProperty::Subject, "SubjectTest");
        doc.SaveDocument();
        doc.CloseDocument();
        doc.OpenDocument(file);
        REQUIRE(doc.GetProperty(XLProperty::Subject) == "SubjectTest");
    }

    SECTION( "SetProperty - Creator" ) {
        doc.SetProperty(XLProperty::Creator, "Kenneth");
        doc.SaveDocument();
        doc.CloseDocument();
        doc.OpenDocument(file);
        REQUIRE(doc.GetProperty(XLProperty::Creator) == "Kenneth");
    }

    SECTION( "SetProperty - Keywords" ) {
        doc.SetProperty(XLProperty::Keywords, "A, B, C");
        doc.SaveDocument();
        doc.CloseDocument();
        doc.OpenDocument(file);
        REQUIRE(doc.GetProperty(XLProperty::Keywords) == "A, B, C");
    }

    // Test other properties

    SECTION( "DeleteProperty - KEywords" ) {
        doc.DeleteProperty(XLProperty::Keywords);
        doc.SaveDocument();
        doc.CloseDocument();
        doc.OpenDocument(file);
        REQUIRE(doc.GetProperty(XLProperty::Keywords).empty());
    }

    doc.CloseDocument();

}
//
// Created by Troldal on 2019-01-13.
//

#include "catch.hpp"
#include <fstream>
#include <OpenXLSX.h>

using namespace OpenXLSX;

TEST_CASE( "Testing of Document Properties", "[create]" ) {

    XLDocument doc;
    std::ifstream f("./DocumentProperties.xlsx");
    if (!f.good())
        doc.CreateDocument("./DocumentProperties.xlsx");
    else
        doc.OpenDocument("./DocumentProperties.xlsx");


    SECTION( "Check empty property" ) {
        REQUIRE(doc.GetProperty(XLProperty::Title).empty());
    }

    SECTION( "Set 'Title' property" ) {
        doc.SetProperty(XLProperty::Title, "TitleTest");
        doc.SaveDocument();
        doc.CloseDocument();
        doc.OpenDocument("./DocumentProperties.xlsx");
        REQUIRE(doc.GetProperty(XLProperty::Title) == "TitleTest");
    }

    // Duplication of above to check that OpenXLSX can handle
    // re-setting of already existing properties.
    SECTION( "Re-Set 'Title' property" ) {
        doc.SetProperty(XLProperty::Title, "TitleTest");
        doc.SaveDocument();
        doc.CloseDocument();
        doc.OpenDocument("./DocumentProperties.xlsx");
        REQUIRE(doc.GetProperty(XLProperty::Title) == "TitleTest");
    }

    SECTION( "Set 'Subject' property" ) {
        doc.SetProperty(XLProperty::Subject, "SubjectTest");
        doc.SaveDocument();
        doc.CloseDocument();
        doc.OpenDocument("./DocumentProperties.xlsx");
        REQUIRE(doc.GetProperty(XLProperty::Subject) == "SubjectTest");
    }

    SECTION( "Set 'Creator' property" ) {
        doc.SetProperty(XLProperty::Creator, "Kenneth");
        doc.SaveDocument();
        doc.CloseDocument();
        doc.OpenDocument("./DocumentProperties.xlsx");
        REQUIRE(doc.GetProperty(XLProperty::Creator) == "Kenneth");
    }

    SECTION( "Set 'Keywords' property" ) {
        doc.SetProperty(XLProperty::Keywords, "A, B, C");
        doc.SaveDocument();
        doc.CloseDocument();
        doc.OpenDocument("./DocumentProperties.xlsx");
        REQUIRE(doc.GetProperty(XLProperty::Keywords) == "A, B, C");
    }

    // Test other properties

    //SECTION( "Set 'Keywords' property" ) {
    //    doc.DeleteProperty("Keywords");
    //    REQUIRE(doc.GetProperty(XLProperty::Keywords).empty());
    //}

    doc.CloseDocument();

}
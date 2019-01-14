//
// Created by Troldal on 2019-01-13.
//

#include "catch.hpp"
#include <fstream>
#include <OpenXLSX.h>

using namespace OpenXLSX;

TEST_CASE( "Setting of Document Properties", "[create]" ) {

    XLDocument doc;
    std::ifstream f("./DocumentProperties.xlsx");
    if (!f.good())
        doc.CreateDocument("./DocumentProperties.xlsx");
    else
        doc.OpenDocument("./DocumentProperties.xlsx");


    SECTION( "Empty property" ) {
        REQUIRE(doc.GetProperty(XLProperty::Title) == "");
    }

    SECTION( "Set 'Title' property" ) {
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

    doc.CloseDocument();

}
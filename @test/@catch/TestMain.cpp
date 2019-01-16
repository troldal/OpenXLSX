//
// Created by Troldal on 2019-01-10.
//

// this tells catch to provide a main()
// only do this in one cpp file

#ifndef OPENXLSX_TESTMAIN_H
#define OPENXLSX_TESTMAIN_H

#define CATCH_CONFIG_RUNNER
#include "catch.hpp"
#include <OpenXLSX.h>
#include <cstdio>

using namespace OpenXLSX;

int main( int argc, char* argv[] ) {
    // global setup...
    std::remove("./DocumentProperties.xlsx");
    std::remove("./WorkbookTests.xlsx");

    XLDocument doc;
    doc.CreateDocument("./WorkbookTests.xlsx");
    doc.SaveDocument();
    doc.CloseDocument();

    int result = Catch::Session().run( argc, argv );

    // global clean-up...

    return result;
}


#endif //OPENXLSX_TESTMAIN_H
//
// Created by Troldal on 2019-01-10.
//

// this tells catch to provide a main()
// only do this in one cpp file

#ifndef OPENXLSX_TESTMAIN_H
#define OPENXLSX_TESTMAIN_H

#define CATCH_CONFIG_RUNNER
#include "catch.hpp"
#include <cstdio>

extern "C" {
#include <XL_Document.h>
};


void PrepareDocument(std::string name) {

    std::remove(name.c_str());
    XL_Document* doc = XL_CreateDocument(name.c_str());
    XL_SaveDocument(doc);
    XL_CloseDocument(doc);
    XL_DestroyDocument(doc);
}

int main( int argc, char* argv[] ) {
    // Global Setup
    
    std::remove("./TestDocumentCreation.xlsx");
    std::remove("./TestDocumentCreationNew.xlsx");

    PrepareDocument("./TestDocumentProperties.xlsx");
    PrepareDocument("./TestWorkbook.xlsx");
    PrepareDocument("./TestSheet.xlsx");
    PrepareDocument("./TestWorksheet.xlsx");
    PrepareDocument("./TestCellReference.xlsx");
    PrepareDocument("./TestCell.xlsx");
    // Global Setup Complete

    // Run Test Suite
    int result = Catch::Session().run( argc, argv );
    // Run Test Suite Complete

    // Global Clean-up
    // ...
    // Global Setup Complete

    return result;
}


#endif //OPENXLSX_TESTMAIN_H
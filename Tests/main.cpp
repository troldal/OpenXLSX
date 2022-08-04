//
// Created by Troldal on 2019-01-10.
//

// this tells catch to provide a main()
// only do this in one cpp file

#ifndef OPENXLSX_TESTMAIN_H
#define OPENXLSX_TESTMAIN_H

#define CATCH_CONFIG_RUNNER

#include <catch.hpp>
#include <filesystem>
#include <OpenXLSX.hpp>
#include <cstdio>

using namespace OpenXLSX;

void PrepareDocument(const std::string& name)
{
    // delete the document/file before creating it:
    std::filesystem::path fileName{name};
    std::filesystem::remove(fileName);
    // create the document:
    XLDocument doc;
    doc.create(fileName);
    // save the document to the given file;
    doc.save();
    // close the document:
    doc.close();
}

int main(int argc, char* argv[])
{
    // Global Setup
    //    XLDocument doc;
    //
    //    std::remove("./TestDocumentCreation.xlsx");
    //    std::remove("./TestDocumentCreationNew.xlsx");
    //
    //    PrepareDocument("./TestDocumentProperties.xlsx");
    //    PrepareDocument("./TestWorkbook.xlsx");
    //    PrepareDocument("./TestSheet.xlsx");
    //    PrepareDocument("./TestWorksheet.xlsx");
    //    PrepareDocument("./TestCellReference.xlsx");
    //    PrepareDocument("./TestCell.xlsx");
    // Global Setup Complete

    // Run Test Suite
    int result = Catch::Session().run(argc, argv);
    // Run Test Suite Complete

    // Global Clean-up
    // ...
    // Global Setup Complete

    return result;
}

#endif    // OPENXLSX_TESTMAIN_H
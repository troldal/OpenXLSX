# OpenXLSX
OpenXLSX is a C++ library for reading, writing, creating and modifying Microsoft Excel® files, with the .xlsx format.

## Motivation
Many programming languages have the ability to modify Excel files, either natively or in the form of open source libraries. This includes Python, Java and C#. For C++, however, things are more scattered. While there are some libraries, they are generally less mature and have a smaller feature set than for other languages. Here is a summary of the main C++ libraries for Excel files:

### libxls
The libxls library (https://sourceforge.net/projects/libxls/) is a C library for reading files in the legacy Excel file format, .xls. It cannot be used for writing or modifying Excel files.

### xlslib
The xlslib library (https://sourceforge.net/projects/xlslib/) is a C/C++ library for creating files in the legacy Excel file format, .xls. It cannot be used for reading or modifying Excel files.

### libxlsxwriter
The libxlsxwriter library (https://libxlsxwriter.github.io) is a C library for creating .xlsx files. It cannot be used for reading or modifying Excel files.

### LibXL
The LibXL library (http://www.libxl.com) can read, write, create and modify Excel files, in both the .xls and .xlsx formats. It is the most feature complete library available and has interfaces for C, C++, C# and Delphi. It is only available for purchase, however.

### QtXlsx
Of the open source libraries, the QtXlsx library (https://github.com/dbzhang800/QtXlsxWriter) is the most feature complete. It is, however, based on the Qt framework. While I'm a big fan of Qt for application programming purposes, I don't believe it is the best option for lower-level libraries.

Because there are no open source library that fully fitted my needs, I decided to develop the OpenXLSX library.

## Ambition
The ambition is that OpenXLSX should be able to read, write, create and modify Excel files (data as well as formatting), and do so with as few dependencies as possible. Currently, OpenXLSX depends on the following 3rd party libraries (all included in the repository):

 - PugiXML
 - LibZip
 - LibZip++
 
 ## Current Status
 OpenXLSX is still work i progress. The following is a list of features which have been implemented and should be working properly:
 
  - Create/open/save files
  - Read/write/modify cell contents
  - Copy cells and cell ranges
  - Copy worksheets
  
  Features related to formatting, plots and figures have not yet been implemented
  
  It should be noted, that creating const XLDocument objects, is currently not working!
  
  ## Compatibility
  OpenXLSX can be built and run on the following platforms/compilers:
  
  - Linux (GCC)
  - MacOS (GCC/LLVM/Xcode)
  - Windows (MinGW)
  
  OpenXLSX can be compiled on Windows using MSVC, but there is an issue with the CMake script, which prevents MSVC finding the libraries.
  Building on Windows using MSYS has not been tested.
  Building using the Intel compiler has not been tested.
  
  ## Usage
  
  ```cpp
  #include <iostream>
  #include <OpenXLSX/OpenXLSX.h>
  
  using namespace std;
  using namespace OpenXLSX;
  
  int main() {
  
    XLDocument doc;
    doc.CreateDocument("MyTest.xlsx");
    auto wks = doc.Workbook()->Worksheet("Sheet1");

    *wks->Cell("A1")->Value() = 3.14159;
    *wks->Cell("B1")->Value() = 42;
    *wks->Cell("C1")->Value() = "Hello OpenXLSX!";
    *wks->Cell("D1")->Value() = true;

    auto A1 = wks->Cell("A1")->Value()->Get<double>();
    auto B1 = wks->Cell("B1")->Value()->Get<int>();
    auto C1 = wks->Cell("C1")->Value()->Get<std::string>();
    auto D1 = wks->Cell("D1")->Value()->Get<bool>();

    cout << "Cell A1: " << A1 << endl;
    cout << "Cell B1: " << B1 << endl;
    cout << "Cell C1: " << C1 << endl;
    cout << "Cell D1: " << D1 << endl;

    doc.SaveDocument();
  
  return 0;
  }
  ```  

# OpenXLSX
OpenXLSX is a C++ library for reading, writing, creating and modyfing Microsoft Excel® files, with the .xlsx format.

## Motivation
Many programming languages have the ability to modify Excel files, either natively or in the form of open source libraries. This includes Python, Java and C#. For C++, however, things are more scattered. While there are some libraries, they are generally less mature and have a smaller featureset than for other languages. Here is a summary of the main C++ libraries for Excel files:

### libxls
The libxls library (https://sourceforge.net/projects/libxls/) is a C library for reading files in the legacy Excel file format, .xls. It cannot be used for writing or modifying Excel files.

### xlsib
The xlslib library (https://sourceforge.net/projects/xlslib/) is a C/C++ library for creating files in the legacy Excel file format, .xls. It cannot be used for reading or modifying Excel files.

### libxlsxwriter
The libxlsxwriter library (https://libxlsxwriter.github.io) is a C library for creating .xlsx files. It cannot be used for reading or modifying Excel files.

### LibXL
The LibXL library (http://www.libxl.com) can read, write, create and modify Excel files, in both the .xls and .xlsx formats. It is the most feature complete library available and has interfaces for C, C++, C# and Delphi. It is only available for purchase, however.

### QtXlsx
Of the open source libraries, the QtXlsx library (https://github.com/dbzhang800/QtXlsxWriter) is the most feature complete. It is, however, based on the Qt framework. While I'm a big fan of Qt for application programming purposes, I don't believe it is the best option for lower-level libraries.

Because there are no open source library that fully fitted my needs, I decided to develop the OpenXLSX library.

## Ambition
The ambition is that OpenXLSX should be able to read, write, create and modify Excel files (data as well as formatting), and do so with as few dependencies as possible. Currently, OpenXLSX depends on the following 3rd party libraries:

 - Boost
 - RapidXML (included)
 - LibZip (and, in turn, zlib)
 - LibZip++ (included)
 
 ## Current Status
 OpenXLSX is still work i progress. The following is a list of features which have been implemented and should be working properly:
 
  - Create/open/save files
  - Read/write/modify cell contents
  - Copy cells and cell ranges
  - Copy worksheets
  
  Features related to formatting, plots and figures have not yet been implemented
  
  It should be noted, that creating const XLDocument objects, is currently not working!
  
  ## Usage
  ```cpp
  
  #include <iostream>
  #include <OpenXLSX/OpenXLSX.h>
  
  using namespace std;
  using namespace OpenXLSX;
  
  int main() {
  
      XLDocument doc;
  
      doc.CreateDocument("Spreadsheet.xlsx");
      doc.Workbook().AddWorksheet("MyWorksheet");
      auto &wks = doc.Workbook().Worksheet("MyWorksheet");
  
      auto arange = wks.Range(XLCellReference(1,1), XLCellReference(10,10));
      for (auto &iter : arange) {
          iter.Value().Set("Hello OpenXLSX!");
      }
  
      doc.SaveDocument();
      doc.CloseDocument();
  
      XLDocument rdoc;
      rdoc.OpenDocument("Spreadsheet.xlsx");
  
      cout << "Content of cell B2: " << rdoc.Workbook().Worksheet("MyWorksheet").Cell("B2").Value().AsString();
  
      rdoc.CloseDocument();
  
  return 0;
  }
  
  ```  
  Note that it is (currently) required to link the program to Boost system Boost filesystem and LibZip libraries, in order to run.

# OpenXLSX
OpenXLSX is a C++ library for reading, writing, creating and modifying Microsoft Excel® files, with the .xlsx format.

## New in version 0.2.0
The internal architecture of OpenXLSX has been significantly re-designed since the previous version. The reason is that the library was turning into a big ball of mud, and it was increasingly difficult to add features and fix bugs. With the new architecture, it will (hopefully) be easier to manage and to add new features.

Due to the re-design of the architecture, there are a few changes to the public interface. These changes, however, are not significant, and it should be easy to update:

* All internal objects are now handled as values rather than pointers, similar to the approach taken in the underlying PugiXML library. This means that when requesting a certain worksheet from a workbook, the resulting worksheet is not returned as a pointer, but as an object that supports both copying and moving.
* The distingction between interface objects and implementation objects are now gone, as it made it very difficult to manage changes. It was an attempt to implement the pimpl idiom, but it wasn't very effective. In the future, I may try to implement pimpl again, but only if it can be done in a simpler way.
* All member functions have been renamed to begin wtih a small letter (camelCase), i.e. the member function WorksheetCount() is renamed to worksheetCount(). This was done mostly for cosmetic reasons.

I realise that these changes may cause problems for some users. Because of that, the previous version of OpenXLSX can be found in the "legacy" branch of this repository. However, I strongly recommend that you transition to the new version instead.

## Motivation
Many programming languages have the ability to modify Excel files, either natively or in the form of open source libraries. This includes Python, Java and C#. For C++, however, things are more scattered. While there are some libraries, they are generally less mature and have a smaller feature set than for other languages. 

Because there are no open source library that fully fitted my needs, I decided to develop the OpenXLSX library.

Here is a summary of the main C++ libraries for Excel files that I'm aware of:

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

### XLNT
Recently, I found the XLNT library on GitHub (https://github.com/tfussell/xlnt). It was not available when I began developing OpenXLSX. To be honest, if it had, I wouldn't have begun OpenXLSX. It has a larger feature set and probably has fewer bugs. However, I decided to continue developing OpenXLSX, because I believe that in a few areas it is better than XLNT. Primarily, OpenXLSX is better able to handle very large spreadsheets (up to a million rows). 

## Ambition
The ambition is that OpenXLSX should be able to read, write, create and modify Excel files (data as well as formatting), and do so with as few dependencies as possible. Currently, OpenXLSX depends on the following 3rd party libraries:

 - PugiXML
 - Zippy (C++ wrapper around miniz)

these libraries are included in the repository, i.e. it's not necessary to download and build separately.

## Compatibility
  OpenXLSX can be built and run on the following platforms/compilers:

| Platform\Compiler | GCC | Clang | MSVC |
|:------------------|:----|:------|:-----|
| Windows           |     |       |      |
| MacOS             |     |       |      |
| Linux             |     |       |      |


  - Linux (GCC/Clang)
  - MacOS (GCC/Clang/Xcode)
  - Windows (MinGW/Clang/MSVC)

  Building on Windows using MSYS or Cygwin has not been tested.
  Building using the Intel compiler has not been tested.

## Current Status
 OpenXLSX is still work i progress. The following is a list of features which have been implemented and should be working properly:

  - Create/open/save files
  - Read/write/modify cell contents
  - Copy cells and cell ranges
  - Copy worksheets

  Features related to formatting, plots and figures have not been implemented, and are not planned to be in the near future.

  It should be noted, that creating const XLDocument objects, is currently not working!

### Performance


### Caveats



## Unicode
  All string manipulations and usage in OpenXLSX uses the C++ std::string, which uses UTF-8 encoding. Also, Excel uses UTF-8 encoding internally (actually, it might be possible to use other encodings, but I'm not sure about that). 
  
  For the above reason, if you work with other text encodings, you have to convert to/from UTF-8 yourself. There are a number of options (e.g. std::codecvt or Boost.Locale). I will also suggest that you see James McNellis' presentation at [CppCon 2014](https://youtu.be/n0GK-9f4dl8).
  
## Example Programs


### Basic Usage
  
  ```cpp
#include <OpenXLSX.hpp>
#include <iomanip>
#include <iostream>

using namespace std;
using namespace OpenXLSX;

int main()
{
    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #01: Basic Usage\n";
    cout << "********************************************************************************\n";

    XLDocument doc;
    doc.create("./Demo01.xlsx");
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.cell(XLCellReference("A1")).value() = 3.14159;
    wks.cell(XLCellReference("B1")).value() = 42;
    wks.cell(XLCellReference("C1")).value() = "  Hello OpenXLSX!  ";
    wks.cell(XLCellReference("D1")).value() = true;
    wks.cell(XLCellReference("E1")).value() = wks.cell(XLCellReference("C1")).value();

    auto A1 = wks.cell(XLCellReference("A1")).value();
    auto B1 = wks.cell(XLCellReference("B1")).value();
    auto C1 = wks.cell(XLCellReference("C1")).value();
    auto D1 = wks.cell(XLCellReference("D1")).value();
    auto E1 = wks.cell(XLCellReference("E1")).value();

    auto valueTypeAsString = [](OpenXLSX::XLValueType type) {
        switch (type) {
            case XLValueType::String:
                return "string";

            case XLValueType::Boolean:
                return "boolean";

            case XLValueType::Empty:
                return "empty";

            case XLValueType::Error:
                return "error";

            case XLValueType::Float:
                return "float";

            case XLValueType::Integer:
                return "integer";
        }
    };

    cout << "Cell A1: (" << valueTypeAsString(A1.valueType()) << ") " << A1.get<double>() << endl;
    cout << "Cell B1: (" << valueTypeAsString(B1.valueType()) << ") " << B1.get<int>() << endl;
    cout << "Cell C1: (" << valueTypeAsString(C1.valueType()) << ") " << C1.get<std::string>() << endl;
    cout << "Cell D1: (" << valueTypeAsString(D1.valueType()) << ") " << D1.get<bool>() << endl;
    cout << "Cell E1: (" << valueTypeAsString(E1.valueType()) << ") " << E1.get<std::string_view>() << endl;

    doc.save();

    return 0;
}
  ```  

## Credits

- Thanks to [@zeux](https://github.com/zeux) for providing the awesome PugiXML library.
- Thanks to [@richgel999](https://github.com/richgel999) for providing the miniz library.
- Thanks to [@tectu](https://github.com/Tectu) for cleaning up my CMakeLists.txt files.
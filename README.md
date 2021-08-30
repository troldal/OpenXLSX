# OpenXLSX

OpenXLSX is a C++ library for reading, writing, creating and modifying
Microsoft Excel® files, with the .xlsx format.

## Table of Contents

- [Motivation](#motivation)
- [Ambition](#ambition)
- [Compatibility](#compatibility)
- [Build Instructions](#build-instructions)
- [Current Status](#current-status)
- [Performance](#performance)
- [Caveats](#caveats)
  - [File Size](#file-size)
  - [Memory Usage](#memory-usage)
  - [Unicode](#unicode)
- [Reference Guide](#reference-guide)
- [Example Programs](#example-programs)
- [Changes](#changes)

## Motivation

Many programming languages have the ability to modify Excel files,
either natively or in the form of open source libraries. This includes
Python, Java and C#. For C++, however, things are more scattered. While
there are some libraries, they are generally less mature and have a
smaller feature set than for other languages.

Because there are no open source library that fully fitted my needs, I
decided to develop the OpenXLSX library.

Here is a summary of the main C++ libraries for Excel files that I'm
aware of:

### libxls

The libxls library (https://sourceforge.net/projects/libxls/) is a C
library for reading files in the legacy Excel file format, .xls. It
cannot be used for writing or modifying Excel files.

### xlslib

The xlslib library (https://sourceforge.net/projects/xlslib/) is a C/C++
library for creating files in the legacy Excel file format, .xls. It
cannot be used for reading or modifying Excel files.

### libxlsxwriter

The libxlsxwriter library (https://libxlsxwriter.github.io) is a C
library for creating .xlsx files. It cannot be used for reading or
modifying Excel files.

### LibXL

The LibXL library (http://www.libxl.com) can read, write, create and
modify Excel files, in both the .xls and .xlsx formats. It is the most
feature complete library available and has interfaces for C, C++, C# and
Delphi. It is only available for purchase, however.

### QtXlsx

Of the open source libraries, the QtXlsx library
(https://github.com/dbzhang800/QtXlsxWriter) is the most feature
complete. It is, however, based on the Qt framework. While I'm a big fan
of Qt for application programming purposes, I don't believe it is the
best option for lower-level libraries.

### XLNT

Recently, I found the XLNT library on GitHub
(https://github.com/tfussell/xlnt). It was not available when I began
developing OpenXLSX. To be honest, if it had, I wouldn't have begun
OpenXLSX. It has a larger feature set and probably has fewer bugs.
However, I decided to continue developing OpenXLSX, because I believe
that in a few areas it is better than XLNT. Primarily, OpenXLSX is
better able to handle very large spreadsheets (up to a million rows).

## Ambition

The ambition is that OpenXLSX should be able to read, write, create and
modify Excel files (data as well as formatting), and do so with as few
dependencies as possible. Currently, OpenXLSX depends on the following
3rd party libraries:

- PugiXML
- Zippy (C++ wrapper around miniz)
- Boost.Nowide (for opening files with non-ASCII names on Windows)

these libraries are included in the repository, i.e. it's not necessary
to download and build separately.

## Compatibility

OpenXLSX has been tested on the following platforms/compilers. Note that
a '-' doesn't mean that OpenXLSX doesn't work; it just means that it
hasn't been tested:

|         | GCC | Clang | MSVC |
|:--------|:----|:------|:-----|
| Windows | N/A | -     | +    |
| MinGW   | +   | +     | N/A  |
| MSYS    | +   | N/A   | N/A  |
| Cygwin  | -   | -     | N/A  |
| MacOS   | +   | +     | N/A  |
| Linux   | +   | +     | N/A  |

The following compiler versions should be able to compile OpenXLSX
without errors:

- GCC: Version 7
- Clang: Version 8
- MSVC: Visual Studio 2019

Clang 7 should be able to compile OpenXLSX, but apparently there is a
bug in the implementation of std::variant, which causes compiler errors.

Visual Studio 2017 should also work, but hasn't been tested.

## Build Instructions

OpenXLSX uses CMake as the build system. If you use GNU makefiles (e.g.
on Linux, MacOS, MinGW or MSYS), you can build OpenXLSX from the
commandline by navigating to the root of the OpenXLSX repository, and
then execute the following commands:

```
mkdir build
cd build
cmake ..
make
```

Depending on your system, you may need to supply cmake with additional
commands. Also, if you use MinGW makefiles, you may need to run
'mingw32-make' instead of 'make'.

If you use an IDE, such as Visual Studio or Xcode, you can supply cmake
with a -G flag, followed by which build system you need. See CMake
documentation for details.

If you wish, you can also use the CMake GUI application.

## Python Interface

A Python interface has been developed for OpenXLSX. You can find
instructions in the Python folder.

## Current Status

OpenXLSX is still work i progress. The following is a list of features
which have been implemented and should be working properly:

- Create/open/save files
- Read/write/modify cell contents
- Copy cells and cell ranges
- Copy worksheets
- Ranges and Iterators

Features related to formatting, plots and figures have not been
implemented, and are not planned to be in the near future.

It should be noted, that creating const XLDocument objects, is currently
not working!

## Performance

The table below is the output from a benchmark (using the Google
Benchmark library), which shows that read/write access can be done at a
rate of around 2,000,000 cells per second. Floating point numbers are
somewhat lower, due to conversion to/from strings in the .xml file.

```
Run on (16 X 2300 MHz CPU s)
CPU Caches:
  L1 Data 32 KiB (x8)
  L1 Instruction 32 KiB (x8)
  L2 Unified 256 KiB (x8)
  L3 Unified 16384 KiB (x1)
Load Average: 2.46, 2.25, 2.19
---------------------------------------------------------------------------
Benchmark                 Time             CPU   Iterations UserCounters...
---------------------------------------------------------------------------
BM_WriteStrings        2484 ms         2482 ms            1 items=8.38861M items_per_second=3.37956M/s
BM_WriteIntegers       1949 ms         1949 ms            1 items=8.38861M items_per_second=4.30485M/s
BM_WriteFloats         4720 ms         4719 ms            1 items=8.38861M items_per_second=1.77767M/s
BM_WriteBools          2167 ms         2166 ms            1 items=8.38861M items_per_second=3.87247M/s
BM_ReadStrings         1883 ms         1882 ms            1 items=8.38861M items_per_second=4.45776M/s
BM_ReadIntegers        1641 ms         1641 ms            1 items=8.38861M items_per_second=5.11252M/s
BM_ReadFloats          4173 ms         4172 ms            1 items=8.38861M items_per_second=2.01078M/s
BM_ReadBools           1898 ms         1898 ms            1 items=8.38861M items_per_second=4.4205M/s
```

## Caveats

### File Size

An .xlsx file is essentially a bunch of .xml files in a .zip archive.
Internally, OpenXLSX uses the miniz library to compress/decompress the
.zip archive, and it turns out that miniz has an upper limit regarding
the file sizes it can handle.

The maximum allowable file size for a file in an archive (i.e. en entry
in a .zip archive, not the archive itself) is 4 GB (uncompressed).
Usually, the largest file in an .xlsx file/archive, will be the .xml
files holding the worksheet data. I.e., the worksheet data may not
exceed 4 GB. What that translates to in terms of rows and columns depend
a lot on the type of data, but 1,048,576 rows x 128 columns filled with
4-digit integers will take up approx. 4 GB. The size of the compressed
archive also depends on the data held in the worksheet, as well as the
compression algorithm used, but a workbook with a single worksheet of 4
GB will usually have a compressed size of 300-350 MB.

The 4 GB limitation is only related to a single entry in an archive, not
the total archive size. That means that if an archive holds multiple
entries with a size of 4GB, miniz can still handle it. For OpenXLSX,
this means that a workbook with several large worksheets can still be
opened.

### Memory Usage

OpenXLSX uses the PugiXML library for parsing and manipulating .xml
files in .xlsx archive. PugiXML is a DOM parser, which reads the entire
.xml document into memory. That makes parsing and manipulation
incredibly fast.

However, all choices have consequences, and using a DOM parser can also
demand a lot of memory. For small spreadsheets, it shouldn't be a
problem, but if you need to manipulate large spreadsheets, you may need
a lot of memory.

The table below gives an indication of how many columns of data can be
handled by OpenXLSX (assuming 1,048,576 rows):

|           | Columns |
|:----------|:--------|
| 8 GB RAM  | 8-16    |
| 16 GB RAM | 32-64   |
| 32 GB RAM | 128-256 |

Your milage may vary. The performance of OpenXLSX will depend on the
type of data in the spreadsheet.

Note also that it is recommended to use OpenXLSX in 64-bit mode. While
it can easily be used in 32-bit mode, it can only access 4 GB of RAM,
which will severely limit the usefulness when handling large
spreadsheets.

If memory consumption is an issue for you, you can build the OpenXLSX
library in compact mode (look for the ENABLE_COMPACT_MODE in the
CMakeLists.txt file), which will enable PugiXML's compact mode. OpenXLSX
will then use less memory, but also run slower. See further details in
the PugiXML documentation
[here](https://pugixml.org/docs/manual.html#dom.memory.compact). A test
case run on Linux VM with 8 GB RAM revealed that OpenXLSX could handle a
worksheet with 1,048,576 rows x 32 columns in compact mode, versus
1,048,576 rows x 16 columns in default mode.

### Unicode

By far the most questions I get about OpenXLSX on Github, is related to Unicode.
It is apparently (and understandably) a source of great confusion for many people.

Early on, I decided that OpenXLSX should focus on the Excel part, and not be a 
text encoding/conversion utility also. Therefore, **all text input/output to 
OpenXLSX MUST be in UTF-8 encoding**... Otherwise it won't work as expected. 
It may also be necessary that the source code files are saved in UTF-8 format. 
If, for example, a source file is saved in UTF-16 format, then any string literals
will also be in UTF-16. So if you have any hard-coded string literals in your source
code, then the source file **must** also be saved in UTF-8 format.

All string manipulations and usage in OpenXLSX uses the C++ std::string,
which is encoding agnostic, but can easily be used for UTF-8 encoding.
Also, Excel uses UTF-8 encoding internally (actually, it might be
possible to use other encodings, but I'm not sure about that).

For the above reason, **if you work with other text encodings, you have
to convert to/from UTF-8 yourself**. There are a number of options (e.g.
Boost.Nowide or Boost.Text). Internally, OpenXLSX uses Boose.Nowide; it has a number
of handy features for opening files and converting between std::string and std::wstring
etc. I will also suggest that you watch James McNellis' presentation at [CppCon 2014](https://youtu.be/n0GK-9f4dl8),
and read [Joel Spolsky's blog](https://www.joelonsoftware.com/2003/10/08/the-absolute-minimum-every-software-developer-absolutely-positively-must-know-about-unicode-and-character-sets-no-excuses/).

Unicode on Windows is particularly challenging. While UTF-8 is well 
supported on Linux and MacOS, support on Windows is more limited. For example, 
output of non-ASCII characters (e.g. Chinese or Japanese characters) to the 
terminal window will look like gibberish. As mentioned, sometimes you also have to be mindful 
of the text encoding of the source files themselves. Some users have had problems
with OpenXLSX crashing when opening/creating .xlsx files with non-ASCII 
filenames, where it turned out that the source code for the test program
was in a non-UTF-8 encoding, and hence the input string to OpenXLSX was also
non-UTF-8. To stay sane, I recommend that source code files are always 
in UTF-8 files; all IDE's I know of can handle source code files in UTF-8 
encoding. Welcome to the wonderful world of unicode on Windows 🤮

## Reference Guide
This section will contain the user reference for usage of OpenXLSX.
HOLD

### XLDocument class
HOLD

### XLWorkbook class
HOLD

### XLWorksheet class
HOLD

### XLCellReference class
The most common way of addressing cells in an Excel spreadsheet is by using the A1 notation, i.e. the columns are 
represented by letters and the rows are represented by numbers. However, it is also possible to use the so-called 
R1C1 notation, where both rows and columns are represented by numbers.

The XLCellReference class encapsulates the concept of a cell address, and can be used to convert between different 
notations. It is also used as an input parameter on XLWorksheet objects to retrieve specific cells.

The most common way of creating an XLCellReference object is to call the constructor with a string with the cell 
address in A1 notation, like this:

```cpp
auto cellRef = XLCellReference("B2");
```
It is also possible to create an XLCellReference object by providing the row and column coordinates of the cell, 
like this:
```cpp
auto cellRef = XLCellReference(2,2);
```

If the default constructor is called (i.e. no input parameters), the resulting object will refer to cell A1.

If you have an XLCellReference object, you can get the row/column number or the address using the `row()`, `column()
`, or `address()` member functions. Similarly, the object can be modified using the `setRow()`, `setColumn()
`, or `setAddress()` member functions:

```cpp
uint16_t row = cellRef.row();            // Get the row number
uint32_t column = cellRef.column();      // Get the column number
std::string address = cellRef.address(); // Get the cell address

cellRef.setRow(3);        // Set the row number;
cellRef.setColumn(5);     // Set the column number;
cellRef.setAddress("E3"); // Set the cell address;
```

### XLCell class
An object of the XLCell class represents a cell in a spreadsheet. The XLCell class itself doesn't do much, as it can 
mostly be considered a container of various cell properties, such as the cell value, cell formulas (not yet 
implemented), and cell formatting (not yet implemented).

An XLCell object can be retrieved via the `cell()` member function of the XLWorksheet class. This method takes an 
XLCellReference object, or alternatively a string with the cell address, as a parameter.

### XLCellValue and XLCellValueProxy classes
As mentioned previously, an object of the XLCell class doesn't contain the actual cell value. Instead, the value is 
contained in an XLCellValue object. XLCellValue is essentially a wrapper around a std::variant object, with some 
handy functionality for conversion and assignment.

An XLCellValue object has no knowledge about the cell it belongs to; it is simply a value that can passed around or 
stored in a container. Hence, assigning a new value to an XLCellValue object does **not** change the value of any 
cells in the spreadsheet.

The XLCellValue class has a templated constructor that can take any integer value (including booleans), any floating 
point value, and any string type with a `c_str()` method for getting a c-string. It is also possible to construct an 
XLCellValue object from raw c-strings and from std::string_view objects, although it doesn't have a `c_str()` member 
function. The copy assignment operator works in the same manner.

For conversion in the other direction, the XLCellValue class provides a templated `get()` function, for converting 
to a basic type. Similar to construction, the template parameter can be any integer type, and floating point type, 
and any type that can be constructed from c-strings. For example:

```c++
auto val1 = cellValue.get<int>();
auto val2 = cellValue.get<double>();
auto val3 = cellValue.get<bool>();
auto val4 = cellValue.get<std::string>();
```

Note, however, that if the value contained in the XLCellValue object cannot be converted to the template parameter 
type (e.g. if the XLCellValue object contains a string, and the template parameter is int), then an XLValueTypeError 
will be thrown.

In addition to the `get()` function, the XLCellValue class has a templated explicit conversion operator, that can 
perform explicit conversion, using the the same rules as above. For example:

```cpp
auto val1 = static_cast<int>(cellValue);
auto val2 = double(cellValue);
```

The XLCell class has a `value()` member function, and one would think that it returns the XLCellValue object for 
that cell. And it does...sort of. The `value()` member function actually returns a reference to an XLCellValueProxy 
object. This objects serves as a placeholder that facilitates conversion to an XLCellValue object, as well as 
assignment of a value to a cell. Whereas assigning a new value to an XLCellValue object does not modify the 
spreadsheet in any way, assigning a value to a XLCellValueProxy object, will.

This is probably better illustrated by example:

```cpp
XLCellValue val = cell.value();
```
While the `value()` member function returns a reference to a XLCellValueProxy object, this will be implicitly 
converted to an XLCellValue object. In addition, the XLCellValueProxy class supports the same conversion options as 
XLCellValue.

But the XLCellValueProxy class also supports assignment operations, that will modify the cell value in the 
underlying spreadsheet. Assignment can be done with XLCellValue objects, any integer type, booleans, any floating 
point types, c-strings, std::string_view and any string type with a `c_str()` member function. For example:

```cpp
cell.value() = 42;
cell.value() = "Hello, OpenXLSX!";
cell.value() = 3.14159;
cell.value() = false;
cell.value() = XLCellValue(123);
```

The benefit of this approach is, that we can have XLCellValue objects that uses value semantics, i.e. contains no 
pointers to the underlying spreadsheet, and that can be copied around, while at the same time can use the same `value
()` member function to assign a new value to the underlying XLCell object. 

This approach works seamlessly...almost. There is one drawback, and that is that it is not possible to use the auto 
keyword:
```cpp
auto val = cell.value(); // Error!
```
Using the auto keyword would invoke the copy constructor of the XLCellValueProxy object, which is undesired. 
Therefore the copy constructor has been made private, and the code above will therefore not compile. Using auto& 
instead would compile, but would not be of any use.

### XLRow class
HOLD

### XLRowData and XLRowDataProxy classes
HOLD

## Example Programs

The following example programs illustrates the key features of OpenXLSX.
The source code is included in the 'examples' directory in the OpenXLSX
repository.

### Basic Usage
This example program illustrates simple usage, such as assigning values to spreadsheet cells, and retrieving cell 
values as XLCellValue objects.
```cpp
#include <OpenXLSX.hpp>
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

    // Here an address string is used in the .cell member function, but an XLCellReference object can
    // also be used. For example, for the cell at 'A1', XLCellReference{1,1} can be usedas an argument.
    wks.cell("A1").value().set(3.14159);
    wks.cell("B1").value() = 42;
    wks.cell("C1").value() = "  Hello OpenXLSX!  ";
    wks.cell("D1").value() = true;
    wks.cell("E1").value() = wks.cell(XLCellReference("C1")).value();

    XLCellValue A1 = wks.cell("A1").value();
    XLCellValue B1 = wks.cell("B1").value();
    XLCellValue C1 = wks.cell("C1").value();
    XLCellValue D1 = wks.cell("D1").value();
    XLCellValue E1 = wks.cell("E1").value();

    cout << "Cell A1: (" << A1.typeAsString() << ") " << A1.get<double>() << endl;
    cout << "Cell B1: (" << B1.typeAsString() << ") " << B1.get<int64_t>() << endl;
    cout << "Cell C1: (" << C1.typeAsString() << ") " << C1.get<std::string>() << endl;
    cout << "Cell D1: (" << D1.typeAsString() << ") " << D1.get<bool>() << endl;
    cout << "Cell E1: (" << E1.typeAsString() << ") " << E1.get<std::string_view>() << endl << endl;
    
    doc.save();

    return 0;
}
```

### Sheet Handling

```cpp
#include <OpenXLSX.hpp>
#include <iostream>

using namespace std;
using namespace OpenXLSX;

int main()
{
    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #02: Sheet Handling\n";
    cout << "********************************************************************************\n";

    XLDocument doc;
    doc.create("./Demo02.xlsx");
    auto wbk = doc.workbook();

    cout << "\nSheets in workbook:\n";
    for (const auto& name : wbk.worksheetNames()) cout << wbk.indexOfSheet(name) << " : " << name << "\n";

    cout << "\nAdding new sheet 'MySheet01'\n";
    wbk.addWorksheet("MySheet01");

    cout << "Adding new sheet 'MySheet02'\n";
    wbk.addWorksheet("MySheet02");

    cout << "Cloning sheet 'Sheet1' to new sheet 'MySheet03'\n";
    wbk.sheet("Sheet1").get<XLWorksheet>().clone("MySheet03");

    cout << "Cloning sheet 'MySheet01' to new sheet 'MySheet04'\n";
    wbk.cloneSheet("MySheet01", "MySheet04");

    cout << "\nSheets in workbook:\n";
    for (const auto& name : wbk.worksheetNames()) cout << wbk.indexOfSheet(name) << " : " << name << "\n";

    cout << "\nDeleting sheet 'Sheet1'\n";
    wbk.deleteSheet("Sheet1");

    cout << "Moving sheet 'MySheet04' to index 1\n";
    wbk.worksheet("MySheet04").setIndex(1);

    cout << "Moving sheet 'MySheet03' to index 2\n";
    wbk.worksheet("MySheet03").setIndex(2);

    cout << "Moving sheet 'MySheet02' to index 3\n";
    wbk.worksheet("MySheet02").setIndex(3);

    cout << "Moving sheet 'MySheet01' to index 4\n";
    wbk.worksheet("MySheet01").setIndex(4);

    cout << "\nSheets in workbook:\n";
    for (const auto& name : wbk.worksheetNames()) cout << wbk.indexOfSheet(name) << " : " << name << "\n";

    wbk.sheet("MySheet01").setColor(XLColor(0, 0, 0));
    wbk.sheet("MySheet02").setColor(XLColor(255, 0, 0));
    wbk.sheet("MySheet03").setColor(XLColor(0, 255, 0));
    wbk.sheet("MySheet04").setColor(XLColor(0, 0, 255));

    doc.save();

    return 0;
}
```

### Unicode

```cpp
#include <OpenXLSX.hpp>
#include <iostream>

using namespace std;
using namespace OpenXLSX;

int main()
{
    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #03: Unicode\n";
    cout << "********************************************************************************\n";

    XLDocument doc1;
    doc1.create("./Demo03.xlsx");
    auto wks1 = doc1.workbook().worksheet("Sheet1");

    wks1.cell(XLCellReference("A1")).value() = "안녕하세요 세계!";
    wks1.cell(XLCellReference("A2")).value() = "你好，世界!";
    wks1.cell(XLCellReference("A3")).value() = "こんにちは世界";
    wks1.cell(XLCellReference("A4")).value() = "नमस्ते दुनिया!";
    wks1.cell(XLCellReference("A5")).value() = "Привет, мир!";
    wks1.cell(XLCellReference("A6")).value() = "Γειά σου Κόσμε!";

    doc1.save();
    doc1.saveAs("./スプレッドシート.xlsx");
    doc1.close();

    XLDocument doc2;
    doc2.open("./Demo03.xlsx");
    auto wks2 = doc2.workbook().worksheet("Sheet1");

    cout << "Cell A1 (Korean)  : " << wks2.cell(XLCellReference("A1")).value().get<std::string>() << endl;
    cout << "Cell A2 (Chinese) : " << wks2.cell(XLCellReference("A2")).value().get<std::string>() << endl;
    cout << "Cell A3 (Japanese): " << wks2.cell(XLCellReference("A3")).value().get<std::string>() << endl;
    cout << "Cell A4 (Hindi)   : " << wks2.cell(XLCellReference("A4")).value().get<std::string>() << endl;
    cout << "Cell A5 (Russian) : " << wks2.cell(XLCellReference("A5")).value().get<std::string>() << endl;
    cout << "Cell A6 (Greek)   : " << wks2.cell(XLCellReference("A6")).value().get<std::string>() << endl;


    cout << "\nNOTE: If you are using a Windows terminal, the above output will look like gibberish,\n"
            "because the Windows terminal does not support UTF-8 at the moment. To view to output,\n"
            "open the Demo03.xlsx file in Excel.\n\n";

    doc2.close();

    return 0;
}
```

### Number Formats

```cpp
#iclude <OpenXLSX.hpp>
#include <iostream>

using namespace std;
using namespace OpenXLSX;

int main()
{
    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #04: Number Formats\n";
    cout << "********************************************************************************\n";

    XLDocument doc1;
    doc1.create("./Demo04.xlsx");
    auto wks1 = doc1.workbook().worksheet("Sheet1");

    wks1.cell("A1").value() = 0.01;
    wks1.cell("B1").value() = 0.02;
    wks1.cell("C1").value() = 0.03;
    wks1.cell("A2").value() = 0.001;
    wks1.cell("B2").value() = 0.002;
    wks1.cell("C2").value() = 0.003;
    wks1.cell("A3").value() = 1e-4;
    wks1.cell("B3").value() = 2e-4;
    wks1.cell("C3").value() = 3e-4;

    wks1.cell("A4").value() = 1;
    wks1.cell("B4").value() = 2;
    wks1.cell("C4").value() = 3;
    wks1.cell("A5").value() = 845287496;
    wks1.cell("B5").value() = 175397487;
    wks1.cell("C5").value() = 973853975;
    wks1.cell("A6").value() = 2e10;
    wks1.cell("B6").value() = 3e11;
    wks1.cell("C6").value() = 4e12;

    doc1.save();
    doc1.close();

    XLDocument doc2;
    doc2.open("./Demo04.xlsx");
    auto wks2 = doc2.workbook().worksheet("Sheet1");

    auto PrintCell = [](const XLCell& cell) {
        cout << "Cell type is ";

        switch (cell.valueType()) {
            case XLValueType::Empty:à
                cout << "XLValueType::Empty";
                break;

            case XLValueType::Float:
                cout << "XLValueType::Float and the value is " << cell.value().get<double>() << endl;
                break;

            case XLValueType::Integer:
                cout << "XLValueType::Integer and the value is " << cell.value().get<int64_t>() << endl;
                break;

            default:
                cout << "Unknown";
        }
    };

    cout << "Cell A1: ";
    PrintCell(wks2.cell("A1"));

    cout << "Cell B1: ";
    PrintCell(wks2.cell("B1"));

    cout << "Cell C1: ";
    PrintCell(wks2.cell("C1"));

    cout << "Cell A2: ";
    PrintCell(wks2.cell("A2"));

    cout << "Cell B2: ";
    PrintCell(wks2.cell("B2"));

    cout << "Cell C2: ";
    PrintCell(wks2.cell("C2"));

    cout << "Cell A3: ";
    PrintCell(wks2.cell("A3"));

    cout << "Cell B3: ";
    PrintCell(wks2.cell("B3"));

    cout << "Cell C3: ";
    PrintCell(wks2.cell("C3"));

    cout << "Cell A4: ";
    PrintCell(wks2.cell("A4"));

    cout << "Cell B4: ";
    PrintCell(wks2.cell("B4"));

    cout << "Cell C4: ";
    PrintCell(wks2.cell("C4"));

    cout << "Cell A5: ";
    PrintCell(wks2.cell("A5"));

    cout << "Cell B5: ";
    PrintCell(wks2.cell("B5"));

    cout << "Cell C5: ";
    PrintCell(wks2.cell("C5"));

    cout << "Cell A6: ";
    PrintCell(wks2.cell("A6"));

    cout << "Cell B6: ";
    PrintCell(wks2.cell("B6"));

    cout << "Cell C6: ";
    PrintCell(wks2.cell("C6"));

    doc2.close();

    return 0;
}
```

### Ranges and Iterators

```cpp
#include <OpenXLSX.hpp>
#include <iostream>
#include <numeric>
#include <random>

using namespace std;
using namespace OpenXLSX;

int main()
{
    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #05: Ranges and Iterators\n";
    cout << "********************************************************************************\n";

    cout << "\nGenerating spreadsheet (1,048,576 rows x 8 columns) ..." << endl;
    XLDocument doc;
    doc.create("./Demo05.xlsx");
    auto wks = doc.workbook().worksheet("Sheet1");
    auto rng = wks.range(XLCellReference("A1"), XLCellReference(1048576, 8));

    std::random_device                 rand_dev;
    std::mt19937                       generator(rand_dev());
    std::uniform_int_distribution<int> distr(0, 99);

    for (auto& cell : rng) cell.value() = distr(generator);

    cout << "Saving spreadsheet (1,048,576 rows x 8 columns) ..." << endl;
    doc.save();
    doc.close();

    cout << "Re-opening spreadsheet (1,048,576 rows x 8 columns) ..." << endl;
    doc.open("./Demo05.xlsx");
    wks = doc.workbook().worksheet("Sheet1");
    rng = wks.range(XLCellReference("A1"), XLCellReference(1048576, 8));

    cout << "Reading data from spreadsheet (1,048,576 rows x 8 columns) ..." << endl;
    cout << "Cell count: " << std::distance(rng.begin(), rng.end()) << endl;
    cout << "Sum of cell values: "
         << accumulate(rng.begin(), rng.end(), 0, [](uint64_t a, XLCell& b) { return a + b.value().get<uint64_t>(); });

    doc.close();

    return 0;
}
```

## Changes

### New in version 0.3.0
HOLD

### New in version 0.2.0

The internal architecture of OpenXLSX has been significantly re-designed
since the previous version. The reason is that the library was turning
into a big ball of mud, and it was increasingly difficult to add
features and fix bugs. With the new architecture, it will (hopefully) be
easier to manage and to add new features.

Due to the re-design of the architecture, there are a few changes to the
public interface. These changes, however, are not significant, and it
should be easy to update:

* All internal objects are now handled as values rather than pointers,
  similar to the approach taken in the underlying PugiXML library. This
  means that when requesting a certain worksheet from a workbook, the
  resulting worksheet is not returned as a pointer, but as an object
  that supports both copying and moving.
* The distinction between interface objects and implementation objects
  are now gone, as it made it very difficult to manage changes. It was
  an attempt to implement the pimpl idiom, but it wasn't very effective.
  In the future, I may try to implement pimpl again, but only if it can
  be done in a simpler way.
* All member functions have been renamed to begin wtih a small letter
  (camelCase), i.e. the member function WorksheetCount() is renamed to
  worksheetCount(). This was done mostly for cosmetic reasons.

I realise that these changes may cause problems for some users. Because
of that, the previous version of OpenXLSX can be found in the "legacy"
branch of this repository. However, I strongly recommend that you
transition to the new version instead.

## Credits

- Thanks to [@zeux](https://github.com/zeux) for providing the awesome
  PugiXML library.
- Thanks to [@richgel999](https://github.com/richgel999) for providing
  the miniz library.
- Thanks to [@tectu](https://github.com/Tectu) for cleaning up my
  CMakeLists.txt files.


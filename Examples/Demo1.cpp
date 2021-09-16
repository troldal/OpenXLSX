#include <OpenXLSX.hpp>
#include <iostream>

using namespace std;
using namespace OpenXLSX;

int main()
{
    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #01: Basic Usage\n";
    cout << "********************************************************************************\n";

    // This example program illustrates basic usage of OpenXLSX, for example creation of a new workbook, and read/write
    // of cell values.


    // First, create a new document and access the sheet named 'Sheet1'.
    // New documents contain a single worksheet named 'Sheet1'
    XLDocument doc;
    doc.create("./Demo01.xlsx");
    auto wks = doc.workbook().worksheet("Sheet1");

    auto cell = wks.cell(1,1);
    cell.value() = 1;
    cout << cell.value().get<int>() << endl;

    // The individual cells can be accessed by using the .cell() method on the worksheet object.
    // The .cell() method can take the celladdress as a string, or alternatively take a XLCellReference
    // object. By using an XLCellReference object, the cells can be accessed by row/column coordinates.
    // The .cell() method returns an XLCell object.

    // The .value() method of an XLCell object can be used for both getting and setting the cell value.
    // Setting the value of a cell can be done bu using the assignment operator on the .value() method
    // as shown below. Alternatively, a .set() can be used. The cell values can be floating point numbers,
    // integers, strings, and booleans. It can also accept XLDateTime objects, but this requires special
    // handling (see later).
    wks.cell("A1").value() = 3.14159;
    wks.cell("B1").value() = 42;
    wks.cell("C1").value() = "  Hello OpenXLSX!  ";
    wks.cell("D1").value() = true;

    // As mentioned, the .value() method can also be used for getting tha value of a cell.
    // The .value() method returns a proxy object that cannot be copied or assigned, but
    // it can be implicitly converted to an XLCellValue object, as shown below.
    // Unfortunately, it is not possible to use the 'auto' keyword, so the XLCellValue
    // type has to be explicitly stated.
    XLCellValue A1 = wks.cell("A1").value();
    XLCellValue B1 = wks.cell("B1").value();
    XLCellValue C1 = wks.cell("C1").value();
    XLCellValue D1 = wks.cell("D1").value();

    // The cell value can be implicitly converted to a basic c++ type. However, if the type does not
    // match the type contained in the XLCellValue object (if, for example, floating point value is
    // assigned to a std::string), then an XLValueTypeError exception will be thrown.
    // To check which type is contained, use the .type() method, which will return a XLValueType enum
    // representing the type. As a convenience, the .typeAsString() method returns the type as a string,
    // which can be useful when printing to console.
    double vA1 = wks.cell("A1").value();
    int vB1 = wks.cell("B1").value();
    std::string vC1 = wks.cell("C1").value();
    bool vD1 = wks.cell("D1").value();

    cout << "Cell A1: (" << A1.typeAsString() << ") " << vA1 << endl;
    cout << "Cell B1: (" << B1.typeAsString() << ") " << vB1 << endl;
    cout << "Cell C1: (" << C1.typeAsString() << ") " << vC1 << endl;
    cout << "Cell D1: (" << D1.typeAsString() << ") " << vD1 << endl << endl;

    // Instead of using implicit (or explicit) conversion, the underlying value can also be retrieved
    // using the .get() method. This is a templated member function, which takes the desired type
    // as a template argument.
    cout << "Cell A1: (" << A1.typeAsString() << ") " << A1.get<double>() << endl;
    cout << "Cell B1: (" << B1.typeAsString() << ") " << B1.get<int64_t>() << endl;
    cout << "Cell C1: (" << C1.typeAsString() << ") " << C1.get<std::string>() << endl;
    cout << "Cell D1: (" << D1.typeAsString() << ") " << D1.get<bool>() << endl << endl;

    // XLCellValue objects can also be copied and assigned to other cells. This following line
    // will copy and assign the value of cell C1 to cell E1. Note tha only the value is copied;
    // other cell properties of the target cell remain unchanged.
    wks.cell("E1").value() = wks.cell(XLCellReference("C1")).value();
    XLCellValue E1 = wks.cell("E1").value();
    cout << "Cell E1: (" << E1.typeAsString() << ") " << E1.get<std::string_view>() << endl << endl;

    // Date/time values is a special case. In Excel, date/time values are essentially just a
    // 64-bit floating point value, that is rendered as a date/time string using special
    // formatting. When retrieving the cell value, it is just a floating point value,
    // and there is no way to identify it as a date/time value.
    // If, however, you know it to be a date time value, or if you want to assign a date/time
    // value to a cell, you can use the XLDateTime class, which falilitates conversion between
    // Excel date/time serial numbers, and the std::tm struct, that is used to store
    // date/time data. See https://en.cppreference.com/w/cpp/chrono/c/tm for more information.

    // An XLDateTime object can be created from a std::tm object:
    std::tm tm;
    tm.tm_year = 121;
    tm.tm_mon = 8;
    tm.tm_mday = 1;
    tm.tm_hour = 12;
    tm.tm_min = 0;
    tm.tm_sec = 0;
    XLDateTime dt (tm);

    // The std::tm object can be assigned to a cell value in the same way as shown previously.
    wks.cell("F1").value() = dt;

    // And as seen previously, an XLCellValue object can be retrieved. However, the object
    // will just contain a floating point value; there is no way to identify it as a date/time value.
    XLCellValue F1 = wks.cell("F1").value();
    cout << "Cell F1: (" << F1.typeAsString() << ") " << F1.get<double>() << endl;

    // If it is known to be a date/time value, the cell value can be converted to an XLDateTime object.
    auto result = F1.get<XLDateTime>();

    // The Excel date/time serial number can be retrieved using the .serial() method.
    cout << "Cell F1: (" << F1.typeAsString() << ") " << result.serial() << endl;

    // Using the .tm() method, the corresponding std::tm object can be retrieved.
    auto tmo = result.tm();
    cout << "Cell F1: (" << F1.typeAsString() << ") " << std::asctime(&tmo);

    doc.save();

    return 0;
}
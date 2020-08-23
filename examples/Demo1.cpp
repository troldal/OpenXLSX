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

    wks.cell(XLCellReference("A1")).value().set(3.14159);
    wks.cell(XLCellReference("B1")).value() = 42;
    wks.cell(XLCellReference("C1")).value() = "  Hello OpenXLSX!  ";
    wks.cell(XLCellReference("D1")).value() = true;
    wks.cell(XLCellReference("E1")).value() = wks.cell(XLCellReference("C1")).value();

    XLCellValue A1 = wks.cell(XLCellReference("A1")).value();
    XLCellValue B1 = wks.cell(XLCellReference("B1")).value();
    XLCellValue C1 = wks.cell(XLCellReference("C1")).value();
    XLCellValue D1 = wks.cell(XLCellReference("D1")).value();
    XLCellValue E1 = wks.cell(XLCellReference("E1")).value();

    cout << "Cell A1: (" << A1.typeAsString() << ") " << A1.get<double>() << endl;
    cout << "Cell B1: (" << B1.typeAsString() << ") " << B1.get<int64_t>() << endl;
    cout << "Cell C1: (" << C1.typeAsString() << ") " << C1.get<std::string>() << endl;
    cout << "Cell D1: (" << D1.typeAsString() << ") " << D1.get<bool>() << endl;
    cout << "Cell E1: (" << E1.typeAsString() << ") " << E1.get<std::string_view>() << endl;

    doc.save();

    return 0;
}
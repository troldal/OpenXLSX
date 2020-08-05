#include <OpenXLSX.hpp>
#include <iomanip>
#include <iostream>

using namespace std;
using namespace OpenXLSX;

int main()
{
    XLDocument doc;
    doc.create("./MyTest.xlsx");
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.cell(XLCellReference("A1")).value() = 3.14159;
    wks.cell(XLCellReference("B1")).value() = 42;
    wks.cell(XLCellReference("C1")).value() = string("  Hello OpenXLSX!  ");
    wks.cell(XLCellReference("D1")).value() = true;
    wks.cell(XLCellReference("E1")).value() = wks.cell(XLCellReference("C1")).value();

    auto A1 = wks.cell(XLCellReference("A1")).value().get<double>();
    auto B1 = wks.cell(XLCellReference("B1")).value().get<unsigned int>();
    auto C1 = wks.cell(XLCellReference("C1")).value().get<std::string>();
    auto D1 = wks.cell(XLCellReference("D1")).value().get<bool>();
    auto E1 = wks.cell(XLCellReference("E1")).value().get<std::string>();

    auto val = wks.cell(XLCellReference("E1")).value();
    cout << val.get<std::string>() << endl;

    cout << "Cell A1: " << A1 << endl;
    cout << "Cell B1: " << B1 << endl;
    cout << "Cell C1: " << C1 << endl;
    cout << "Cell D1: " << D1 << endl;
    cout << "Cell E1: " << E1 << endl;

    doc.save();

    return 0;
}
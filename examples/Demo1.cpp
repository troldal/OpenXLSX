#include <OpenXLSX.hpp>
#include <iomanip>
#include <iostream>

using namespace std;
using namespace OpenXLSX;

int main()
{
    XLDocument doc;
    doc.create("./MyTest.xlsx");
    auto wks = doc.workbook().Worksheet("Sheet1");

    wks.Cell(XLCellReference("A1")).Value() = 3.14159;
    wks.Cell(XLCellReference("B1")).Value() = 42;
    wks.Cell(XLCellReference("C1")).Value() = string("  Hello OpenXLSX!  ");
    wks.Cell(XLCellReference("D1")).Value() = true;
    wks.Cell(XLCellReference("E1")).Value() = wks.Cell(XLCellReference("C1")).Value();

    auto A1 = wks.Cell(XLCellReference("A1")).Value().Get<double>();
    auto B1 = wks.Cell(XLCellReference("B1")).Value().Get<unsigned int>();
    auto C1 = wks.Cell(XLCellReference("C1")).Value().Get<std::string>();
    auto D1 = wks.Cell(XLCellReference("D1")).Value().Get<bool>();
    auto E1 = wks.Cell(XLCellReference("E1")).Value().Get<std::string>();

    auto val = wks.Cell(XLCellReference("E1")).Value();
    cout << val.Get<std::string>() << endl;

    cout << "Cell A1: " << A1 << endl;
    cout << "Cell B1: " << B1 << endl;
    cout << "Cell C1: " << C1 << endl;
    cout << "Cell D1: " << D1 << endl;
    cout << "Cell E1: " << E1 << endl;

    doc.save();

    return 0;
}
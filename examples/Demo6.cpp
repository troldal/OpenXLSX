#include <OpenXLSX.hpp>
#include <iostream>

using namespace std;
using namespace OpenXLSX;

int main()
{
    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #06: Row Handling\n";
    cout << "********************************************************************************\n";

    XLDocument doc;
    doc.create("./Demo06.xlsx");
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.cell(XLCellReference("A1")).value() = 3.14159;
    wks.cell(XLCellReference("B1")).value() = 42;
    wks.cell(XLCellReference("C1")).value() = "  Hello OpenXLSX!  ";
    wks.cell(XLCellReference("D1")).value() = true;
    wks.cell(XLCellReference("E1")).value() = wks.cell(XLCellReference("C1")).getValue();

    auto row = wks.row(2);
    cout << row.cellCount() << endl;
    row.setHeight(50);

    doc.save();

    return 0;
}
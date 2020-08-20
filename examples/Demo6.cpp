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

    wks.cell(XLCellReference("A1")) = 3.14159;
    wks.cell(XLCellReference("B1")) = 42;
    wks.cell(XLCellReference("C1")) = "  Hello OpenXLSX!  ";
    wks.cell(XLCellReference("D1")) = true;
    wks.cell(XLCellReference("E1")) = wks.cell(XLCellReference("C1")).value();

    auto row = wks.row(2);
    cout << row.cellCount() << endl;
    row.setHeight(50);

    XLCellValue val;
    cout << val.get<std::string>() << endl;

    XLCellValue val1(1234);
    cout << val1.get<int>() << endl;

    XLCellValue val2(3.14159);
    cout << val2.get<float>() << endl;

    XLCellValue val3(true);
    cout << val3.get<bool>() << endl;

    XLCellValue val4("Hello OpenXLSX!");
    cout << val4.get<std::string_view>() << endl;

    val = val4;
    cout << val.get<std::string>() << endl;

    auto val5 = val2;
    cout << val2.get<float>() << endl;

    doc.save();

    return 0;
}
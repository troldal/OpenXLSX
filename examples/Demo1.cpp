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

    std::tm tmo;
    tmo.tm_year = 121;
    tmo.tm_mon = 8;
    tmo.tm_mday = 1;
    tmo.tm_hour = 12;
    tmo.tm_min = 0;
    tmo.tm_sec = 0;
    XLDateTime dt (tmo);

    wks.cell("A1").value() = 3.14159;
    wks.cell("B1").value() = 42;
    wks.cell("C1").value() = "  Hello OpenXLSX!  ";
    wks.cell("D1").value() = true;
    wks.cell("E1").value() = dt;
    wks.cell("F1").value() = wks.cell(XLCellReference("C1")).value();

    XLCellValue A1 = wks.cell("A1").value();
    XLCellValue B1 = wks.cell("B1").value();
    XLCellValue C1 = wks.cell("C1").value();
    XLCellValue D1 = wks.cell("D1").value();
    XLCellValue E1 = wks.cell("E1").value();
    XLCellValue F1 = wks.cell("F1").value();

    tmo = E1.get<XLDateTime>().tm();
    cout << "Cell A1: (" << A1.typeAsString() << ") " << A1.get<double>() << endl;
    cout << "Cell B1: (" << B1.typeAsString() << ") " << B1.get<int64_t>() << endl;
    cout << "Cell C1: (" << C1.typeAsString() << ") " << C1.get<std::string>() << endl;
    cout << "Cell D1: (" << D1.typeAsString() << ") " << D1.get<bool>() << endl;
    cout << "Cell E1: (" << E1.typeAsString() << ") " << std::asctime(&tmo);
    cout << "Cell F1: (" << F1.typeAsString() << ") " << F1.get<std::string_view>() << endl << endl;

    doc.save();

    return 0;
}
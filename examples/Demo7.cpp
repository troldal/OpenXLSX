#include <iostream>
#include <iomanip>
#include <OpenXLSX.hpp>

using namespace std;
using namespace OpenXLSX;

int main()
{
    XLDocument doc;
    doc.open("./スプレッドシート.xlsx");
    auto wks = doc.workbook()->Worksheet("Sheet1");

    auto A1 = wks->Cell(XLCellReference("A1")).Value().AsString();
    auto A2 = wks->Cell(XLCellReference("A2")).Value().AsString();

    cout << "Cell A1: " << A1 << endl;
    cout << "Cell A2: " << A2 << endl;

    doc.close();

    return 0;
}
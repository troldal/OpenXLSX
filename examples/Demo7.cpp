#include <iostream>
#include <iomanip>
#include <OpenXLSX.hpp>

using namespace std;
using namespace OpenXLSX::Impl;

int main() {

    XLDocument doc;
    doc.OpenDocument("./スプレッドシート.xlsx");
    auto wks = doc.Workbook()->Worksheet("Sheet1");

    auto A1 = wks->Cell(XLCellReference("A1")).Value().AsString();
    auto A2 = wks->Cell(XLCellReference("A2")).Value().AsString();

    cout << "Cell A1: " << A1 << endl;
    cout << "Cell A2: " << A2 << endl;

    doc.CloseDocument();

    return 0;
}
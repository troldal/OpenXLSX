#include <iostream>
#include <iomanip>
#include <cmath>
#include <OpenXLSX/OpenXLSX.h>

using namespace std;
using namespace OpenXLSX;

int main() {

    XLDocument doc;
    doc.CreateDocument("./SizeTest.xlsx");
    auto wks = doc.Workbook().Worksheet("Sheet1");

    auto arange = wks.Range(XLCellReference("A1"), XLCellReference(1000000, 50));

    for (auto iter : arange) {
        iter.Value() = M_PI;
    }

    doc.SaveDocument();

    return 0;
}
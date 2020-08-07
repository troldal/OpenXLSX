#include <OpenXLSX.hpp>
#include <chrono>
#include <iomanip>
#include <iostream>

using namespace std;
using namespace OpenXLSX;

int main()
{
    XLDocument doc;
    doc.open("./Names.xlsx");
    auto wbk = doc.workbook();
    auto wks = wbk.worksheet("Long Sheet Name");

    wks.setName("Even Longer Sheet Name");
    wbk.worksheet("Sheet1").setName("NewName");
    wbk.worksheet("Sheet2").setName("Number Two");
    wbk.worksheet("Sheet3").setName("Number Three");
    wbk.MoveSheet("NewName", 4);
    wbk.deleteSheet("Number Three");

    doc.saveAs("Names2.xlsx");

    return 0;
}

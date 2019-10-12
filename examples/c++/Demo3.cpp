#include <iostream>
#include <chrono>
#include <iomanip>
#include <OpenXLSX/OpenXLSX.h>

using namespace std;
using namespace OpenXLSX;

int main() {

    XLDocument doc;
    doc.OpenDocument("./Names.xlsx");
    auto wbk = doc.Workbook();
    auto wks = wbk.Worksheet("Long Sheet Name");

    wks.SetName("Even Longer Sheet Name");
    wbk.Worksheet("Sheet1").SetName("NewName");
    wbk.Worksheet("Sheet2").SetName("Number Two");
    wbk.Worksheet("Sheet3").SetName("Number Three");
    wbk.MoveSheet("NewName", 4);
    wbk.DeleteSheet("Number Three");

    doc.SaveDocumentAs("Names2.xlsx");

    return 0;
}

#include <iostream>
#include <OpenXLSX.hpp>

using namespace std;
using namespace OpenXLSX;

int main()
{
    XLDocument doc;
    doc.create("./MyTest.xlsx");
    auto wbk = doc.workbook();

    wbk.AddWorksheet("MySheet01");     // Append new sheet
    wbk.AddWorksheet("MySheet02");     // Prepend new sheet
    wbk.AddWorksheet("MySheet03");     // Prepend new sheet
    wbk.AddWorksheet("MySheet04");     // Insert new sheet
    wbk.setSheetIndex("Sheet1", 2);    // Move Sheet1 to second place
    wbk.Worksheet("Sheet1").Cell(XLCellReference("A1")).Value() = "Hello OpenXLSX";
    wbk.DeleteSheet("MySheet01");
    wbk.cloneSheet("Sheet1", "Sheet1Clone");

    for (const auto& name : wbk.WorksheetNames()) cout << name << ": " << wbk.IndexOfSheet(name) << endl;

    cout << endl;

    wbk.Worksheet("Sheet1").SetName("BLAH");

    for (auto iter = 1; iter <= wbk.SheetCount(); ++iter) {
        cout << iter << ": " << wbk.Sheet(iter).Name() << endl;
    }

    doc.save();

    return 0;
}

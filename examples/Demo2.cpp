#include <iostream>
#include <OpenXLSX.hpp>

using namespace std;
using namespace OpenXLSX;

int main()
{
    XLDocument doc;
    doc.create("./MyTest.xlsx");
    auto wbk = doc.workbook();

    wbk.addWorksheet("MySheet01");     // Append new sheet
    wbk.addWorksheet("MySheet02");     // Prepend new sheet
    wbk.addWorksheet("MySheet03");     // Prepend new sheet
    wbk.addWorksheet("MySheet04");     // Insert new sheet
    wbk.setSheetIndex("Sheet1", 2);    // Move Sheet1 to second place
    wbk.worksheet("Sheet1").cell(XLCellReference("A1")).value() = "Hello OpenXLSX";
    wbk.deleteSheet("MySheet01");
    // wbk.cloneSheet("Sheet1", "Sheet1Clone");
    wbk.sheet("Sheet1").get<XLWorksheet>().clone("Sheet1Clone");

    for (const auto& name : wbk.worksheetNames()) cout << name << ": " << wbk.indexOfSheet(name) << endl;

    cout << endl;

    wbk.worksheet("Sheet1").setName("BLAH");

    for (auto iter = 1; iter <= wbk.sheetCount(); ++iter) {
        cout << iter << ": " << wbk.sheet(iter).name() << endl;
    }

    doc.save();

    return 0;
}

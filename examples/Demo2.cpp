#include <OpenXLSX.hpp>
#include <iostream>

using namespace std;
using namespace OpenXLSX;

int main()
{
    XLDocument doc;
    doc.create("./MyTest.xlsx");
    auto wbk = doc.workbook();

    wbk.addWorksheet("MySheet01");    // Append new sheet
    wbk.addWorksheet("MySheet02");    // Prepend new sheet
    wbk.addWorksheet("MySheet03");    // Prepend new sheet
    wbk.addWorksheet("MySheet04");    // Insert new sheet
    wbk.worksheet("Sheet1").setIndex(3);
    //    wbk.setSheetIndex("Sheet1", 3);    // Move Sheet1 to second place
    wbk.worksheet("Sheet1").cell(XLCellReference("A1")).value() = "Hello OpenXLSX";
    wbk.deleteSheet("MySheet01");
    // wbk.cloneSheet("Sheet1", "Sheet1Clone");
    wbk.sheet("Sheet1").get<XLWorksheet>().clone("Sheet1Clone");
    wbk.sheet("Sheet1Clone").setColor(XLColor(0, 253, 106, 2));

    for (const auto& name : wbk.worksheetNames()) cout << name << ": " << wbk.indexOfSheet(name) << endl;

    cout << endl;

    wbk.worksheet("Sheet1").setName("BLAH");

    for (uint32_t iter = 1; iter <= wbk.sheetCount(); ++iter) {
        cout << iter << ": " << wbk.sheet(iter).name() << endl;
    }

    doc.save();

    return 0;
}

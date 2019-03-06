#include <iostream>
#include <chrono>
#include <iomanip>
#include <OpenXLSX/OpenXLSX.h>

using namespace std;
using namespace OpenXLSX;

int main() {

    XLDocument doc;
    doc.CreateDocument("./MyTest.xlsx");
    auto wbk = doc.Workbook();

    wbk.AddWorksheet("MySheet01");    // Append new sheet
    wbk.AddWorksheet("MySheet02", 1); // Prepend new sheet
    wbk.AddWorksheet("MySheet03", 1); // Prepend new sheet
    wbk.AddWorksheet("MySheet04", 2); // Insert new sheet
    wbk.MoveSheet("Sheet1", 2);       // Move Sheet1 to second place
    wbk.DeleteSheet("MySheet01");

    wbk.Worksheet("Sheet1").Row(1);

    for (const auto& name : wbk.WorksheetNames())
        cout << name << ": " << wbk.IndexOfSheet(name) << endl;

    cout << endl;

    for (auto iter = 1; iter <= wbk.SheetCount(); ++iter)
        cout << iter << ": " << wbk.Sheet(iter).Name() << endl;

    doc.SaveDocument();

    return 0;
}

#include <iostream>
#include <iomanip>
#include <OpenXLSX/OpenXLSX.h>

using namespace std;
using namespace OpenXLSX;

/*
 * TODO: Sheet iterator
 * TODO: Handling of named ranges
 * TODO: Column/Row iterators
 * TODO: correct copy/move operations for all classes
 * TODO: Find a way to handle overwriting of shared formulas.
 * TODO: Handling of Cell formatting
 * TODO: Handle chartsheets
 * TODO: Update formulas when changing Sheet Name.
 * TODO: Get vector for a Row or Column.
 * TODO: Conditional formatting
 */

int main() {

    XLDocument doc;
    doc.CreateDocument("./MyTest.xlsx");
    auto wks = doc.Workbook().Worksheet("Sheet1");

    wks.Cell("A1").Value() = 3.14159;
    wks.Cell("B1").Value() = 42;
    wks.Cell("C1").Value() = "Hello OpenXLSX!";
    wks.Cell("D1").Value() = true;
    wks.Cell("E1").Value() = wks.Cell("C1").Value();

    auto A1 = wks.Cell("A1").Value().Get<double>();
    auto B1 = wks.Cell("B1").Value().Get<unsigned int>();
    auto C1 = wks.Cell("C1").Value().Get<std::string>();
    auto D1 = wks.Cell("D1").Value().Get<bool>();
    auto E1 = wks.Cell("E1").Value().Get<std::string>();

    auto val = wks.Cell("E1").Value();
    cout << val.Get<std::string>() << endl;

    cout << "Cell A1: " << A1 << endl;
    cout << "Cell B1: " << B1 << endl;
    cout << "Cell C1: " << C1 << endl;
    cout << "Cell D1: " << D1 << endl;
    cout << "Cell E1: " << E1 << endl;

    doc.SaveDocument();

    return 0;
}

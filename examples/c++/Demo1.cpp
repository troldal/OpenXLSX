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

/*int main() {

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
}*/

#include <iostream>
#include <iomanip>
#include <OpenXLSX/OpenXLSX.h>

using namespace std;
using namespace OpenXLSX;

int main(int argc, char** argv)
{
    XLDocument doc;

    if (argc < 2)
    {
        cout << "Usage: " << argv[0] << " <file.xlsx>" << endl;
        return 1;
    }

    doc.OpenDocument(argv[1]);
    auto wks = doc.Workbook().Worksheet("Tabelle1");

    auto tA2 = wks.Cell("A2").ValueType();

    if (tA2 == OpenXLSX::XLValueType::Error)
        cout << "Format error at cell A2" << endl;

    auto tC2 = wks.Cell("C2").ValueType();

    if (tC2 == OpenXLSX::XLValueType::Error)
        cout << "Format error at cell C2" << endl;

    auto tD2 = wks.Cell("D2").ValueType();

    if (tD2 == OpenXLSX::XLValueType::Error)
        cout << "Format error at cell D2" << endl;

    auto A1 = wks.Cell("A1").Value().Get<std::string>();
    auto A2 = wks.Cell("A2").Value().Get<unsigned int>();
    auto B1 = wks.Cell("B1").Value().Get<std::string>();
    auto B2 = wks.Cell("B2").Value().Get<std::string>();

    cout << "Cell A1: " << A1 << endl;
    cout << "Cell A2: " << A2 << endl;
    cout << "Cell B1: " << B1 << endl;
    cout << "Cell B2: " << B2 << endl;

    return 0;
}

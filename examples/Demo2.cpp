#include <OpenXLSX.hpp>
#include <iostream>

using namespace std;
using namespace OpenXLSX;

int main()
{
    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #02: Sheet Handling\n";
    cout << "********************************************************************************\n";

    XLDocument doc;
    doc.create("./Demo02.xlsx");
    auto wbk = doc.workbook();

    cout << "\nSheets in workbook:\n";
    for (const auto& name : wbk.worksheetNames()) cout << wbk.indexOfSheet(name) << " : " << name << "\n";

    cout << "\nAdding new sheet 'MySheet01'\n";
    wbk.addWorksheet("MySheet01");

    cout << "Adding new sheet 'MySheet02'\n";
    wbk.addWorksheet("MySheet02");

    cout << "Cloning sheet 'Sheet1' to new sheet 'MySheet03'\n";
    wbk.sheet("Sheet1").get<XLWorksheet>().clone("MySheet03");

    cout << "Cloning sheet 'MySheet01' to new sheet 'MySheet04'\n";
    wbk.cloneSheet("MySheet01", "MySheet04");

    cout << "\nSheets in workbook:\n";
    for (const auto& name : wbk.worksheetNames()) cout << wbk.indexOfSheet(name) << " : " << name << "\n";

    cout << "\nDeleting sheet 'Sheet1'\n";
    wbk.deleteSheet("Sheet1");

    cout << "Moving sheet 'MySheet04' to index 1\n";
    wbk.worksheet("MySheet04").setIndex(1);

    cout << "Moving sheet 'MySheet03' to index 2\n";
    wbk.worksheet("MySheet03").setIndex(2);

    cout << "Moving sheet 'MySheet02' to index 3\n";
    wbk.worksheet("MySheet02").setIndex(3);

    cout << "Moving sheet 'MySheet01' to index 4\n";
    wbk.worksheet("MySheet01").setIndex(4);

    cout << "\nSheets in workbook:\n";
    for (const auto& name : wbk.worksheetNames()) cout << wbk.indexOfSheet(name) << " : " << name << "\n";

    wbk.sheet("MySheet01").setColor(XLColor(0, 0, 0));
    wbk.sheet("MySheet02").setColor(XLColor(255, 0, 0));
    wbk.sheet("MySheet03").setColor(XLColor(0, 255, 0));
    wbk.sheet("MySheet04").setColor(XLColor(0, 0, 255));

    doc.save();

    return 0;
}

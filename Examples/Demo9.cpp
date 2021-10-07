#include <OpenXLSX.hpp>
#include <deque>
#include <iostream>
#include <list>
#include <random>

using namespace std;
using namespace OpenXLSX;

int main()
{

    XLDocument doc;
    doc.open("./read-tdne.xlsx");
    auto wks = doc.workbook().worksheet("TEST");

    for (auto row : wks.rows()) {
        for (auto cell : row.cells()) {
            cout << cell.value().get<int>() << endl;
        }
    }

    doc.workbook().addWorksheet("Other");
    auto other = doc.workbook().worksheet("Other");
    other.cell("A1").value() = "Blah!";

    doc.saveAs("Demo9.xlsx");
    doc.close();

    return 0;
}

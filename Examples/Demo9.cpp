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
    doc.open("./test.xlsx");
    auto wks = doc.workbook().worksheet("sheet");

    cout << wks.cell("A1").value() << endl;
    cout << wks.cell("B1").value() << endl;

    doc.close();

    return 0;
}

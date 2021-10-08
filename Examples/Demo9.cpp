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

    if (wks.cell("A1").value().type() == XLValueType::Integer)
        cout << wks.cell("A1").value() << endl;

    if (wks.cell("A2").value().type() == XLValueType::Integer)
        cout << wks.cell("A2").value() << endl;

    doc.close();

    return 0;
}

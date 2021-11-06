#include <OpenXLSX.hpp>
#include <iostream>

using namespace std;
using namespace OpenXLSX;

int main()
{

    XLDocument doc;
    doc.open("test_case.xlsx");
    auto wks = doc.workbook().worksheet("equipment");

    return 0;
}
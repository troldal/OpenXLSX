#include <OpenXLSX.hpp>
#include <iostream>

using namespace std;
using namespace OpenXLSX;

int main()
{
    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #06: Reset Calculation Chain\n";
    cout << "********************************************************************************\n";

    XLDocument doc;
    doc.open("/Users/troldal/Desktop/test.xlsx");
    doc.resetCalcChain();
    doc.saveAs("/Users/troldal/Desktop/test2.xlsx");
    doc.close();

    return 0;
}
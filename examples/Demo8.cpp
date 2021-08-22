  
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
    doc.open("./Demo01.xlsx");
    doc.resetCalcChain();
    doc.saveAs("./Demo01_CalcReset.xlsx");
    doc.close();

    return 0;
}
#include <iostream>
#include <OpenXLSX.hpp>

using namespace std;
using namespace OpenXLSX;

int main()
{
    XLDocument doc1;
    doc1.create("./StringTest.xlsx");
    auto wks1 = doc1.workbook().worksheet("Sheet1");

    wks1.cell("A1").value() = "Hello OpenXLSX!";
    wks1.cell("B1").value() = " Hello OpenXLSX! ";
    wks1.cell("C1").value() = "  Hello OpenXLSX!  ";
    wks1.cell("A2").value() = "";
    wks1.cell("B2").value() = " ";
    wks1.cell("C2").value() = "  ";

    doc1.save();
    doc1.close();

    XLDocument doc2;
    doc2.open("./StringTest.xlsx");
    auto wks2 = doc2.workbook().worksheet("Sheet1");

    cout << "Cell A1: " << wks2.cell("A1").value().get<std::string>().length() << endl;
    cout << "Cell B1: " << wks2.cell("B1").value().get<std::string>().length() << endl;
    cout << "Cell C1: " << wks2.cell("C1").value().get<std::string>().length() << endl;
    cout << "Cell A2: " << wks2.cell("A2").value().get<std::string>().length() << endl;
    cout << "Cell B2: " << wks2.cell("B2").value().get<std::string>().length() << endl;
    cout << "Cell C2: " << wks2.cell("C2").value().get<std::string>().length() << endl;

    doc2.close();

    return 0;
}
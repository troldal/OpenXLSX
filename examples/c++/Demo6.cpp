#include <iostream>
#include <OpenXLSX/OpenXLSX.h>

using namespace std;
using namespace OpenXLSX;

int main() {

    XLDocument doc1;
    doc1.CreateDocument("./StringTest.xlsx");
    auto wks1 = doc1.Workbook().Worksheet("Sheet1");

    wks1.Cell("A1").Value() = "Hello OpenXLSX!";
    wks1.Cell("B1").Value() = " Hello OpenXLSX! ";
    wks1.Cell("C1").Value() = "  Hello OpenXLSX!  ";
    wks1.Cell("A2").Value() = "";
    wks1.Cell("B2").Value() = " ";
    wks1.Cell("C2").Value() = "  ";

    doc1.SaveDocument();
    doc1.CloseDocument();

    XLDocument doc2;
    doc2.OpenDocument("./StringTest.xlsx");
    auto wks2 = doc2.Workbook().Worksheet("Sheet1");

    cout << "Cell A1: " << wks2.Cell("A1").Value().Get<std::string>().length() << endl;
    cout << "Cell B1: " << wks2.Cell("B1").Value().Get<std::string>().length() << endl;
    cout << "Cell C1: " << wks2.Cell("C1").Value().Get<std::string>().length() << endl;
    cout << "Cell A2: " << wks2.Cell("A2").Value().Get<std::string>().length() << endl;
    cout << "Cell B2: " << wks2.Cell("B2").Value().Get<std::string>().length() << endl;
    cout << "Cell C2: " << wks2.Cell("C2").Value().Get<std::string>().length() << endl;

    doc2.CloseDocument();

    return 0;
}
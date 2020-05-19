#include <iostream>
#include <iomanip>
#include <OpenXLSX/OpenXLSX.h>

using namespace std;
using namespace OpenXLSX;

int main() {

    XLDocument document;
    document.OpenDocument(R"(/Users/Troldal/Downloads/ress.xlsx)");
    XLWorkbook xlWorkbook = document.Workbook();
    XLWorksheet worksheet = xlWorkbook.Worksheet("Sheet1");

    string a1 = worksheet.Cell("A1").Value().AsString();
    string b1 = worksheet.Cell("B1").Value().AsString();
    string c1 = worksheet.Cell("C1").Value().AsString();
    cout << a1.size() << " " << b1.size() <<" "<< c1.size() << endl;  //0 4 3
    cout << a1 << " " << b1 << c1 << endl;  // abcd 분

    return 0;
}